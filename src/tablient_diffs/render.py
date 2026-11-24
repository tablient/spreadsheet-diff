"""Render FileDiffs as one Markdown PR comment."""
from __future__ import annotations

from .diff import FileDiff

MARKER = "<!-- xlsxdiff -->"  # the old name on purpose: comments posted before the rename are still found and updated
MAX_COMMENT = 60_000  # GitHub's limit is 65,536 characters
PER_SHEET_LINES = 200


def _fence(text: str) -> str:
    fence = "```"
    while fence in text:
        fence += "`"
    return f"{fence}diff\n{text}\n{fence}"


def _sheet_block(s, budget: int) -> str:
    lines = s.diff.splitlines()
    truncated = len(lines) > PER_SHEET_LINES
    body = "\n".join(lines[:PER_SHEET_LINES])
    if len(body) > budget:
        body, truncated = body[:max(budget, 0)], True
    note = "\n_Diff truncated; see the file in the Files tab._\n" if truncated else ""
    label = f"{s.name}: {s.status}, +{s.added} / -{s.removed} lines"
    return f"<details><summary>{label}</summary>\n\n{_fence(body)}\n{note}</details>\n"


def render(files: list[FileDiff], summary: str | None = None) -> str:
    parts = [MARKER, "### Spreadsheet changes", ""]
    changed = [f for f in files if f.changed]
    if not changed:
        parts.append("No cell, formula or format changes in the spreadsheets touched by this PR.")
        return "\n".join(parts) + "\n"
    if summary:
        parts += [summary.strip()[:MAX_COMMENT // 2], ""]
    footer = ["", "<sub>Lines starting with `+` are after the change, `-` before. Format: open Cellprose text.</sub>"]
    # Everything counts against the limit; keep room for the omission notes at the end.
    room = MAX_COMMENT - len("\n".join(parts + footer)) - 400

    def add(lines: list[str]) -> bool:
        nonlocal room
        size = len("\n".join(lines)) + 1
        if size > room:
            return False
        parts.extend(lines)
        room -= size
        return True

    omitted = 0
    for i, f in enumerate(changed):
        head = [f"#### `{f.path}` ({f.status})"]
        if f.status == "error":
            head += [f"Could not compare this workbook: {f.error}.", ""]
        else:
            if f.sheets:
                head += [f"{len(f.sheets)} sheet(s) changed: " + ", ".join(f"`{s.name}`" for s in f.sheets), ""]
            if f.parts:
                head += ["Charts, pivot tables and other parts:", *[f"- {p}" for p in f.parts], ""]
        if not add(head):
            omitted = len(changed) - i
            break
        full = False
        for s in f.sheets:
            if not add([_sheet_block(s, room // 2)]):
                parts.append(f"_More changes omitted (comment size limit): `{s.name}` and later sheets._")
                full = True
                break
        if full:
            omitted = len(changed) - i - 1
            break
    if omitted:
        parts.append(f"_{omitted} more file(s) omitted (comment size limit); see the Files tab._")
    parts += footer
    return "\n".join(parts) + "\n"
