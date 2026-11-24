"""Turn two versions of a workbook into a per-sheet text diff."""
from __future__ import annotations

import bisect
import difflib
import re
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

from .parts import diff_parts

SHEET_RE = re.compile(r"^# Sheet: (.*)$")
# Keys for the parts of the text that are not sheets. Excel forbids [ and ] in sheet names, so no real sheet collides.
FRONT_MATTER, NAMES = "[front matter]", "[defined names]"


@dataclass
class SheetChange:
    name: str
    status: str  # "added" | "removed" | "changed"
    added: int = 0
    removed: int = 0
    diff: str = ""


@dataclass
class FileDiff:
    path: str
    status: str  # "added" | "removed" | "modified" | "error"
    sheets: list[SheetChange] = field(default_factory=list)
    error: str = ""
    parts: list[str] = field(default_factory=list)  # charts, pivots, macros, images: see parts.py

    @property
    def changed(self) -> bool:
        return bool(self.sheets or self.parts) or self.status in ("added", "removed", "error")


def to_text(data: bytes) -> str:
    """Cellprose text of a workbook given as bytes (no recalculation, no source name: both would add noise)."""
    from cellprose.gitflow import textconv
    with tempfile.TemporaryDirectory() as d:
        p = Path(d) / "book.xlsx"
        p.write_bytes(data)
        return textconv(p)


def split_sheets(text: str) -> dict[str, list[str]]:
    """Split Cellprose text into {sheet name: lines}. Lines before the first sheet and the `# Names` section
    are kept under their own keys so a change to either still shows up."""
    out: dict[str, list[str]] = {FRONT_MATTER: []}
    cur = FRONT_MATTER
    for line in text.splitlines():
        m = SHEET_RE.match(line)
        if m:
            cur = m.group(1)
            out.setdefault(cur, [])
        elif line.startswith("# Names"):
            cur = NAMES
            out.setdefault(cur, [])
        out[cur].append(line)
    return {k: v for k, v in out.items() if any(s.strip() for s in v)}


ROW_RE = re.compile(r"^\|?\d+\|")


def canonical(lines: list[str]) -> list[str]:
    """Reorder a sheet's lines so the diff does not depend on how the writer split it into regions.

    Region headings, separator rows and blank lines are dropped; table rows are grouped under their column header (a
    repeated identical header, as after an empty row splits one region in two, joins the first group); fenced blocks
    such as `formulas` and `formats` are collected by name after the tables.
    """
    head: list[str] = []
    runs: dict[str, list[str]] = {}
    fences: dict[str, list[str]] = {}
    header = None
    fence = None
    for line in lines:
        if fence is not None:
            if line.startswith("```"):
                fence = None
            elif line.strip():
                fences[fence].append(line)
        elif line.startswith("```"):
            fence = line[3:].strip() or "(block)"
            fences.setdefault(fence, [])
        elif line.startswith("## ") or line.startswith(("|-", "-|")) or not line.strip():
            continue
        elif line.startswith(("|#|", "#|")):
            header = line
            runs.setdefault(header, [])
        elif ROW_RE.match(line) and header is not None:
            runs[header].append(line)
        else:
            head.append(line)
    out = list(head)
    for h, rows in runs.items():
        out += ["", h, *rows]
    for name, body in fences.items():
        out += ["", f"```{name}", *body, "```"]
    return out


def sheet_key(name: str) -> str:
    """Sheet name as a formula reference spells it, without quoting: `'My Sheet'` -> `My Sheet`."""
    if len(name) > 1 and name[0] == name[-1] == "'":
        name = name[1:-1].replace("''", "'")
    return name


def _key(line: str) -> str:
    """Comparison key: a table row without its row number, so inserting a row does not mark every row below it."""
    return ROW_RE.sub("|#|", line, count=1)


_REF_RE = re.compile(r"(?<![A-Za-z0-9_.])((?:'[^']+'|[A-Za-z0-9_.]+)!)?(\$?[A-Z]{1,3})(\$?)(\d+)(?::(\$?[A-Z]{1,3})(\$?)(\d+))?(?![A-Za-z0-9_(])")
_ROW_NO = re.compile(r"^\|?(\d+)\|")


def _row_map(old: list[str], new: list[str]):
    """Old row number -> new row number, learned from table rows that matched by content on both sides."""
    sm = difflib.SequenceMatcher(None, [_key(l) for l in old], [_key(l) for l in new], autojunk=False)
    pairs: list[tuple[int, int]] = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for a, b in zip(old[i1:i2], new[j1:j2]):
                ma, mb = _ROW_NO.match(a), _ROW_NO.match(b)
                if ma and mb:
                    pairs.append((int(ma[1]), int(mb[1])))
    anchors: list[tuple[int, int]] = []
    for o, n in pairs:                       # keep a monotone chain so a stray match cannot bend the map
        if not anchors or (o > anchors[-1][0] and n > anchors[-1][1]):
            anchors.append((o, n))
    olds = [o for o, _ in anchors]

    def f(r: int) -> int:
        i = bisect.bisect_right(olds, r) - 1
        return r if i < 0 else anchors[i][1] + (r - anchors[i][0])
    return f


def _shift_refs(line: str, f, maps: dict | None = None) -> str:
    """Rewrite A1 references and ranges in a non-table line through row maps: references without a sheet use `f` (the
    sheet this line is on), references to another sheet use that sheet's own map from `maps`."""
    def sub(m):
        g = f
        if m[1]:
            g = (maps or {}).get(sheet_key(m[1][:-1]))
            if g is None:                      # unknown sheet, or one that was not changed
                return m[0]
        a = f"{m[1] or ''}{m[2]}{m[3]}{g(int(m[4]))}"
        return a + (f":{m[5]}{m[6]}{g(int(m[7]))}" if m[5] else "")
    parts = line.split('"')                    # leave text inside string literals alone
    return '"'.join(_REF_RE.sub(sub, p) if i % 2 == 0 else p for i, p in enumerate(parts))


def unified(old: list[str], new: list[str], context: int = 2, maps: dict | None = None) -> list[str]:
    """Unified-diff body (no file header) of two line lists. Table rows are matched by content, not row number, and
    row references in region headings, formulas and formats are compared after the same row shift, so inserting a row
    shows as that row (and the heading it grows) rather than as every formula below it."""
    f = _row_map(old, new)
    okeys = [_key(l) if ROW_RE.match(l) else _shift_refs(l, f, maps) for l in old]
    nkeys = [_key(l) for l in new]
    sm = difflib.SequenceMatcher(None, okeys, nkeys, autojunk=False)
    out: list[str] = []
    for group in sm.get_grouped_opcodes(context):
        out.append(f"@@ -{group[0][1] + 1},{group[-1][2] - group[0][1]} +{group[0][3] + 1},{group[-1][4] - group[0][3]} @@")
        for tag, i1, i2, j1, j2 in group:
            if tag == "equal":
                out += [" " + l for l in new[j1:j2]]
                continue
            out += ["-" + l for l in old[i1:i2]]
            out += ["+" + l for l in new[j1:j2]]
    return out


def diff_texts(old: str | None, new: str | None, context: int = 2) -> list[SheetChange]:
    a = split_sheets(old) if old is not None else {}
    b = split_sheets(new) if new is not None else {}
    for d in (a, b):
        for k in d:
            if k not in (FRONT_MATTER, NAMES):
                d[k] = canonical(d[k])
    # Row maps of every sheet present in both versions, so a formula on one sheet that points at rows of another
    # is compared after that other sheet's rows moved.
    maps = {sheet_key(k): _row_map(a[k], b[k]) for k in a if k in b and k not in (FRONT_MATTER, NAMES)}
    changes = []
    for name in list(b) + [n for n in a if n not in b]:
        old_lines, new_lines = a.get(name), b.get(name)
        if old_lines == new_lines:
            continue
        status = "added" if old_lines is None else "removed" if new_lines is None else "changed"
        body = unified(old_lines or [], new_lines or [], context, maps)
        if not body:                       # differs only by row shifts that the map already explains
            continue
        changes.append(SheetChange(
            name=name, status=status,
            added=sum(1 for l in body if l.startswith("+")),
            removed=sum(1 for l in body if l.startswith("-")),
            diff="\n".join(body)))
    return changes


def _explain(side: str, e: Exception) -> str:
    """One line saying which version failed and why, instead of a bare parser exception."""
    kind = type(e).__name__
    if kind in ("BadZipFile", "InvalidFileException"):
        why = "it is not a valid .xlsx file (damaged, or another format saved under this name)"
    elif "XML" in kind or "ParseError" in kind or "syntax" in str(e).lower():
        why = "its XML could not be parsed (some tools write files Excel accepts but this reader does not)"
    elif "password" in str(e).lower() or "encrypt" in str(e).lower():
        why = "it is password-protected"
    else:
        why = f"{kind}: {e}"
    return f"the {side} version could not be read: {why}"


def diff_workbooks(path: str, old: bytes | None, new: bytes | None) -> FileDiff:
    if old is None and new is None:
        return FileDiff(path, "error", error="no content on either side")
    if old == new:
        return FileDiff(path, "modified")
    texts = []
    for side, data in (("before", old), ("after", new)):
        try:
            texts.append(to_text(data) if data is not None else None)
        except Exception as e:  # corrupt or password-protected files must not fail the whole run
            return FileDiff(path, "error", error=_explain(side, e))
    old_t, new_t = texts
    status = "added" if old is None else "removed" if new is None else "modified"
    try:
        parts = diff_parts(old, new)
    except Exception:  # an unreadable zip would already have failed above; never fail the run over extras
        parts = []
    return FileDiff(path, status, diff_texts(old_t, new_t), parts=parts)
