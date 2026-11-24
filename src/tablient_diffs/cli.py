"""`tablient-diffs` command: compare workbooks locally, or run as the PR check inside the GitHub Action."""
from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path

from . import github
from .diff import FileDiff, diff_workbooks
from .render import render
from .summarize import summarize

STATUS = {"A": "added", "D": "removed", "M": "modified"}


def _event_pr(path: str | None) -> int | None:
    if not path or not Path(path).exists():
        return None
    ev = json.loads(Path(path).read_text())
    return (ev.get("pull_request") or {}).get("number") or ev.get("number")


def run_pr(a) -> int:
    files: list[FileDiff] = []
    # Old contents come from the merge base, the same commit `base...head` diffs against, so changes
    # made on the base branch after the PR branched off do not show up as PR changes.
    old_ref = a.base_sha or github.merge_base(a.base, a.head) or a.base
    for status, path in github.changed_workbooks(a.base, a.head):
        old = None if status == "A" else github.blob(old_ref, path)
        new = None if status == "D" else github.blob(a.head, path)
        files.append(diff_workbooks(path, old, new))
    summary = None
    key = os.environ.get("ANTHROPIC_API_KEY")
    if a.summarize and key:
        summary = summarize([f for f in files if f.changed], key, a.model)
    body = render(files, summary)
    if not files:
        print("No spreadsheet files changed.")
        return 0
    pr, token, repo = _event_pr(os.environ.get("GITHUB_EVENT_PATH")), os.environ.get("GITHUB_TOKEN"), os.environ.get("GITHUB_REPOSITORY")
    if a.dry_run or not (pr and token and repo):
        print(body)
        return 0
    print(f"Comment {github.upsert_comment(repo, pr, token, body)} on #{pr}")
    return 0


def run_files(a) -> int:
    old = Path(a.old).read_bytes() if a.old != "-" else None
    new = Path(a.new).read_bytes() if a.new != "-" else None
    print(render([diff_workbooks(a.new if new is not None else a.old, old, new)]))
    return 0


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="tablient-diffs", description="Readable diffs of .xlsx changes.")
    sub = p.add_subparsers(dest="cmd", required=True)
    f = sub.add_parser("files", help="diff two workbook files (use - for a missing side)")
    f.add_argument("old")
    f.add_argument("new")
    r = sub.add_parser("pr", help="diff workbooks changed between two git refs and comment on the PR")
    r.add_argument("--base", required=True, help="base ref or sha")
    r.add_argument("--base-sha", help="sha to read the old file from (default: merge base of --base and --head)")
    r.add_argument("--head", default="HEAD")
    r.add_argument("--summarize", action="store_true", help="add an LLM summary (needs ANTHROPIC_API_KEY)")
    r.add_argument("--model", default="claude-haiku-5-5")
    r.add_argument("--dry-run", action="store_true", help="print the comment instead of posting it")
    a = p.parse_args(argv)
    return run_files(a) if a.cmd == "files" else run_pr(a)


if __name__ == "__main__":
    sys.exit(main())
