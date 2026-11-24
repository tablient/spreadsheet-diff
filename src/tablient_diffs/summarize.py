"""Optional plain-English summary of a diff from an LLM. Off unless an API key is given."""
from __future__ import annotations

import json
import urllib.request

from .diff import FileDiff

PROMPT = (
    "You review changes to spreadsheet models in a pull request. Below are line diffs of the workbooks in "
    "Cellprose text (rows are `|row|cells|`, `formulas` blocks hold fill ranges such as `B4:E4 = B2*B3`). "
    "Write 2-5 short bullet points a reviewer can read in ten seconds: which assumptions or inputs changed, "
    "which formulas changed, and any typed-in value that replaced a formula. Be specific with cell addresses. "
    "Do not speculate beyond the diff.\n\n"
)


def build_input(files: list[FileDiff], limit: int = 30_000) -> str:
    chunks = []
    for f in files:
        for s in f.sheets:
            chunks.append(f"## {f.path} / {s.name} ({s.status})\n{s.diff}")
    return "\n\n".join(chunks)[:limit]


def summarize(files: list[FileDiff], api_key: str, model: str, base_url: str = "https://api.anthropic.com") -> str | None:
    text = build_input(files)
    if not text.strip():
        return None
    body = json.dumps({"model": model, "max_tokens": 600,
                       "messages": [{"role": "user", "content": PROMPT + text}]}).encode()
    req = urllib.request.Request(f"{base_url}/v1/messages", data=body, method="POST", headers={
        "x-api-key": api_key, "anthropic-version": "2023-06-01", "content-type": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            data = json.load(r)
        return "**Summary (AI-generated, check against the diff below)**\n\n" + data["content"][0]["text"]
    except Exception:  # a summary is a nicety; never fail the check over it
        return None
