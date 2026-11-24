"""Minimal GitHub REST helpers (stdlib only) for finding changed workbooks and posting one sticky comment."""
from __future__ import annotations

import json
import subprocess
import urllib.request

from .render import MARKER

API = "https://api.github.com"


def _req(method: str, url: str, token: str, payload: dict | None = None):
    req = urllib.request.Request(url, method=method, data=json.dumps(payload).encode() if payload else None, headers={
        "authorization": f"Bearer {token}", "accept": "application/vnd.github+json",
        "content-type": "application/json", "user-agent": "tablient-diffs"})
    with urllib.request.urlopen(req, timeout=30) as r:
        body = r.read()
        return json.loads(body) if body else None


def git(*args: str) -> bytes:
    return subprocess.run(["git", *args], check=True, capture_output=True).stdout


def changed_workbooks(base: str, head: str) -> list[tuple[str, str]]:
    """[(status, path)] for .xlsx/.xlsm files that differ between base and head (merge-base diff, like a PR)."""
    out = git("diff", "--name-status", "--no-renames", "-z", f"{base}...{head}").decode().split("\0")
    pairs = zip(out[0::2], out[1::2])
    return [(s, p) for s, p in pairs if p.lower().endswith((".xlsx", ".xlsm"))]


def merge_base(base: str, head: str) -> str | None:
    """The commit the PR branched from, or None if git cannot find one (e.g. a shallow clone)."""
    try:
        return git("merge-base", base, head).decode().strip() or None
    except subprocess.CalledProcessError:
        return None


def blob(ref: str, path: str) -> bytes | None:
    try:
        return git("show", f"{ref}:{path}")
    except subprocess.CalledProcessError:
        return None


def upsert_comment(repo: str, pr: int, token: str, body: str) -> str:
    """Create the comment, or update the one this tool posted before. Returns "created" or "updated"."""
    url = f"{API}/repos/{repo}/issues/{pr}/comments"
    page = 1
    while True:
        items = _req("GET", f"{url}?per_page=100&page={page}", token)
        for c in items:
            if MARKER in c["body"]:
                _req("PATCH", f"{API}/repos/{repo}/issues/comments/{c['id']}", token, {"body": body})
                return "updated"
        if len(items) < 100:
            break
        page += 1
    _req("POST", url, token, {"body": body})
    return "created"
