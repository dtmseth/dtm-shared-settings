#!/usr/bin/env python3
"""Publish merged settings files to SharePoint /Settings/.

Runs after a PR merges into main. Reads the list of files changed by the
merge (via git diff against the merge's first parent) and uploads each one
that lives under SETTINGS_SUBDIR to /Settings/{filename} on SharePoint.

Each uploaded file is accompanied by a small ``{filename}.meta.json``
sidecar holding the merge SHA, PR number/title, and merger's username.
Power Automate Flow A reads from /Settings/ — the sidecar gives the flow
something richer than just "a file changed" to put in the notification.

Auth: GRAPH_ACCESS_TOKEN in env (same pattern as the pickup workflow).
"""

from __future__ import annotations

import json
import logging
import os
import subprocess
import sys
import urllib.parse
from pathlib import Path

import requests

logger = logging.getLogger("publish")

GRAPH = "https://graph.microsoft.com/v1.0"


def env(name: str, *, required: bool = True, default: str = "") -> str:
    val = os.environ.get(name, default).strip()
    if required and not val:
        sys.exit(f"Missing required env var: {name}")
    return val


def graph_headers(token: str, *, content_type: str | None = None) -> dict[str, str]:
    headers = {"Authorization": f"Bearer {token}"}
    if content_type:
        headers["Content-Type"] = content_type
    return headers


def run(cmd: list[str]) -> str:
    logger.info("$ %s", " ".join(cmd))
    return subprocess.run(
        cmd,
        check=True,
        text=True,
        capture_output=True,
    ).stdout


def changed_files_under(subdir: str, merge_sha: str) -> list[Path]:
    """List files under *subdir* changed by the merge commit *merge_sha*.

    For a merge commit M with first parent P, ``git diff P M`` shows exactly
    the changes the merge introduced. For a fast-forward merge the merge SHA
    is the tip commit and ``HEAD~1..HEAD`` works the same way.
    """
    raw = run(["git", "diff", "--name-only", f"{merge_sha}^", merge_sha])
    out: list[Path] = []
    subdir_normalized = subdir.rstrip("/") + "/"
    for line in raw.splitlines():
        line = line.strip()
        if not line.startswith(subdir_normalized):
            continue
        path = Path(line)
        if not path.exists():
            logger.info("Skipping deleted/renamed-away file: %s", line)
            continue
        out.append(path)
    return out


def upload(token: str, site_id: str, drive_id: str, remote_name: str, data: bytes) -> None:
    encoded = urllib.parse.quote(remote_name, safe="/")
    url = f"{GRAPH}/sites/{site_id}/drives/{drive_id}/root:/Settings/{encoded}:/content"
    resp = requests.put(
        url,
        headers=graph_headers(token, content_type="application/octet-stream"),
        data=data,
        timeout=60,
    )
    resp.raise_for_status()
    logger.info("Uploaded %s (%d bytes)", remote_name, len(data))


def build_metadata(*, merge_sha: str, settings_subdir: str) -> dict:
    pr_number = os.environ.get("PR_NUMBER", "").strip()
    pr_title = os.environ.get("PR_TITLE", "").strip()
    pr_author = os.environ.get("PR_AUTHOR", "").strip()
    return {
        "merge_sha": merge_sha,
        "settings_subdir": settings_subdir,
        "pr_number": pr_number,
        "pr_title": pr_title,
        "pr_author": pr_author,
    }


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    token = env("GRAPH_ACCESS_TOKEN")
    site_id = env("SHAREPOINT_SITE_ID")
    drive_id = env("SHAREPOINT_DRIVE_ID")
    settings_subdir = env("SETTINGS_SUBDIR", required=False, default="resources/config")
    merge_sha = env("MERGE_SHA")

    files = changed_files_under(settings_subdir, merge_sha)
    if not files:
        logger.info("No settings files changed in this merge — nothing to publish")
        return 0

    metadata = build_metadata(merge_sha=merge_sha, settings_subdir=settings_subdir)
    failures = 0

    for path in files:
        remote_name = path.name
        try:
            # Upload the metadata sidecar BEFORE the settings file. Power
            # Automate Flow A triggers on the settings file changing, and
            # the flow reads the sidecar for PR context — landing the
            # sidecar first guarantees it's there when the flow fires.
            meta_payload = {**metadata, "filename": remote_name}
            upload(
                token,
                site_id,
                drive_id,
                f"{remote_name}.meta.json",
                json.dumps(meta_payload, indent=2).encode("utf-8"),
            )
            upload(token, site_id, drive_id, remote_name, path.read_bytes())
        except Exception:
            logger.exception("Failed to publish %s", path)
            failures += 1

    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
