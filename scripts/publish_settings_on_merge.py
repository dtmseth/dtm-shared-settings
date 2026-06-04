#!/usr/bin/env python3
"""Publish merged settings files to SharePoint /Settings/.

Runs after a PR merges into main. Reads the list of files changed by the
merge (via git diff against the merge's first parent) and uploads each one
that lives under SETTINGS_SUBDIR to /Settings/{relative_path} on SharePoint,
preserving any subdirectory structure under the settings root.

Each uploaded file is accompanied by a small ``{filename}.meta.json``
sidecar holding the merge SHA, PR number/title, and merger's username.
Power Automate Flow A reads from /Settings/ — the sidecar gives the flow
something richer than just "a file changed" to put in the notification.

Two modes:
- ``diff`` (default): publish only the files changed by ``MERGE_SHA``. Used
  by the on-merge trigger and the post-auto-merge dispatch from the pickup
  workflow.
- ``bulk``: walk SETTINGS_SUBDIR recursively and publish everything. Used
  to rebuild /Settings/ after an outage or as an initial backfill.

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
    """List files under *subdir* whose content changed in the merge.

    Excludes deletions — those are handled separately by
    deleted_files_under() so we can run a Graph DELETE against SharePoint
    instead of an upload.

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
            # Treated as a delete; deleted_files_under picks these up.
            continue
        out.append(path)
    return out


def deleted_files_under(subdir: str, merge_sha: str) -> list[Path]:
    """List files under *subdir* that were REMOVED by the merge commit.

    `git diff --name-only --diff-filter=D` returns paths that existed in
    the first parent but not in the merge. Those are what we need to
    mirror-delete on SharePoint /Settings/ so a pickup-workflow delete
    proposal actually propagates to teammates.
    """
    raw = run([
        "git", "diff", "--name-only", "--diff-filter=D",
        f"{merge_sha}^", merge_sha,
    ])
    out: list[Path] = []
    subdir_normalized = subdir.rstrip("/") + "/"
    for line in raw.splitlines():
        line = line.strip()
        if not line.startswith(subdir_normalized):
            continue
        # Don't sanity-check existence — the file is supposed to be gone.
        out.append(Path(line))
    return out


def all_files_under(subdir: str) -> list[Path]:
    """Every regular file in *subdir*, recursively. Used by bulk-publish mode
    to rebuild SharePoint /Settings/ from scratch after an outage."""
    base = Path(subdir)
    if not base.exists():
        return []
    return sorted(p for p in base.rglob("*") if p.is_file())


def remote_name_for(path: Path, subdir: str) -> str:
    """Return the SharePoint-relative path that *path* should land at.

    Preserves the subdirectory structure under SETTINGS_SUBDIR so per-record
    entities like ``resources/config/agencies/foo.json`` upload to
    ``/Settings/agencies/foo.json`` instead of getting flattened. The
    sidecar uses the same relative path with ``.meta.json`` appended.
    """
    return path.relative_to(subdir).as_posix()


def delete_remote(token: str, site_id: str, drive_id: str, remote_name: str) -> bool:
    """DELETE a file under /Settings/ via Graph. Returns True on success,
    True for 404 (already gone), False for other failures (logged)."""
    encoded = urllib.parse.quote(remote_name, safe="/")
    url = f"{GRAPH}/sites/{site_id}/drives/{drive_id}/root:/Settings/{encoded}"
    resp = requests.delete(url, headers=graph_headers(token), timeout=30)
    if resp.status_code == 404:
        logger.info("Remote %s already absent — nothing to delete", remote_name)
        return True
    if not resp.ok:
        logger.error("Failed to delete remote %s: HTTP %d", remote_name, resp.status_code)
        return False
    logger.info("Deleted remote %s", remote_name)
    return True


def upload(token: str, site_id: str, drive_id: str, remote_name: str, data: bytes) -> None:
    # quote(safe="/") preserves the slashes between subdir segments so they
    # round-trip through the Graph path-addressing convention intact.
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


def select_files(*, mode: str, settings_subdir: str, merge_sha: str) -> list[Path]:
    """Pick the set of files to publish based on the requested mode."""
    if mode == "bulk":
        return all_files_under(settings_subdir)
    if mode != "diff":
        sys.exit(f"Unknown PUBLISH_MODE: {mode!r} (expected 'diff' or 'bulk')")
    if not merge_sha:
        sys.exit("MERGE_SHA is required for diff mode")
    return changed_files_under(settings_subdir, merge_sha)


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    token = env("GRAPH_ACCESS_TOKEN")
    site_id = env("SHAREPOINT_SITE_ID")
    drive_id = env("SHAREPOINT_DRIVE_ID")
    settings_subdir = env("SETTINGS_SUBDIR", required=False, default="resources/config")
    mode = env("PUBLISH_MODE", required=False, default="diff").lower()
    merge_sha = env("MERGE_SHA", required=False, default="")

    files = select_files(mode=mode, settings_subdir=settings_subdir, merge_sha=merge_sha)
    # Only diff mode tracks deletions — bulk-publish is a full re-upload that
    # doesn't need to enumerate what's gone. Deletions on bulk mode would
    # require listing /Settings/ + diffing against the repo, which is a more
    # involved cleanup than we want a bulk run to perform.
    deletions: list[Path] = []
    if mode == "diff" and merge_sha:
        deletions = deleted_files_under(settings_subdir, merge_sha)

    if not files and not deletions:
        logger.info("Nothing to publish (mode=%s)", mode)
        return 0

    metadata = build_metadata(merge_sha=merge_sha, settings_subdir=settings_subdir)
    logger.info(
        "Publishing %d upload(s) and %d deletion(s) in mode=%s",
        len(files), len(deletions), mode,
    )
    failures = 0

    for path in files:
        remote_name = remote_name_for(path, settings_subdir)
        # Don't recursively publish stale sidecars — they get regenerated
        # alongside their data file in the same pass.
        if remote_name.endswith(".meta.json"):
            logger.info("Skipping sidecar (will be regenerated): %s", remote_name)
            continue
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

    # Mirror deletions: a merge that removed a file from resources/config/
    # should also remove it from SharePoint /Settings/. Drop the sidecar
    # in the same pass so Power Automate notifications don't reference a
    # file that's no longer there.
    for path in deletions:
        remote_name = remote_name_for(path, settings_subdir)
        if remote_name.endswith(".meta.json"):
            continue
        if not delete_remote(token, site_id, drive_id, remote_name):
            failures += 1
            continue
        # Best-effort sidecar cleanup; ignore failures so a missing sidecar
        # doesn't fail the whole publish.
        delete_remote(token, site_id, drive_id, f"{remote_name}.meta.json")

    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
