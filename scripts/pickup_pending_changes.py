#!/usr/bin/env python3
"""Pickup script for the dtm-shared-settings repo.

Runs inside GitHub Actions on a five-minute cron. For each JSON file under
/PendingChanges/ in the team's SharePoint library, it:

  1. Reads and validates the proposal payload
  2. Creates a branch in this repo named change/<user-slug>/<timestamp>
  3. Writes new_content to the proposal's target_file on that branch
  4. Opens a PR titled "[<user>] <summary>"
  5. Deletes the source proposal from SharePoint

Auth: expects GRAPH_ACCESS_TOKEN in the environment. The workflow obtains it
via azure/login@v2 (OIDC federated credential) plus `az account get-access-token
--resource https://graph.microsoft.com`. Git operations use GITHUB_TOKEN
auto-provided by Actions; PR creation uses the `gh` CLI.

Settings files in this repo are expected under ``resources/config/`` so that
the layout mirrors how the app bundles them. Override with SETTINGS_SUBDIR.
"""

from __future__ import annotations

import json
import logging
import os
import subprocess
import sys
import urllib.parse
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import requests

logger = logging.getLogger("pickup")

GRAPH = "https://graph.microsoft.com/v1.0"
SUPPORTED_SCHEMA_VERSIONS = {1}


def env(name: str, *, required: bool = True, default: str = "") -> str:
    val = os.environ.get(name, default).strip()
    if required and not val:
        sys.exit(f"Missing required env var: {name}")
    return val


def graph_headers(token: str, *, json_body: bool = False) -> dict[str, str]:
    headers = {"Authorization": f"Bearer {token}"}
    if json_body:
        headers["Content-Type"] = "application/json"
    return headers


def run(cmd: list[str], *, check: bool = True, capture: bool = False) -> str:
    logger.info("$ %s", " ".join(cmd))
    result = subprocess.run(
        cmd,
        check=check,
        text=True,
        capture_output=capture,
    )
    return (result.stdout or "").strip()


def list_pending(token: str, site_id: str, drive_id: str) -> list[dict[str, Any]]:
    url = (
        f"{GRAPH}/sites/{site_id}/drives/{drive_id}/root:/PendingChanges:/children"
    )
    out: list[dict[str, Any]] = []
    while url:
        resp = requests.get(url, headers=graph_headers(token), timeout=30)
        if resp.status_code == 404:
            logger.warning("PendingChanges folder not found")
            return []
        resp.raise_for_status()
        payload = resp.json()
        for item in payload.get("value", []):
            if "folder" in item:
                continue
            name = item.get("name", "")
            if not name.endswith(".json"):
                continue
            out.append(item)
        url = payload.get("@odata.nextLink")
    return out


def read_proposal(token: str, site_id: str, drive_id: str, name: str) -> dict[str, Any]:
    encoded = urllib.parse.quote(name, safe="/")
    url = f"{GRAPH}/sites/{site_id}/drives/{drive_id}/root:/PendingChanges/{encoded}:/content"
    resp = requests.get(url, headers=graph_headers(token), timeout=30)
    resp.raise_for_status()
    return resp.json()


def delete_proposal(token: str, site_id: str, drive_id: str, name: str) -> None:
    encoded = urllib.parse.quote(name, safe="/")
    url = f"{GRAPH}/sites/{site_id}/drives/{drive_id}/root:/PendingChanges/{encoded}"
    resp = requests.delete(url, headers=graph_headers(token), timeout=30)
    if resp.status_code == 404:
        logger.warning("Proposal %s already gone — race with another runner", name)
        return
    resp.raise_for_status()


def validate(proposal: dict[str, Any]) -> str | None:
    version = proposal.get("schema_version")
    if version not in SUPPORTED_SCHEMA_VERSIONS:
        return f"unsupported schema_version: {version!r}"
    for field in ("proposal_id", "target_file", "new_content", "summary", "submitted_by"):
        if not proposal.get(field):
            return f"missing required field: {field}"
    if "/" in proposal["target_file"] or proposal["target_file"].startswith("."):
        return f"target_file must be a bare filename, got {proposal['target_file']!r}"
    return None


def open_pr_for_proposal(
    proposal: dict[str, Any],
    *,
    settings_subdir: str,
    main_branch: str,
) -> bool:
    """Returns True if a PR was opened (proposal should be deleted from SharePoint)."""
    target = proposal["target_file"]
    submitted_by = proposal["submitted_by"]
    summary = proposal["summary"]
    slug = proposal.get("submitted_by_slug") or "user"
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    branch = f"change/{slug}/{timestamp}"

    target_path = Path(settings_subdir) / target

    # Fresh branch off the latest main.
    run(["git", "fetch", "origin", main_branch])
    run(["git", "checkout", "-B", branch, f"origin/{main_branch}"])
    target_path.parent.mkdir(parents=True, exist_ok=True)
    target_path.write_text(proposal["new_content"], encoding="utf-8")

    diff = run(["git", "status", "--porcelain", str(target_path)], capture=True)
    if not diff:
        logger.info("Proposal %s is a no-op vs current %s — skipping PR",
                    proposal["proposal_id"], target_path)
        return True  # still delete the source; nothing to do

    run(["git", "add", str(target_path)])
    commit_msg = f"[{submitted_by}] {summary}\n\nProposal ID: {proposal['proposal_id']}"
    run(["git", "commit", "-m", commit_msg])
    run(["git", "push", "-u", "origin", branch])

    body_lines = [
        f"**Submitted by:** {submitted_by}",
        f"**Email:** {proposal.get('submitted_by_email', '')}",
        f"**Submitted at:** {proposal.get('submitted_at', '')}",
        f"**Proposal ID:** `{proposal['proposal_id']}`",
        "",
        "## Summary",
        summary,
        "",
        "_This PR was opened automatically by `pickup-pending-changes.yml`._",
    ]
    pr_body = "\n".join(body_lines)
    run(
        [
            "gh", "pr", "create",
            "--base", main_branch,
            "--head", branch,
            "--title", f"[{submitted_by}] {summary}",
            "--body", pr_body,
        ]
    )
    return True


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    token = env("GRAPH_ACCESS_TOKEN")
    site_id = env("SHAREPOINT_SITE_ID")
    drive_id = env("SHAREPOINT_DRIVE_ID")
    settings_subdir = env("SETTINGS_SUBDIR", required=False, default="resources/config")
    main_branch = env("MAIN_BRANCH", required=False, default="main")

    # Identify the bot for the commit history.
    run(["git", "config", "user.name", "dtm-pickup-bot"])
    run(["git", "config", "user.email", "noreply@dtmfleet.com"])

    items = list_pending(token, site_id, drive_id)
    logger.info("Found %d pending proposal(s)", len(items))
    failures = 0

    for item in items:
        name = item["name"]
        try:
            proposal = read_proposal(token, site_id, drive_id, name)
        except Exception:
            logger.exception("Could not read proposal %s — leaving in place", name)
            failures += 1
            continue

        problem = validate(proposal)
        if problem:
            logger.error("Proposal %s rejected: %s — leaving in place for manual triage",
                         name, problem)
            failures += 1
            continue

        try:
            open_pr_for_proposal(
                proposal,
                settings_subdir=settings_subdir,
                main_branch=main_branch,
            )
            delete_proposal(token, site_id, drive_id, name)
        except Exception:
            logger.exception("Failed to process proposal %s", name)
            failures += 1

    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
