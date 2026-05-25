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
import re
import subprocess
import sys
import urllib.parse
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

import requests

logger = logging.getLogger("pickup")

GRAPH = "https://graph.microsoft.com/v1.0"
# v1 is the pre-Phase-2-β format with no `category`. v2 (Phase 2-β) adds
# `category: "general" | "advanced"` which drives two-tier review behavior.
SUPPORTED_SCHEMA_VERSIONS = {1, 2}

# Categories accepted in the v2 schema. "general" → auto-merge after opening
# the PR; "advanced" → leave the PR for owner review on github.com.
KNOWN_CATEGORIES = {"general", "advanced"}

# Default category for v1 proposals (no `category` field). Advanced is the
# safe choice — anything written under the old contract continues to require
# manual review, exactly as it does today.
DEFAULT_CATEGORY = "advanced"

# GitHub caps PR titles at 256 chars. Leaving headroom for the `[user] `
# prefix while still keeping the title scannable.
PR_TITLE_SUMMARY_MAX = 200

# Characters allowed in the user-slug segment of a branch name. Anything
# else gets normalized to a dash so git ref-name rules are never violated.
_SAFE_SLUG_RE = re.compile(r"[^a-z0-9_-]+")


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


def _validate_target_file(target: str) -> str | None:
    """Allow `foo.json` or `subdir/foo.json`. Reject anything that could
    traverse out of the settings tree.

    Per-record entities (agencies, sales reps, presets) live in subdirectories
    under ``resources/`` and need the one-subdir form. Anything deeper or with
    ``..`` segments is rejected — the workflow holds repo-write credentials,
    so path traversal is a real risk to guard against.
    """
    if not target or target.startswith("/") or target.startswith("."):
        return f"target_file must be relative, got {target!r}"
    parts = target.split("/")
    if len(parts) > 2:
        return f"target_file may have at most one subdirectory, got {target!r}"
    for seg in parts:
        if not seg or seg in {".", ".."} or seg.startswith(".") or "\\" in seg or "\x00" in seg:
            return f"target_file segment invalid: {seg!r}"
    if not parts[-1].endswith(".json"):
        return f"target_file must end in .json, got {target!r}"
    return None


def resolve_category(proposal: dict[str, Any]) -> str:
    """Map the proposal payload's category onto a known value.

    v1 proposals (no field) default to advanced. v2 proposals with an unknown
    category also default to advanced — the safe choice if a future version
    introduces a new tier the workflow doesn't yet understand.
    """
    raw = str(proposal.get("category", "")).strip().lower()
    if raw in KNOWN_CATEGORIES:
        return raw
    return DEFAULT_CATEGORY


def validate(proposal: dict[str, Any]) -> str | None:
    version = proposal.get("schema_version")
    if version not in SUPPORTED_SCHEMA_VERSIONS:
        return f"unsupported schema_version: {version!r}"
    for field in ("proposal_id", "target_file", "new_content", "summary", "submitted_by"):
        if not proposal.get(field):
            return f"missing required field: {field}"
    target_problem = _validate_target_file(proposal["target_file"])
    if target_problem:
        return target_problem
    return None


def _sanitize_slug(raw: str) -> str:
    """Normalize a slug to characters safe inside a git ref name."""
    cleaned = _SAFE_SLUG_RE.sub("-", raw.lower()).strip("-")
    return cleaned or "user"


def _build_commit_message(submitted_by: str, summary: str, proposal_id: str) -> str:
    """Compose a git commit message that respects the subject/body convention.

    Git tooling assumes a blank line separates the first-line subject from the
    body; without one, ``git log --oneline`` and many UIs garble multi-line
    summaries. The proposal_id trailer always lands in the body.
    """
    summary_lines = summary.splitlines() if summary else [""]
    subject = f"[{submitted_by}] {summary_lines[0]}".rstrip()
    extra_body = "\n".join(summary_lines[1:]).strip()
    trailer = f"Proposal ID: {proposal_id}"
    body = f"{extra_body}\n\n{trailer}" if extra_body else trailer
    return f"{subject}\n\n{body}"


def open_pr_for_proposal(
    proposal: dict[str, Any],
    *,
    settings_subdir: str,
    main_branch: str,
) -> bool:
    """Returns True if a PR was opened (proposal should be deleted from SharePoint).

    For category=general, also enables immediate squash-merge so the PR lands
    without owner review. Branch protection on the settings repo must permit
    ``github-actions[bot]`` to bypass the review requirement; otherwise the
    merge call fails and the PR is left open for manual handling.
    """
    target = proposal["target_file"]
    submitted_by = proposal["submitted_by"]
    summary = proposal["summary"]
    category = resolve_category(proposal)
    # PR title field has practical limits and breaks on newlines — use the
    # first line only, truncated to GitHub's effective ceiling. The full
    # multi-line summary still lands in the body.
    title_summary = (
        summary.splitlines()[0][:PR_TITLE_SUMMARY_MAX] if summary else ""
    )
    slug = _sanitize_slug(proposal.get("submitted_by_slug") or "user")
    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    # proposal_id segment guarantees uniqueness even if two runners process
    # different proposals from the same user in the same second.
    branch = f"change/{slug}/{timestamp}-{proposal['proposal_id']}"

    target_path = Path(settings_subdir) / target

    # Fresh branch off the latest main — caller already fetched once per run.
    run(["git", "checkout", "-B", branch, f"origin/{main_branch}"])
    target_path.parent.mkdir(parents=True, exist_ok=True)
    # The schema says new_content is a string, but be forgiving — older or
    # manually-authored proposals may carry a JSON object instead.
    content = proposal["new_content"]
    if not isinstance(content, str):
        content = json.dumps(content, indent=2)
    target_path.write_text(content, encoding="utf-8")

    diff = run(["git", "status", "--porcelain", str(target_path)], capture=True)
    if not diff:
        logger.info("Proposal %s is a no-op vs current %s — skipping PR",
                    proposal["proposal_id"], target_path)
        return True  # still delete the source; nothing to do

    run(["git", "add", str(target_path)])
    commit_msg = _build_commit_message(submitted_by, summary, proposal["proposal_id"])
    run(["git", "commit", "-m", commit_msg])
    run(["git", "push", "-u", "origin", branch])

    body_lines = [
        f"**Submitted by:** {submitted_by}",
        f"**Email:** {proposal.get('submitted_by_email', '')}",
        f"**Submitted at:** {proposal.get('submitted_at', '')}",
        f"**Category:** {category}",
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
            "--title", f"[{submitted_by}] {title_summary}",
            "--body", pr_body,
        ]
    )

    if category == "general":
        # Branch protection bypass for github-actions[bot] is what makes this
        # work. If the merge call fails (bypass not configured, merge conflict,
        # etc.) the PR stays open for owner review — equivalent to advanced.
        try:
            run(
                [
                    "gh", "pr", "merge", branch,
                    "--squash", "--delete-branch",
                ]
            )
        except subprocess.CalledProcessError:
            logger.exception(
                "Auto-merge failed for general-category PR %s — leaving open for review",
                branch,
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
    # Fetch once per run — every proposal branches off the same origin/main.
    run(["git", "fetch", "origin", main_branch])

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
