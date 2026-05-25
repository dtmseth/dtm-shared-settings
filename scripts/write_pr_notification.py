#!/usr/bin/env python3
"""Write a PR-notification marker file to SharePoint /PRNotifications/.

Power Automate watches the folder; a new file there fires the email-to-Seth
flow. This script is invoked by .github/workflows/notify-new-pr.yml after
the Ollama-based AI summarizer has commented on the PR.

Required env:
  GRAPH_ACCESS_TOKEN   client_credentials token for Graph API
  SHAREPOINT_DRIVE_ID  drive containing the /PRNotifications/ folder
  PR_NUMBER            integer, the PR being announced
  PR_TITLE             string
  PR_URL               https://github.com/.../pull/N
  PR_AUTHOR            github login (may not match the DTM submitter)
  PR_BODY              raw markdown body — contains the proposer's summary
  AI_SUMMARY           text produced by the Ollama action (may be empty)
  OPEN_PR_COUNT        integer (count of currently open PRs in the repo)
  PRS_URL              link to the repo's open PRs page
  REPO_FULL_NAME       owner/repo string
"""

from __future__ import annotations

import json
import logging
import os
import sys
import urllib.parse
from datetime import datetime, timezone

import requests

logger = logging.getLogger("notify_pr")

GRAPH = "https://graph.microsoft.com/v1.0"


def env(name: str, *, required: bool = True, default: str = "") -> str:
    val = os.environ.get(name, default).strip()
    if required and not val:
        sys.exit(f"Missing required env var: {name}")
    return val


def upload(token: str, drive_id: str, remote_path: str, data: bytes) -> None:
    encoded = urllib.parse.quote(remote_path, safe="/")
    url = f"{GRAPH}/drives/{drive_id}/root:/{encoded}:/content"
    resp = requests.put(
        url,
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        },
        data=data,
        timeout=30,
    )
    resp.raise_for_status()


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")

    token = env("GRAPH_ACCESS_TOKEN")
    drive_id = env("SHAREPOINT_DRIVE_ID")
    pr_number = env("PR_NUMBER")
    payload = {
        "schema_version": 1,
        "pr_number": int(pr_number),
        "pr_title": env("PR_TITLE"),
        "pr_url": env("PR_URL"),
        "pr_author": env("PR_AUTHOR"),
        "pr_body": env("PR_BODY", required=False),
        "ai_summary": env("AI_SUMMARY", required=False),
        "open_pr_count": int(env("OPEN_PR_COUNT", required=False, default="0") or 0),
        "prs_url": env("PRS_URL"),
        "repo_full_name": env("REPO_FULL_NAME"),
        "notified_at": datetime.now(timezone.utc).isoformat(),
    }
    body = json.dumps(payload, indent=2).encode("utf-8")
    remote_path = f"PRNotifications/{pr_number}.json"
    upload(token, drive_id, remote_path, body)
    logger.info("Wrote notification marker: %s (%d bytes)", remote_path, len(body))
    return 0


if __name__ == "__main__":
    sys.exit(main())
