# dtm-shared-settings — workflow drafts

**These files do not belong to the DTM Vehicle Builder app.** They are drafts
intended for the separate `dtm-shared-settings` GitHub repository — the
review backend for cloud-mode settings changes. They live here so they can
be co-evolved with the gateway code that produces the proposals they
process, but nothing in the app reads or executes them.

## What's in this directory

```
.github/workflows/
  pickup-pending-changes.yml      ← cron */5 min — turns SharePoint
                                    PendingChanges files into PRs
  publish-settings-on-merge.yml   ← on PR merge — pushes merged settings
                                    file back to SharePoint /Settings/

scripts/
  pickup_pending_changes.py       ← invoked by pickup workflow
  publish_settings_on_merge.py    ← invoked by publish workflow
  requirements.txt                ← Python deps for both scripts
```

## How to land these in the settings repo

1. In `dtm-shared-settings`, mirror this layout:
   ```
   .github/workflows/pickup-pending-changes.yml
   .github/workflows/publish-settings-on-merge.yml
   scripts/pickup_pending_changes.py
   scripts/publish_settings_on_merge.py
   scripts/requirements.txt
   resources/config/              ← actual settings JSONs live here
   ```
2. Verify GitHub Actions secrets are set in `dtm-shared-settings`:
   - `AZURE_CLIENT_ID` — the **CI** Azure AD app's client ID (not the desktop app)
   - `AZURE_TENANT_ID`
   - `AZURE_CLIENT_SECRET` — the CI app's client secret (rotated every ~6 months)
   - `SHAREPOINT_SITE_ID`
   - `SHAREPOINT_DRIVE_ID`
3. Branch protection on `main`: 1 reviewer required, linear history, no
   direct pushes. Use **classic** branch protection (not Rulesets) — the
   repo is on GitHub Free and Rulesets require Pro/Team to enforce on the
   default branch.

**Note on auth**: these workflows use the **client-credentials + client-secret**
flow (same as the existing `test-sharepoint-connection.yml`). The roadmap
recommends migrating to OIDC federated credentials for least-privilege
review later, but client-credentials is already proven and removes a setup
step. Switching to OIDC later is a workflow-only change (no script edits).

## Schema contract

The pickup script accepts proposal JSONs with `schema_version: 1`, matching
what `SharePointPendingChangesGateway` writes in this app. The fields are
documented in
[`src/dtm_buildsheet/app/adapters/cloud/sharepoint_proposals_gateway.py`](../src/dtm_buildsheet/app/adapters/cloud/sharepoint_proposals_gateway.py).

Bump the version in both places simultaneously if you change the shape, and
add the new version to `SUPPORTED_SCHEMA_VERSIONS` in the pickup script.

## What the workflows do (and don't do)

**pickup-pending-changes.yml** runs every 5 minutes:
- Lists `/PendingChanges/*.json` on SharePoint
- For each one: validates schema, creates a branch
  `change/<user-slug>/<timestamp>`, writes the proposed content to
  `resources/config/<target_file>`, opens a PR, deletes the source
- No-op proposals (where new_content matches the current file) are
  recorded and the source is deleted — no empty PR is created
- Failures leave the source file in place for manual triage

**publish-settings-on-merge.yml** runs once per merged PR:
- Diffs the merge against its first parent to find changed settings files
- Uploads each one to `/Settings/{filename}` on SharePoint
- Also uploads a `{filename}.meta.json` sidecar with the PR number, title,
  and merger username — this is what Power Automate Flow A reads to build
  a useful notification email

**Not handled here** (intentional scope cuts for this milestone):
- Verifying the proposal author has commit rights in the settings repo —
  proposals are open to any tenant user with SharePoint write access
- Squash-vs-merge policy — relies on whatever the repo enforces
- Releases workflow (DMG/EXE → `/Releases/`) — lives in the app repo

## Local dry-run

The pickup script can be smoke-tested locally with a manually obtained
Graph token:

```bash
cd dtm-shared-settings
pip install -r scripts/requirements.txt
export GRAPH_ACCESS_TOKEN=...   # from az account get-access-token
export SHAREPOINT_SITE_ID=...
export SHAREPOINT_DRIVE_ID=...
git checkout -b dryrun-pickup
python scripts/pickup_pending_changes.py
```

Don't push the resulting branch unless you mean to.
