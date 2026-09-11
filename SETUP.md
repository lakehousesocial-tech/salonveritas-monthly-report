# Environment Setup

This pipeline runs unattended as a scheduled cloud Routine. Setting up a new
client from this template (copy this repo, rebrand `generate_report.py` and
`client-context.md`, rename the `<CLIENT>_*` env vars) requires the following
one-time environment setup — do all of it before the first scheduled run, or
the Routine will fail on its first fire.

## 1. Create a dedicated cloud environment

claude.ai → Settings → Environments → **New Environment**. Name it
`<Client> Marketing Report` (e.g. "Salon Veritas Marketing Report") — each
client gets its own environment so secrets and network access stay isolated
per client, matching the VanDeVelde/Ocean Crest pattern.

## 2. Attach the repo

Point the environment at this client's repo
(`lakehousesocial-tech/<client>-monthly-report`).

## 3. Network policy — select Custom, not the default

**This step is easy to miss and produces a misleading error if skipped.**
The default/restrictive network policy blocks outbound requests entirely; the
failure looks exactly like a bad API key/token (`403 Forbidden`), not a
network problem, unless you read the response body closely.

Under the environment's network settings, choose **Custom domains** and add
both of the following — the pipeline calls both APIs directly from inside
the environment:

- `api.buffer.com` — used by `buffer-metrics.js` (Buffer's GraphQL API)
- `graph.facebook.com` — used by `fetch_followers.js` (Meta Graph API)

If a run fails with `Host not in allowlist: <domain>` in `run-log.txt`, this
step was skipped or a domain was missed — add it here, not by second-guessing
the credentials.

## 4. Add the four secrets

On that same environment's secrets/environment-variables settings, add:

- `<CLIENT>_BUFFER_API_KEY` — a Buffer Personal Access Token (generated at
  `publish.buffer.com/settings/api` while logged into the client's Buffer
  account, or an agency account with the client's channels connected)
- `<CLIENT>_META_ACCESS_TOKEN` — a long-lived **Page** Access Token (not a
  user token — see `fetch_followers.js`'s docstring for the full derivation:
  short-lived user token → long-lived user token → Page token via
  `me/accounts`)
- `<CLIENT>_FB_PAGE_ID` — the numeric Facebook Page ID
- `<CLIENT>_IG_BUSINESS_ID` — the linked Instagram Business Account ID (can
  be read off the same `me/accounts` call by requesting the
  `instagram_business_account` field)

Paste each value directly from its source with no intermediate stop in
another app — that's the most common way stray whitespace/quotes creep in.

## 5. Create the scheduled Routine

Point it at this environment, with a prompt that runs `bash
run_monthly_report.sh` and commits the result on `OK_TO_UPLOAD:<filename>`
(see the VanDeVelde/Ocean Crest Routines for the exact prompt pattern).

## 6. Test-fire before trusting the schedule

Fire the Routine once by hand and check `run-log.txt` / the commit history
rather than waiting for the first scheduled run to find out something's
misconfigured. A clean run ends with a `.pptx` committed to the repo; any
failure logs a specific reason to `run-log.txt`.
