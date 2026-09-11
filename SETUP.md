# Environment Setup — Client Monthly Report Template

This pipeline runs unattended as a scheduled cloud Routine. This doc is written
to be **copied verbatim into the next client's repo** — everything below is
battle-tested against real failures hit setting up Salon Veritas, not
theoretical. Follow it in order; skipping steps produces failures that look
like something else entirely (see the Error Decoder at the bottom).

## 0. Copy and rebrand the pipeline

From an existing client repo (this one, or VanDeVelde's), copy:
`buffer-metrics.js`, `fetch_followers.js`, `generate_report.py`,
`validate_report.py`, `run_monthly_report.sh`, `.gitignore`. Start
`client-context.md` fresh (brand facts + empty Goal Log) and
`follower-history.json` as `[]`.

Rebrand by find-replacing the env var prefix (`SALONVERITAS_` →
`<NEWCLIENT>_`) in `buffer-metrics.js`, `fetch_followers.js`, and
`run_monthly_report.sh`; the `--client` default in `generate_report.py`; the
`COLORS`/`TITLE_FONT`/`BODY_FONT` design tokens if the new client has its own
brand kit or an existing Claude-built report artifact to source them from
(prefer free Google Fonts over a client's licensed brand font — see the
Design System note below); and `context_tie_in_goal()`'s static goal text to
reference whatever this client's evergreen content initiative actually is
(was Printify/Wix for VanDeVelde, a Milbon package for Salon Veritas).

**Design System note:** if the client has a prior Claude-built artifact
(an audit, a report, a landing page) that already established a color/font
system, treat THAT as the source of truth for anything Claude generates —
not the raw brand kit and not a work-in-progress site. Raw brand kits often
specify licensed fonts (Gill Sans MT, etc.) that aren't embeddable in a
generated `.pptx`; an existing report artifact will have already solved that
by picking free Google Font equivalents.

## 1. Make the script self-sufficient (do this before first use, not after it breaks)

Two fixes that cost real debugging time on Salon Veritas — bake them into
`run_monthly_report.sh` from the start for every new client:

**a. Self-install `python-pptx`.** Not every environment has it pre-installed
(unclear why some do and some don't). Add this near the top of the script,
before anything else runs:
```bash
echo "=== Step 0: Ensuring Python dependencies ==="
python3 -c "import pptx" 2>/dev/null || pip3 install --quiet python-pptx
```

**b. Commit the `.pptx` from inside the script itself — do not rely on the
calling Routine session to read `OK_TO_UPLOAD:<filename>` and commit it
separately.** That hand-off looked correct on paper (and is what the
VanDeVelde/Ocean Crest Routine prompts still do) but proved unreliable in
practice: across several live test runs, the script would finish cleanly and
print `OK_TO_UPLOAD:<filename>`, and the outer session would just... not
follow through on the commit, without erroring either. Have the script do it
directly in its own final step:
```bash
if [ "$VALIDATION_OK" -eq 1 ]; then
  git add -f "$OUTPUT" >/dev/null 2>&1
  git commit -m "Automated monthly run: ${MONTH} ${YEAR} (report generated)" >/dev/null 2>&1 || true
  git push origin HEAD:main
  echo "OK_TO_UPLOAD:${OUTPUT}"
  exit 0
fi
```
The Routine's own prompt should then just verify the file landed (`git log -1
--stat`), not attempt its own `git add`/`commit`/`push` for it.

**c. Make the follower-snapshot step non-fatal.** Meta's API is the most
fragile link in this pipeline (see Error Decoder). Don't let a Meta hiccup
block the whole report when the actual post/engagement data (Buffer) is
already in hand:
```bash
if ! node fetch_followers.js > /tmp/followers_out.log 2>/tmp/followers_err.log; then
  echo "[$TIMESTAMP] FOLLOWER FETCH FAILED (non-fatal): $(cat /tmp/followers_err.log)" >> run-log.txt
else
  cat /tmp/followers_out.log
fi
```
`generate_report.py` already handles a missing/empty `follower-history.json`
gracefully (no "Current Followers" card, no error) — this is safe.

## 2. Buffer setup

1. Confirm the client's Instagram/TikTok/Facebook channels are connected in
   Buffer (their own account, or an agency account that manages them) —
   check via Buffer's own UI, or query `list_channels` if a Buffer MCP
   connector is available in-session.
2. Generate a Personal Access Token at `publish.buffer.com/settings/api`,
   logged in as whichever account has those channels connected. **A token
   generated under a different client's Buffer account will not see this
   client's channels** — generate fresh per client, don't reuse.
3. Sanity-check the token works before wiring it into secrets: `curl -X POST
   https://api.buffer.com/graphql -H "Content-Type: application/json" -H
   "Authorization: Bearer YOUR_TOKEN" -d '{"query":"{ account { organizations
   { id } } } "}'`

## 3. Meta (Facebook Page + Instagram) credentials

This is the fiddly part. Follow it exactly — every shortcut here cost real
time on Salon Veritas.

**a. Confirm the Instagram account is a Business/Creator account, linked to
the Facebook Page**, and that both are added as Business Assets in Meta
Business Suite (business.facebook.com → Business Settings → Accounts →
Pages, and → Accounts → Instagram Accounts). **If the Page isn't listed
there, none of the API calls below will expose audience/engagement fields no
matter how many permissions you grant** — this is a separate gate from OAuth
permissions and from the app's own products, and it's easy to miss because
basic fields (`id`, `name`) work fine without it, so the token looks valid
right up until you ask for anything useful.

**b. If the client (or you) administers more than one Page with a similar or
identical name, get the correct Page's numeric ID from Business Settings
FIRST** (click the Page there — the ID is shown at the top), before doing
anything else. Matching a Page by name alone in a token/ID lookup is how a
wrong-but-plausible ID sneaks in and produces confusing "nonexisting field"
errors that look like a permissions problem but are actually just the wrong
Page.

**c. Confirm the app has the right products added.** developers.facebook.com
→ your app → left sidebar → look at **Added products**. For Instagram data,
**Instagram Graph API** needs to be added (not just the `instagram_basic`
permission — permission and product are separate gates). Click "Set up" if
it's only listed under "Available products."

**d. Generate the token chain**, all in Graph API Explorer
(developers.facebook.com/tools/explorer):
   1. Select your app, token type **User Token**, check permissions
      `pages_show_list`, `pages_read_engagement`, `instagram_basic` →
      **Generate Access Token**.
   2. Exchange for a long-lived user token (60 days — confirm `expires_in:
      5184000` in the response):
      ```
      https://graph.facebook.com/v21.0/oauth/access_token?grant_type=fb_exchange_token&client_id=APP_ID&client_secret=APP_SECRET&fb_exchange_token=SHORT_LIVED_TOKEN
      ```
      Run in a browser address bar, or curl **with the whole URL in quotes**
      — an unquoted `&` in a shell is a command separator, not part of the
      URL, and produces a `parse error near '&'`.
   3. Get the Page token, ID, and IG Business ID **all in one call** — don't
      do three separate lookups:
      ```
      https://graph.facebook.com/v21.0/me/accounts?fields=id,name,access_token,instagram_business_account&access_token=LONG_LIVED_USER_TOKEN
      ```
      This returns one entry per Page you administer — **if you manage
      multiple Pages, match by the numeric `id` from step (b), not by
      name.** From the matching entry: `access_token` is the real Meta
      token, `id` is the Page ID, `instagram_business_account.id` is the IG
      Business ID.
   4. Sanity-check before wiring into secrets:
      ```
      https://graph.facebook.com/v21.0/PAGE_ID?fields=followers_count,fan_count,instagram_business_account{followers_count}&access_token=PAGE_TOKEN
      ```
      Real follower numbers back = you're done. Any error here means
      something above is still wrong — see the Error Decoder before
      touching the environment secrets.

**Longevity note:** a Page token derived this way is long-lived but not
permanent — it breaks if the admin who generated it changes their Facebook
password or the app's access is revoked. For a Routine expected to run
indefinitely, a Business Manager System User token (Business Settings →
Users → System Users) is more durable and worth setting up once a client
relationship is established as long-term.

## 4. Cloud environment

1. claude.ai → Settings → Environments → **New Environment**, named
   `<Client> Marketing Report`. One per client — isolates secrets and keeps
   this client's automation from touching another's.
2. Attach the client's repo.
3. **Network policy → Custom domains** (not the default). Add BOTH:
   - `api.buffer.com`
   - `graph.facebook.com`

   Skipping this is the single most misleading failure in this whole setup:
   the request never reaches either API, and the error (`403 Forbidden`)
   looks exactly like a bad credential. The tell is in the full response
   body: `Host not in allowlist: <domain>`. If `buffer-metrics.js`/
   `fetch_followers.js` don't already surface the response body on failure
   (older copies just log the bare status code), fix that first — it's what
   makes this diagnosable at all instead of a guessing game.
4. Add the four secrets (`<CLIENT>_BUFFER_API_KEY`,
   `<CLIENT>_META_ACCESS_TOKEN`, `<CLIENT>_FB_PAGE_ID`,
   `<CLIENT>_IG_BUSINESS_ID`), pasting each value directly from its source
   with no intermediate stop in another app/chat window — that's the most
   common way stray whitespace, quotes, or a trailing newline creep in. If a
   run fails in a way that doesn't match anything in the Error Decoder,
   re-copy the secret fresh rather than assuming the value is right.

## 5. The scheduled Routine

Point it at the client's environment. Prompt pattern (assumes the
self-contained-commit fix from Step 1b is in the script):

```
You are running the fully automated monthly <CLIENT> social media report
pipeline. This runs unattended -- do not ask for confirmation or approval.

1. Run: bash run_monthly_report.sh from the repo root.

2. Read the script's final output line:
- OK_TO_UPLOAD:<filename> -- the script already committed and pushed it.
  Verify with `git log -1 --stat`; do NOT commit it yourself.
- VALIDATION_FAILED_DO_NOT_UPLOAD or ABORTED_BEFORE_UPLOAD -- already
  logged to run-log.txt and pushed. Stop; don't try to fix it yourself.

Do not modify generate_report.py, buffer-metrics.js, fetch_followers.js,
validate_report.py, or run_monthly_report.sh. Do not ask clarifying
questions -- this is a fully unattended run.
```

Schedule to match the client's actual cadence — a monthly meeting on a fixed
date (not necessarily the 1st) may mean the report should be labeled by the
CURRENT month rather than "last month," since `MONTH=$(date -u +%B)` differs
from `MONTH=$(date -u -d "last month" +%B)` depending on which the schedule
implies. Buffer's own current-vs-prior data window is always trailing-30-days
regardless of this label.

## 6. Test-fire before trusting the schedule

Fire the Routine by hand and check `run-log.txt` + commit history rather than
waiting for the real schedule to find out something's misconfigured. Expect
to iterate — this setup has enough moving pieces (Buffer token, Meta token
chain, network allowlist, Python deps) that a clean first fire is the
exception, not the rule.

**One side effect of repeated test-firing:** each successful run appends a
new entry to `client-context.md`'s Goal Log unconditionally. If you fire the
Routine multiple times while debugging in the same month, you'll get
duplicate `### <Month>` sections. Check for this (`grep -n "^### "
client-context.md`) and collapse to one before considering setup done — it's
cosmetic, not a functional bug (checkoff logic always reads the *last*
section), but it's messy for a document the client may see.

## Error Decoder

| Error text (in `run-log.txt`) | Root cause | Fix |
|---|---|---|
| `GraphQL request failed: 403 Forbidden` (bare, no body) | Either a bad Buffer token, OR the environment's network policy silently blocking the request — indistinguishable without the response body | Make sure `buffer-metrics.js`'s error handling includes `res.text()` on failure, not just the status code, then re-run to see which |
| `... 403 Forbidden -- Host not in allowlist: <domain>` | Environment network policy doesn't allow this host | Step 4.3 above |
| `Error validating access token: Session has expired on <time>, ~1-2hrs after generation` | Used a short-lived token (or a Page token derived from one) instead of the long-lived chain | Redo Step 3.d.2-3, don't skip the exchange |
| `(#100) Tried accessing nonexisting field (X)` on a field you know is real, with all permissions confirmed granted | Almost always the Page isn't linked as a Business Asset (Step 3.a), OR it's the wrong Page (Step 3.b — check by ID, not name), OR the required product isn't added (Step 3.c) | Check 3.a → 3.b → 3.c in that order; basic `id,name` fields will still work even when this is broken, which is what makes it confusing |
| `ModuleNotFoundError: No module named 'pptx'` | Environment doesn't have `python-pptx` pre-installed | Step 1.a self-heal |
| `zsh: parse error near '&'` | Ran a multi-parameter URL in a shell without quoting it | Wrap the whole URL in double quotes for curl; browser address bars don't have this problem |
| Script prints `OK_TO_UPLOAD:<filename>` but no `.pptx` shows up in the repo | The old commit hand-off pattern (routine session reads stdout, commits separately) is unreliable | Step 1.b — commit from inside the script |
