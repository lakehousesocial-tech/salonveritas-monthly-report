#!/usr/bin/env bash
# Monthly Salon Veritas report pipeline -- run by the scheduled cloud
# Routine (see the Routine's own prompt for what happens with this script's
# output). Designed to run unattended: every failure path logs to
# run-log.txt and commits/pushes it so it's visible in the repo, then exits
# nonzero so the caller knows not to upload.
#
# Requires exactly these four environment variables (configured in the
# Routine's Environment secrets, not in this script or the repo):
#   SALONVERITAS_BUFFER_API_KEY
#   SALONVERITAS_META_ACCESS_TOKEN
#   SALONVERITAS_FB_PAGE_ID
#   SALONVERITAS_IG_BUSINESS_ID
#
# Step 2 (follower snapshot) is NON-FATAL, unlike VanDeVelde/Ocean Crest --
# Meta's Graph API has been rejecting every Page-level field on this account
# beyond basic id/name (followers_count, fan_count, instagram_business_account
# all error as "nonexisting field" despite the token carrying every relevant
# permission) even though the same Meta-setup steps work for other clients.
# Root cause not yet resolved -- see run-log.txt history and SETUP.md. Rather
# than block the whole monthly report on an unresolved Meta-side issue,
# a failure here just logs a warning and the pipeline continues without a
# "Current Followers" card (generate_report.py already handles a missing/
# empty follower-history.json gracefully). Revisit and make this fatal again
# once the Meta follower-count query is confirmed working.
#
# Unlike VanDeVelde (scheduled for the 1st, reporting on the month that just
# ended), this Routine is scheduled for the 15th of each month, ahead of the
# client's mid-month check-in -- so MONTH/YEAR label the CURRENT month, not
# last month. buffer-metrics.js's own current-vs-prior window is unaffected
# by this label (it's always a trailing 30 days from run time); this only
# changes what the deck's cover/title/filename call the reporting period.
#
# IMPORTANT: cloud Routine sessions check out their own throwaway branch
# (claude/<random-name>) instead of committing directly to main. A plain
# `git push` would push THAT branch and leave main untouched -- silently
# stranding this run's client-context.md / follower-history.json / run-log.txt
# updates somewhere next month's fresh clone of main will never see, which
# breaks goal-checkoff and follower-growth tracking even on a fully
# successful run. Every push below explicitly targets main regardless of
# the local branch name.
set -uo pipefail
cd "$(dirname "$0")"

echo "=== Step 0: Ensuring Python dependencies ==="
python3 -c "import pptx" 2>/dev/null || pip3 install --quiet python-pptx

MONTH=$(date -u +%B)
YEAR=$(date -u +%Y)
OUTPUT="Salon Veritas Monthly Report — ${MONTH} ${YEAR}.pptx"
TIMESTAMP="$(date -u +%FT%TZ)"

log_failure_and_push() {
  echo "[$TIMESTAMP] $1" >> run-log.txt
  git add run-log.txt >/dev/null 2>&1
  git commit -m "Run failed (${MONTH} ${YEAR}): $2" >/dev/null 2>&1
  git push origin HEAD:main >/dev/null 2>&1
  echo "ABORTED_BEFORE_UPLOAD"
}

echo "=== Step 1: Pulling Buffer metrics (current vs. prior period) ==="
if ! node buffer-metrics.js > /tmp/buffer_data.json 2>/tmp/buffer_err.log; then
  log_failure_and_push "BUFFER FETCH FAILED: $(cat /tmp/buffer_err.log)" "buffer-metrics fetch"
  exit 1
fi

echo "=== Step 2: Pulling follower snapshot (Meta Graph API) ==="
if ! node fetch_followers.js > /tmp/followers_out.log 2>/tmp/followers_err.log; then
  echo "[$TIMESTAMP] FOLLOWER FETCH FAILED (non-fatal -- continuing without a follower snapshot this period): $(cat /tmp/followers_err.log)" >> run-log.txt
  echo "WARNING: follower snapshot failed, continuing without it -- see run-log.txt for details"
else
  cat /tmp/followers_out.log
fi

echo "=== Step 3: Generating deck + auto goals ==="
if ! python3 generate_report.py --data /tmp/buffer_data.json --output "$OUTPUT" --month "$MONTH" \
     > /tmp/generate_out.log 2>&1; then
  log_failure_and_push "REPORT GENERATION FAILED: $(tail -30 /tmp/generate_out.log)" "generate_report"
  exit 1
fi
cat /tmp/generate_out.log   # surface the generated goals + checkoff results in the routine's own run log

echo "=== Step 4: Sanity-checking the deck before upload ==="
VALIDATION_OK=1
python3 validate_report.py "$OUTPUT" || VALIDATION_OK=0

echo "=== Step 5: Persisting updated state back to the repo ==="
git add client-context.md follower-history.json run-log.txt >/dev/null 2>&1

if [ "$VALIDATION_OK" -eq 1 ]; then
  # Commit the .pptx here directly rather than leaving it to the calling
  # Routine session to read this script's stdout and commit it separately --
  # that hand-off proved unreliable in practice (the script would finish
  # cleanly and print OK_TO_UPLOAD, but the outer session wouldn't always
  # follow through on the commit). Doing it here removes that failure mode.
  git add -f "$OUTPUT" >/dev/null 2>&1
  git commit -m "Automated monthly run: ${MONTH} ${YEAR} (report generated)" >/dev/null 2>&1 || true
  git push origin HEAD:main
  echo "OK_TO_UPLOAD:${OUTPUT}"
  exit 0
else
  git commit -m "Automated monthly run: ${MONTH} ${YEAR} (validation failed, report not committed)" >/dev/null 2>&1 || true
  git push origin HEAD:main
  echo "VALIDATION_FAILED_DO_NOT_UPLOAD"
  exit 1
fi
