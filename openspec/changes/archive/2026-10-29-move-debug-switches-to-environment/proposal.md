# Proposal

## Why

Testing notifications for a specific day currently means editing `common.py`: `DEBUG_FLAG = True` suppresses mail/SMS and pins "today", `CHANGE_DAY = True` only pins "today" to the hard-coded `DEBUG_TODAY`. Both must be reset to `False` before every commit, and `operations.py` has to inspect the source constants during preflight to catch a forgotten switch. In addition `notifier.py` captures `DEBUG_FLAG` at import time (`from common import DEBUG_FLAG`), so changing `common.DEBUG_FLAG` at runtime does not affect sending while it does affect the date — the two halves of one switch can disagree.

Moving these local-only overrides to environment variables removes the "remember to reset" risk, keeps the committed code free of test dates and lets production refuse them explicitly.

## What Changes

- Replace `DEBUG_FLAG`, `CHANGE_DAY` and `DEBUG_TODAY` with two independent environment variables read at call time:
  - `NULIGAHELPER_TODAY=YYYY-MM-DD` pins `common.effective_today()` (replaces `CHANGE_DAY` + `DEBUG_TODAY`).
  - `NULIGAHELPER_DRY_RUN=1` suppresses outbound mail and SMS in `Notifier.send_Mail` / `send_SMS` (replaces the sending half of `DEBUG_FLAG`). Unlike today it does not pin the date; combine both variables for the old `DEBUG_FLAG` behavior.
- Validate the values strictly (ISO date; `1`/unset) and fail with a German error on invalid input instead of silently using the real date.
- **BREAKING (production):** web startup, the daily job and `operations.py preflight|launch-check` refuse to run when `NULIGAHELPER_ENV=production` and either override is set. This replaces the source-constant check in `operations.py`.
- `run.sh` / `run_webapp.sh` pass the variables through unchanged; README, `AGENTS.md` and `CLAUDE.md` document the new switches and drop the "reset before committing" rule.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `production-operations`: Production preflight and startup reject development date and delivery overrides.

## Impact

- `common.py` (constants removed, `effective_today()` reads `NULIGAHELPER_TODAY`, new `dry_run()` helper), `notifier.py` (call-time check), `operations.py` (environment-based check), `webapp.py` / `main.py` startup guard via one shared helper.
- Tests that set `common.DEBUG_FLAG`/`CHANGE_DAY`/`DEBUG_TODAY` or patch `notifier.DEBUG_FLAG` (`test_age_eligibility.py`, `test_cake_reminders.py`, operations tests) switch to environment patching; existing `patch("common.effective_today", …)` calls keep working.
- Documentation: `README.MD`, `AGENTS.md`, `CLAUDE.md`.
- No schema, message wording or deployment-asset change; production systemd units must not set the new variables (they currently do not).
