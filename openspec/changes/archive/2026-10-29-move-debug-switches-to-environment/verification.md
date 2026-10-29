# Implementation verification

Verified on 2026-10-29 with synthetic settings and mocked providers.

- `common.py`: `DEBUG_FLAG`, `CHANGE_DAY` and `DEBUG_TODAY` are gone. `effective_today()`
  reads `NULIGAHELPER_TODAY` on every call via `pinned_today()`; `dry_run()` reads
  `NULIGAHELPER_DRY_RUN`. Invalid values raise `OverrideError` with a German message
  naming the variable and without echoing the value. Empty values count as unset.
- Production: `refuse_overrides_in_production()` is called first in
  `webapp.create_app()` and `main.main()`; `pinned_today()` and `dry_run()` also refuse
  their variable in production, so tools without an explicit guard cannot use them.
  `runtime_check.check()` (systemd `ExecStartPre` of every unit) returns the variable
  name first, and `operations.py preflight|launch-check` replaced the source-constant
  check with one `p.require` per variable (`operation=failed field=<variable>`).
- `notifier.py` checks `common.dry_run()` at send time; the import-time
  `from common import DEBUG_FLAG` that could disagree with `common` is removed.
- Tests: `test_age_eligibility.py` and `test_cake_reminders.py` use `patch.dict(os.environ)`;
  `test/helpers.py` drops both variables from the caller's shell. New
  `test/test_development_overrides.py` (7 tests) covers pinned/unset/empty dates, invalid
  values, dry run without date pinning, date override still sending, refusal by web
  startup, daily job, accessors and `runtime_check` without creating the database,
  preflight and launch-check field names, and production without overrides.
- Documentation: `AGENTS.md` (module map and environment rules, "reset before
  committing" removed), `README.MD` (local example), `docs/betrieb/private-deployment.md`,
  `test/README.md`, `docs/REFACTORING-VORSCHLAG.md`. `CLAUDE.md` imports `AGENTS.md` and
  needed no change.
- `test/run_tests.sh`: 487 passing checks (480 + 7 new). The only failure is the known
  host-disk dependency in `test_monitor_detects_staleness_missing_timers_clock_disk_and_tls`
  (addressed by `tidy-maintenance-helpers`).
- `openspec validate move-debug-switches-to-environment --strict`: valid.
