# Implementation verification

Verified on 2026-10-24 using synthetic databases and mocked notification transports.

- `test/run_tests.sh`: exit 0; 487 passing checks across 58 test files, including
  the existing authentication, refusal, concurrency, migration, notification and
  dependency-free JavaScript regressions.
- `test/test_schedule_visibility.py`: 13 feature regressions covering inclusive
  calendar dates, persistent settings, active-admin/CSRF/expected-state checks,
  concurrent edits and bounded contention, schedule filtering and block times,
  expired boundaries, season changes, unchanged season totals, hidden-date direct
  claims/takeovers and read-only previews versus ordinary notification/lock flows.
- `test/test_schedule_visibility_migration.py`: four guarded migration regressions
  covering the single successor head, exact fresh/upgrade storage, unrestricted
  singleton initialization, retention of all domain rows, invalid state/drift
  rejection, configured-head reruns and restoration of the validated snapshot.
  Existing earlier migration tests cover the supported historical chains and
  their previously reviewed transformations.
- `test/browser_schedule_visibility.mjs`: eight passing Chromium checks with
  JavaScript disabled and enabled; native save/clear, keyboard navigation,
  administrator full-date mode across both pages, 320/360 CSS pixels and doubled
  text at 640 CSS pixels. Screenshots were inspected for the new controls.
  The opt-in command and runtime options are documented in `test/README.md`.
- `openspec validate limit-visible-home-game-days --strict`: valid.
- `git diff --check`: clean. `DEBUG_FLAG` and `CHANGE_DAY` remain `False`.

The four delta specifications and README usage/deployment instructions were
reviewed against the implementation. The visibility predicate is called only
from schedule preparation and outstanding-statistics preparation. Setting writes
are outside the due-lock mutation preflight. No production database migration or
provider traffic was part of this verification.

Deployment requires the documented stopped-writers migration to
`0009_schedule_visibility`; recovery restores the validated pre-upgrade snapshot
with the matching prior application and all writers stopped.
