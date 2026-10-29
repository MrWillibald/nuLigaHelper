# Implementation verification

Verified on 2026-10-29.

- `CLAUDE.md` reduced to an orientation plus `@AGENTS.md`; its architecture overview,
  corrected route list (taken from the `webapp.py` route decorators), style and
  test-data notes moved into the new `AGENTS.md` "Architecture" section. The obsolete
  "no migration layer" statement was dropped; `AGENTS.md` already holds the Alembic rule.
- `AGENTS.md` module map lists all 25 top-level modules grouped by area; the project
  summary describes the accepted production deployment. The task-description
  maintenance note moved from the README into the conventions.
- `README.MD` reduced from 758 to 90 lines: overview, features, documentation index,
  getting started, testing. Operator sections moved verbatim to `docs/betrieb/`
  (`sqlite-and-backups.md`, `schema-migration.md`, `cli.md`,
  `notification-wording.md`, `private-deployment.md`, index `README.md`); the table
  list moved to `docs/entwicklung/database-schema.md`.
- Removed README behavior chapters were checked against `openspec/specs/`, both user
  manuals and `AGENTS.md`: every bold UI label and code term from those chapters was
  searched in that corpus; the only misses were line-wrapped phrases, examples
  (`+491701234567`, `14:30`, `5/5`) and developer notes, which were moved instead.
- The stale statement that production activation still awaited first-cutover
  acceptance was corrected in a separate commit (acceptance confirmed by the
  maintainer on 2026-10-29). The README claim that the daily job logs to
  `helper.log` was wrong (it logs to stderr / the journal) and was corrected.
- `test/README.md` per-file table (24 of 58 files) replaced by area groups; a script
  confirmed every `test/test_*.py` matches one group pattern.
- `run_webapp.sh` and the `manage_db.py` header fixed; `bash -n run_webapp.sh` passes,
  `manage_db.py --help` output unchanged.
- Link check over all 61 tracked Markdown files outside the archive: 0 missing targets
  or anchors.
- `test/run_tests.sh`: 486 passing checks; the single failure
  `test_monitor_detects_staleness_missing_timers_clock_disk_and_tls` is the known
  host-disk dependency also present before this change (fixed by
  `tidy-maintenance-helpers`).
- `openspec validate refresh-developer-documentation --strict`: valid.
