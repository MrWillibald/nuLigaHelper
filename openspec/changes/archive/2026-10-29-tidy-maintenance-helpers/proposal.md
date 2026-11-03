# Proposal

## Why

Several small inconsistencies make the codebase harder to read and the test suite less reliable, and they would otherwise be carried into the restructured layout:

- `_readonly_uri()` exists twice (`backup.py`, `schema_migrations.py`) with slightly different signatures.
- `db.ROLE_MV` ("MV" is no longer a per-game duty role) and `db.ROLE_SUPPORT = ROLE_CASH` (alias left from the Unterstützung/Kasse rename) are public domain constants used only by tests, which suggests to readers that they are current roles.
- Development installs `requirements.txt` with unpinned versions, while CI and production install the pinned `requirements-test.txt` / `requirements-production.txt`. Local runs can therefore use different library versions than CI (pandas and SQLAlchemy behavior differences are already documented pitfalls).
- `test_operations.py::test_monitor_detects_staleness_missing_timers_clock_disk_and_tls` passes the real `os.statvfs()` of a temporary directory into the monitor, so the test fails on machines with a nearly full disk (observed on 2026-10-29: `app_disk: 1`).

## What Changes

- Keep one read-only SQLite URI helper in `backup.py` (public `readonly_uri(path, *, immutable=False)`) and use it from `schema_migrations.py`. `release-assets/recovery_check.py` stays standalone by design (it must run without the application package) and is not touched.
- Remove `ROLE_MV` and `ROLE_SUPPORT` from `db.py`; tests that need the literals define them locally (`"MV"`, or `db.ROLE_CASH`).
- Make `requirements.txt` a one-line `-r requirements-test.txt` so local development, CI and production share the reviewed pins; update install instructions.
- Give the monitor test a deterministic fake `statvfs` result for the "healthy" cases so only intended thresholds are exercised.
- The dormant newspaper-article feature (`Notifier.send_article`, commented call in `main.py`, `dormant` message templates) is deliberately left unchanged.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

None. Internal refactoring and test robustness; observable behavior is unchanged.

## Impact

- `backup.py`, `schema_migrations.py`, `db.py`, `requirements.txt`, `test/test_operations.py`, tests referencing the removed constants (`test_day_blocks.py`, `test_game_duties.py`, `test_birth_date_migration.py`, `test_task_descriptions.py`), install instructions in `README.MD`, `AGENTS.md`, `CLAUDE.md`, `test/run_tests.sh`.
- Developers re-run `pip install -r requirements.txt` once; the pinned set targets CPython 3.13 like CI.
