## Why

The Flask UI and daily job now share one SQLite database, but connections do not explicitly enforce foreign keys or wait for transient locks, the daily job can overlap itself, and Dropbox backup reads the live database file directly. On a Raspberry Pi this can produce avoidable lock failures, unenforced relationships, or an inconsistent backup—especially once write-ahead logging is enabled for concurrent access.

## What Changes

- Establish a supported SQLite runtime profile for a local Raspberry Pi filesystem: foreign-key enforcement and verification on every connection, a bounded busy timeout, WAL journal mode, and explicit startup validation.
- Support concurrent access by one threaded Flask web process/worker plus one daily-job process, with bounded lock waiting and actionable failure reporting; preserve the existing per-slot assignment compare-and-swap and audit semantics under contention.
- Prevent two daily jobs targeting the same database from running at the same time, including direct invocation paths rather than relying only on cron configuration.
- Replace raw reads of the live SQLite file with transactionally consistent snapshots produced through SQLite's online backup API. Treat the main database, WAL, and shared-memory files as one live database state and never upload a copied main file independently of its WAL.
- Upload a stable latest backup plus bounded, dated recovery points to Dropbox, prune only backups owned by this application, and surface snapshot, upload, listing, and retention failures to logs and the daily job's exit status.
- Document and test backup restoration, including stopping database users, avoiding stale `-wal`/`-shm` sidecars, validating the restored database, and exercising concurrent web/job access.
- Retain the project's current schema lifecycle. This change does not introduce schema migrations or require a database rebuild solely for the SQLite runtime settings.
- Explicit non-goals: PostgreSQL support, a general database migration framework, multi-worker web serving, network-filesystem SQLite, high-availability replication, and public-internet deployment hardening.

## Capabilities

### New Capabilities

- `sqlite-runtime-safety`: Defines connection invariants, WAL and busy-timeout behavior, the initially supported process topology, non-overlapping daily execution, and preservation of assignment CAS behavior during concurrent access.
- `database-backup-recovery`: Defines consistent SQLite snapshots, Dropbox latest/dated retention behavior, complete backup failure visibility, and a verified restore procedure that accounts for WAL sidecars.

### Modified Capabilities

None. Existing `task-self-service` compare-and-swap behavior and `assignment-audit` guarantees remain unchanged and become regression constraints for the new SQLite runtime profile.

## Impact

- Database setup and lifecycle in `db.py`, including SQLAlchemy/SQLite connection hooks and startup checks.
- Daily-job orchestration and Dropbox backup behavior in `main.py` and `run.sh`, plus a guarded recovery command in `manage_db.py`.
- The supported web deployment topology documented around `run_webapp.sh`; production WSGI selection remains separate.
- Dropbox/database configuration examples in `config_template.json` and operational guidance in `README.MD` and test documentation.
- Offline tests, especially `test/test_concurrency.py`, plus new coverage for pragmas, lock contention, overlap prevention, backup consistency/retention/failure handling, and restore validation.
- No new database schema and no migration machinery; WAL mode may create runtime `-wal` and `-shm` sidecar files beside the configured database while it is open.
