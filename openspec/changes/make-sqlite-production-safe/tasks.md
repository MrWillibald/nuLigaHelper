## 1. Establish the SQLite runtime profile

- [ ] 1.1 Centralize the 5-second sqlite3/SQLAlchemy timeout and per-connection `foreign_keys=ON`, `busy_timeout=5000`, and `synchronous=FULL` setup in `db.make_engine()`, and verify a fresh connection reports every configured value in focused database tests.
- [ ] 1.2 Extend database initialization to establish and verify WAL mode and run `foreign_key_check` before application work, and verify tests cover successful startup plus actionable failures for a non-WAL result and pre-existing orphaned references.
- [ ] 1.3 Replace direct application and test `create_engine()` calls with `db.make_engine()` where the production profile is required, and verify the suite no longer creates unconfigured file-backed application engines.
- [ ] 1.4 Audit web, daily-job, management-CLI, and authentication transaction boundaries so commits/rollbacks occur before network I/O and sessions close promptly, and verify targeted tests demonstrate no write transaction remains open while a mocked scraper, notifier, or Dropbox call executes.

## 2. Preserve concurrent assignment semantics

- [ ] 2.1 Add bounded SQLite lock/`BUSY_SNAPSHOT` classification and fresh-transaction revalidation for assignment claims/releases without a generic whole-request retry, and verify stale expectations still produce `SlotConflictError` with the current occupant while exhausted lock deadlines produce a distinct temporary-unavailability error.
- [ ] 2.2 Keep assignment mutation and audit creation in one transaction across success, conflict, integrity failure, and lock timeout paths, and verify focused tests assert exactly one winner/audit for concurrent claims and no audit for stale or timed-out mutations.
- [ ] 2.3 Map exhausted SQLite contention in Flask to HTTP 503 with the established page/JSON response shapes while retaining HTTP 409 for CAS conflicts, and verify endpoint tests distinguish the two statuses and roll back the request session.
- [ ] 2.4 Expand `test/test_concurrency.py` to use independent WAL-configured engines/connections and cover overlapping read/write, a lock released within the wait, timeout beyond the wait, stale claim, stale release after replacement, and one-person-per-game races; verify the concurrency test module passes offline.

## 3. Prevent overlapping daily jobs

- [ ] 3.1 Implement a process-lifetime non-blocking `fcntl.flock` guard at `<resolved-db-path>.daily.lock` and acquire it in `main.main()` before database initialization or external effects, and verify lock ownership is released when its context/process ends.
- [ ] 3.2 Add tests proving a second invocation for the same canonical database fails before scrape/sync/backup/notification calls, a leftover unlocked file does not block, and different database paths remain independent; verify the daily-lock tests pass without spawning network work.
- [ ] 3.3 Keep `run.sh` as a failure-propagating launcher and document the application-level overlap guard for cron/manual execution, verifying a simulated overlap returns a nonzero shell-visible status.

## 4. Create and validate online SQLite snapshots

- [ ] 4.1 Extract backup/recovery helpers with staged error reporting and implement a unique local snapshot using sqlite3 `Connection.backup()` with the 5-second busy policy and a 30-second monotonic overall deadline; verify tests prove a permanently busy source fails within the bound rather than hanging.
- [ ] 4.2 Validate completed snapshots with `quick_check` and `foreign_key_check` before exposing bytes for upload, and verify corrupt/orphaned candidates are rejected before any mocked Dropbox method is called.
- [ ] 4.3 Guarantee cleanup of temporary snapshot files and sidecars on success and every failure path while preserving/logging both primary and cleanup errors, and verify filesystem tests leave no active temporary artifact and retain all staged failure information.
- [ ] 4.4 Add WAL-specific snapshot tests with independent connections and committed WAL-resident data, uncommitted data, and a concurrent writer, verifying each snapshot is self-contained, contains one complete committed state, and restores without source `-wal`/`-shm` files.

## 5. Publish and retain Dropbox recovery points

- [ ] 5.1 Add and validate optional `club.dropbox.dated_retention` configuration with default 14, plus exact latest/dated path generation from the database basename and `common.effective_today()`, and verify absent, valid, zero, negative, non-integer, unusual-basename, and same-day cases.
- [ ] 5.2 Replace live-file upload with validated snapshot publication in dated-then-latest order, and verify a fake Dropbox client receives identical snapshot bytes at the expected paths and never reads/uploads the live `.db`, `-wal`, or `-shm` files.
- [ ] 5.3 Implement fully paginated Dropbox listing and anchored oldest-first pruning of only matching dated objects beyond the configured limit, and verify tests preserve latest, unrelated files, malformed names, similarly prefixed databases, and the newest configured number across multiple pages.
- [ ] 5.4 Make snapshot, validation, read, Dropbox construction, dated/latest upload, pagination, prune, and cleanup errors fail-visible without logging credentials, and verify fault-injection tests assert the exact failed stage, partial publication outcome, and unsuccessful job result for each class of failure.
- [ ] 5.5 Update daily orchestration to remember backup failure, continue safe independent notification work, log all fatal outcomes, and exit nonzero afterward without retrying notifications; verify tests cover backup-only failure and combined backup/notification failure without losing either error.

## 6. Provide WAL-safe restoration

- [ ] 6.1 Add `manage_db.py restore-snapshot SNAPSHOT --confirm-stopped` using shared snapshot validation, same-filesystem temporary copy/fsync, preservation of the active `.db`/`-wal`/`-shm` set in a unique rollback directory, atomic replacement, and recovery on installation failure; verify the CLI refuses to run without explicit stopped-service confirmation.
- [ ] 6.2 Add offline restore tests proving validation happens before mutation, invalid candidates leave the active database and sidecars untouched, an injected replacement failure recovers the preserved set, and a valid standalone WAL snapshot restores committed data without source sidecars.
- [ ] 6.3 Document backup names/retention, all failure semantics, local-filesystem and one-web-worker limits, expected WAL sidecars/checkpoint behavior, deployment preflight, the complete stop-preserve-restore-smoke-test procedure, and rollback in `README.MD`/test docs; verify every documented command and configuration key matches the implemented CLI/template.
- [ ] 6.4 Update `config_template.json` with the retention setting and operationally safe example, explicitly retain the no-migration policy and PostgreSQL/multi-worker non-goals, and verify the template remains valid JSON and `common.load_config()` accepts both old omitted and new explicit retention configurations.

## 7. Validate the integrated change

- [ ] 7.1 Run the focused database, concurrency, web contention, daily-lock, backup/retention, and restore tests and verify they pass offline with synthetic databases, secrets, and Dropbox fakes only.
- [ ] 7.2 Run `test/run_tests.sh` and verify the complete existing suite remains green, including original task-self-service CAS and assignment-audit scenarios.
- [ ] 7.3 Inspect a throwaway database after the tests to verify WAL, 5-second busy timeout, `synchronous=FULL`, foreign-key enforcement, clean `quick_check`/`foreign_key_check`, bounded dated retention, and successful restore, documenting that no schema migration or database rebuild was introduced.
