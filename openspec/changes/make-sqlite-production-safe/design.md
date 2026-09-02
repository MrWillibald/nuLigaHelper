## Context

See `proposal.md` for motivation and the two delta specs for behavior. Today `db.make_engine()` creates a default SQLAlchemy SQLite engine, `init_db()` creates tables without connection or integrity checks, Flask keeps one SQLAlchemy session per request, and the daily job uses a session across synchronization and notification reads. `main.backup_to_dropbox()` opens the configured live `.db` path as ordinary bytes and suppresses only Dropbox API errors. Both launch scripts can run independently, and nothing prevents a second daily process.

The target is a Raspberry Pi with the database and lock/snapshot files on a local filesystem. SQLite is embedded: the Flask process and daily process each own connection pools, so all safety settings must be applied per connection and all cross-process coordination must use SQLite or OS primitives rather than Python-only locks. Existing databases may already contain foreign-key violations because SQLite enforcement has not been enabled. The schema itself is unchanged, and project policy deliberately has no migration layer.

## Goals / Non-Goals

**Goals:**

- Define one centralized, testable SQLite engine profile shared by web, daily, CLI, and tests.
- Improve read/write coexistence while favoring integrity and bounded failure over unlimited retries.
- Preserve the externally specified assignment CAS winner/conflict behavior and atomic audit writes under WAL-specific contention.
- Produce self-contained, validated snapshots whose correctness does not depend on checkpoint timing or copying WAL sidecars.
- Give cron an unambiguous nonzero result for overlap and every backup-stage failure without unnecessarily suppressing independent notifications.
- Make restore a rehearsable operation rather than an undocumented file replacement.

**Non-Goals:**

- Detecting or supporting network filesystems; operators must place the database on a local filesystem.
- Proving safe operation with multiple WSGI worker processes. One process/worker may remain threaded.
- Changing tables, adding schema-version tables, repairing inconsistent production data automatically, or introducing migration tooling.
- Replacing Dropbox, adding replication, or selecting/hardening a production WSGI/TLS stack.
- Guaranteeing that a backup and an external notification form one atomic operation; external side effects cannot participate in the SQLite transaction.

## Decisions

### 1. Centralize connection settings and fail startup when invariants are not met

`db.make_engine()` will become the only application/test engine construction path. It will set the sqlite3 connect timeout to 5 seconds and install a SQLAlchemy connect hook that applies and verifies `PRAGMA foreign_keys=ON`, `PRAGMA busy_timeout=5000`, and `PRAGMA synchronous=FULL` for every DB-API connection. `FULL` is chosen because the Pi can lose power and this database is small enough that durability is more important than maximizing write throughput.

Database initialization will establish `PRAGMA journal_mode=WAL` on a dedicated connection and verify the returned mode before normal work. Because journal mode is persistent database metadata rather than a table/schema change, this requires no migration. Initialization will then run `PRAGMA foreign_key_check`; any rows are an actionable startup failure that identifies the affected table/row references but does not attempt repair. The web process must finish this before accepting requests, and the daily job before scraping or other side effects.

Tests and management commands must use this factory instead of constructing unconfigured engines. A small inspection helper will make startup assertions and test expectations share one source of truth.

Alternatives considered:

- **Keep rollback-journal mode:** simpler sidecar behavior, but readers and the daily writer block each other more often. Rejected for the shared UI/job workload.
- **Set pragmas once:** `foreign_keys`, `busy_timeout`, and `synchronous` are connection-scoped. Rejected because each process/pool opens new connections.
- **Use `synchronous=NORMAL`:** faster and still consistent after power loss, but recently committed WAL transactions can be lost. Rejected in favor of `FULL` durability for this low-volume application.
- **Silently continue if WAL or integrity checks fail:** rejected because it creates an unobservable weaker production mode.

### 2. Use WAL without treating the main file as the whole live database

WAL allows web readers to continue while the daily job commits, but only one writer exists at a time. The 5-second busy timeout covers ordinary `SQLITE_BUSY`; it does not make every stale WAL read transaction safely upgradable, because SQLite can return `SQLITE_BUSY_SNAPSHOT`. Sessions must therefore end transactions promptly, and no write transaction may span scraping, mail/SMS delivery, Dropbox calls, or other network I/O. Web request sessions continue to close at teardown; the daily sync commits before backup and notification I/O.

Normal SQLite automatic checkpoints remain enabled. The backup path will not force `wal_checkpoint(TRUNCATE)`: a forced checkpoint is unnecessary for online-backup correctness and can add contention. Operators should expect transient `<db>-wal` and `<db>-shm` files while users are connected. Monitoring unexpectedly persistent WAL growth is useful operationally, but no new monitoring service is added.

Alternatives considered:

- **Copy `.db`, `-wal`, and `-shm` together:** coordinating a correct live three-file copy is fragile and not transactionally atomic. Rejected in favor of SQLite's backup API.
- **Checkpoint then copy the main file:** checkpoint completion can be blocked by readers and a writer can commit between checkpoint and copy. Rejected.
- **Delete sidecars during normal operation:** unsafe while connections are live. Sidecars are handled only by the stopped-system restore procedure.

### 3. Revalidate assignment CAS after WAL contention; never blindly replay stale intent

The unique constraints on `(game_id, role, slot)` and `(game_id, person_id)`, conditional release predicate, and append-only audit stay authoritative. Assignment mutation handling will distinguish three outcomes:

1. the mutation and audit commit together;
2. current state no longer matches the expected occupant, yielding `SlotConflictError` with the fresh occupant and no audit;
3. the bounded contention deadline expires, yielding temporary unavailability and no committed mutation/audit.

If a write encounters `SQLITE_BUSY_SNAPSHOT` or a retryable lock, the transaction must roll back, open a fresh transaction, reload the slot and relevant person's assignments, and compare them with the original expectation before any bounded retry. It may retry only while the expectation still matches and the common 5-second deadline remains. It must return a stale conflict immediately when fresh state differs. Integrity errors are translated only after a fresh read identifies whether the slot or one-task-per-game constraint won; unrelated integrity failures remain errors. This prevents a generic retry wrapper from replaying stale releases or manufacturing audit entries.

Web lock exhaustion will roll back and return HTTP 503 (including the normal JSON error shape for JSON endpoints); command/job paths log context and exit nonzero. Conflict remains HTTP 409 and is not conflated with lock timeout.

Alternatives considered:

- **Retry the whole request automatically:** rejected because it can replay stale user intent and external effects.
- **Use a process-local mutex:** cannot coordinate Flask and daily-job processes.
- **Serialize all writes with `BEGIN IMMEDIATE`:** viable but broad; it would require carefully restructuring request authentication/read transactions and can increase writer hold time. Targeted fresh-transaction CAS retry preserves current boundaries with less lock amplification.

### 4. Prevent daily overlap with an OS advisory lock keyed by canonical database path

At the very start of `main.main()`, after resolving configuration/database path but before initialization, scraping, or notifications, the job will open `<resolved-db-path>.daily.lock` and take a non-blocking exclusive `fcntl.flock`. The descriptor remains open for the full workflow. A held lock produces a clear log/console message and a distinct nonzero exit; a leftover file without a live lock is harmless. Canonical resolved database paths produce independent locks for independent databases.

The guard belongs in Python rather than only `run.sh`, so `python main.py`, tests, and future schedulers cannot bypass it accidentally. `run.sh` remains a thin launcher and documents cron behavior. `fcntl` is appropriate because the supported platform is Linux/Raspberry Pi and the database must be local.

Alternatives considered:

- **Cron configuration alone:** easy to bypass and cannot protect manual invocations.
- **PID-file existence:** stale files after crashes require unsafe cleanup heuristics.
- **Database lock row:** would add schema and migration concerns and would hold or repeatedly update the same database it protects.

### 5. Build a bounded online snapshot, then validate it before upload

A backup helper will open the live database through sqlite3 with the same bounded busy policy and copy it to a uniquely named temporary `.db` in the live database directory using `Connection.backup()`. Keeping the destination on the same filesystem supports later atomic restore operations and predictable permissions. A monotonic 30-second overall backup deadline, enforced through backup progress/retry handling, prevents pathological contention from hanging cron forever; it is separate from the per-lock 5-second timeout.

The online backup API selects a transactionally consistent committed state and reads committed pages from WAL as needed; no source checkpoint is required. Once both connections close, the temporary destination is self-contained. Before upload, a read-only validation connection requires `PRAGMA quick_check` to return exactly `ok` and `PRAGMA foreign_key_check` to return no rows. Snapshot files and any temporary sidecars are cleaned in `finally`; cleanup errors are recorded as backup failures without hiding an earlier primary error.

Alternatives considered:

- **Read the live file into memory:** incorrect in WAL mode and can capture changing pages.
- **Use `VACUUM INTO`:** can produce a consistent copy, but the SQLite backup API directly matches the requested online-copy semantics and gives progress/busy control without also rebuilding the file.
- **Keep snapshots indefinitely on local disk:** unnecessary on constrained Pi storage; upload attempts clean their temporary artifacts.

### 6. Publish dated first, latest second, then enforce a narrow retention policy

The existing stable Dropbox object keeps the configured database basename (for example `nuliga_helper.db`). Dated objects use `<stem>-YYYY-MM-DD<suffix>` (for example `nuliga_helper-2026-09-02.db`) based on `common.effective_today()`, satisfying the project's date abstraction; same-day reruns overwrite that date. `club.dropbox.dated_retention` is an optional positive integer with default 14.

For a validated snapshot, upload order is:

1. overwrite the dated object;
2. overwrite latest with the same bytes;
3. list the configured folder through all pagination pages;
4. parse only exact filenames generated from this database basename and delete oldest matching dated objects beyond the limit.

Uploading dated first ensures a failed first upload cannot replace latest. If latest fails, the dated recovery point remains useful and the partial outcome is reported. If retention fails, valid uploads remain, but the job still fails because the bound was not enforced. Matching is anchored to the configured basename and ISO date; unrelated Dropbox content is never pruned. The stable latest object is excluded from the dated count.

The helper will expose typed/staged failures rather than swallowing exceptions. Dropbox credentials and snapshot bytes never enter logs.

Alternatives considered:

- **Only latest:** gives no rollback window if corruption is uploaded successfully.
- **Timestamp every run:** creates churn when a job is rerun on one day; one point per effective date is adequate for a daily cron.
- **Delete before upload:** risks reducing recovery coverage when the new upload later fails.
- **Prune the whole folder by age:** could delete unrelated club files.

### 7. Preserve backup failure until process exit while attempting safe notifications

Backup remains after the synchronization commit so its snapshot includes the new game state. `main()` will catch a backup-domain failure at the orchestration boundary, log it with traceback/stage, retain it as the final unsuccessful result, and continue notification queries/delivery because those are independent and the source database has already passed startup and remains usable. After safe independent work completes, the job raises/returns failure so `run.sh` and cron see nonzero. A later notification failure must not erase the already logged backup failure; all encountered failures should be logged before exit.

No notification is retried as part of backup handling, avoiding duplicate mail/SMS side effects.

Alternatives considered:

- **Raise immediately at the backup call:** visible but skips due notifications for an unrelated Dropbox outage.
- **Log warning and return success:** current behavior; rejected because cron cannot alert on lost recovery coverage.

### 8. Restore is an offline, validate-before-replace operation

`manage_db.py restore-snapshot SNAPSHOT --confirm-stopped` will provide a guarded, reusable validate-before-replace operation against the resolved configured database. The explicit flag makes the operator acknowledge that the command cannot prove all Flask/daily connections are stopped. The command will validate the candidate first, preserve the current `.db`, `-wal`, and `-shm` files together in a uniquely named rollback directory, copy the candidate to a temporary file on the target filesystem, fsync it, atomically install it, and recover the preserved set if installation fails. It will never interpret or transform application rows and therefore is recovery tooling, not migration machinery.

`README.MD` (and focused test documentation if useful) will provide a concrete checklist:

1. stop the web process and disable/wait for the daily job; verify no process holds the run lock or database;
2. download the chosen latest/dated object without overwriting the active database;
3. run the guarded restore command, which validates the candidate before preserving and replacing the active database and sidecars;
4. start one application process so normal startup re-verifies WAL/foreign keys, perform a smoke check, then resume cron/web operation;
5. retain the command-created rollback set until the restored service is accepted.

An offline automated test will create a synthetic WAL database with committed WAL-resident data, make an online snapshot, restore from only that snapshot into a clean target, validate before replacement, and verify the restored data. Negative cases prove that missing confirmation or invalid input leaves the current target and sidecars untouched. This is operational recovery tooling/documentation, not schema migration machinery.

Alternatives considered:

- **Replace the main file while Flask is running:** unsafe because existing connections and sidecars can continue referring to the old database state.
- **Delete old sidecars before preserving rollback:** loses potentially necessary state from the current database.

## Risks / Trade-offs

- **[Existing databases contain orphaned rows]** → Foreign-key preflight will stop startup. Log exact `foreign_key_check` findings and require explicit operator repair or restore; do not auto-migrate or auto-delete data.
- **[WAL cannot operate on the chosen filesystem]** → Fail startup and document the local-filesystem requirement rather than silently falling back to rollback journal.
- **[Writer contention still exists]** → Keep writes short, wait only 5 seconds, revalidate CAS after snapshot contention, return 503/nonzero after the deadline, and test real independent connections.
- **[Long readers delay checkpoints and grow WAL]** → Close request/job transactions promptly, leave automatic checkpointing enabled, and avoid network I/O inside transactions. Backup correctness does not depend on checkpoint completion.
- **[Power loss after SQLite commit]** → Use `synchronous=FULL`; WAL provides consistency but is not a substitute for storage/media health or backups.
- **[Power/network loss creates partial Dropbox publication]** → Dated-first ordering preserves any successful recovery point, logs exact stage outcomes, and never reports the whole backup as successful after a partial failure.
- **[Retention parsing deletes the wrong object]** → Match only an anchored basename-plus-ISO-date convention and test unrelated names and paginated listings.
- **[Snapshot cleanup itself fails]** → Log cleanup separately and keep the job unsuccessful; unique temporary names prevent the next run from treating leftovers as valid backups.
- **[One-worker limit constrains future scaling]** → State the supported boundary clearly. Multi-worker validation or PostgreSQL is a separate change.
- **[Backup succeeds but later notifications fail]** → Preserve each outcome independently in logs and return nonzero for any fatal stage; external notification side effects cannot be rolled back.

## Migration Plan

1. Before deployment, stop the web app and daily cron and preserve the current rollback-journal database as an offline rollback copy.
2. Run the full offline suite, including pragma, concurrent-CAS, lock, backup, retention, and restore tests.
3. Deploy code/config/documentation with `dated_retention` omitted to accept the default 14, or set an explicit positive count.
4. Start exactly one web process/worker. Startup changes the existing database's persistent journal mode to WAL, verifies foreign keys, and creates runtime sidecars; it does not alter tables or require deleting the database.
5. If foreign-key violations are reported, keep services stopped and explicitly repair or restore data under owner control. Do not add an automatic migration.
6. Run one manual daily job, confirm latest plus dated Dropbox objects, verify pruning scope, inspect logs/exit status, and rehearse `manage_db.py restore-snapshot` against a throwaway configured database before re-enabling cron.
7. For rollback, stop all database users first, preserve the `.db` plus any `-wal`/`-shm` files, checkpoint through SQLite if possible, and use the guarded restore command with the pre-deployment copy or a validated online-backup snapshot before reverting code. Never revert by copying only a live WAL main file.
