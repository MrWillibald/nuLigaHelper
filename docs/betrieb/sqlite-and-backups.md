# SQLite production profile, backups and restore

All application database connections enable foreign-key enforcement, wait up to five seconds for transient locks and use `synchronous=FULL`. Startup establishes WAL journal mode and refuses to continue if WAL, the connection settings or `PRAGMA foreign_key_check` cannot be verified. Existing foreign-key violations are reported for explicit operator repair or restore; they are never deleted automatically.

The supported deployment uses the database on a local netcup VPS filesystem, exactly one web process/worker, and at most one daily-job process. Multi-worker WSGI operation, network-filesystem SQLite and PostgreSQL are outside the currently verified topology. WAL creates temporary `<database>-wal` and `<database>-shm` files while connections are active. Do not copy, move or delete these files while the application is running.

The daily job takes a non-blocking OS lock at `<resolved-database>.daily.lock` before database revision verification or external work. A concurrent run for the same database fails immediately with a nonzero exit status; an unlocked file left after a restart is harmless. `run.sh` uses `exec`, so cron receives the Python process status directly.

Backups use SQLite's online backup API and therefore include one complete committed state even when committed pages still reside in WAL. The live `.db` file and its sidecars are never uploaded directly. Every snapshot must pass `PRAGMA quick_check` and `PRAGMA foreign_key_check` before publication. Dropbox receives, in this order:

1. `<database-stem>-YYYY-MM-DD<suffix>` for `common.effective_today()`;
2. the stable configured database basename as the latest backup;
3. retention cleanup of only older names that exactly match that database's dated convention.

`NULIGAHELPER_DROPBOX_DATED_RETENTION` is required in production and must be a positive integer. A same-day rerun replaces that day's recovery point. Retry-safe Dropbox operations recover from transient connection, DNS, timeout, rate-limit and server failures with three bounded retries after 1, 2 and 4 seconds; authentication, bad-input and other permanent failures fail immediately. Snapshot, validation, read, Dropbox client, upload, listing, pagination, deletion and local-cleanup failures are logged concisely by stage and make the daily job exit unsuccessfully. When the source database remains usable, due notification work is still attempted after a backup failure, but notifications are not retried automatically.

## Restore a downloaded snapshot

Restoration is an offline operation. Copying only a live WAL-mode main file is not a valid backup or restore method.

1. Stop the web process and disable or wait for the daily job. Verify that no process is using the database or holding its daily-run lock.
2. Download the selected latest or dated Dropbox object to a separate path. Never download it over the active database.
3. Run the guarded command against the configured target:

   ```bash
   python3 -m nuligahelper.cli --db /absolute/path/nuliga_helper.db \
       restore-snapshot /path/to/downloaded-snapshot.db --confirm-stopped
   ```

4. The command validates the candidate before changing the target, fsyncs a same-filesystem temporary copy, preserves the current `.db`, `-wal` and `-shm` set in the printed rollback directory, and atomically installs the snapshot. If installation fails, it attempts to recover the complete preserved set. Keep all services stopped and inspect the reported paths after any failure.
5. Start exactly one application process. Normal startup re-establishes WAL and repeats the runtime and foreign-key checks. Verify the schedule, authentication and one non-destructive database query before deliberately re-enabling normal traffic and the daily/cleanup timers.
6. Retain the command-created rollback directory until the restored service and the next validated Dropbox backup have been accepted.

Before deploying these settings to an existing database, stop writers, preserve the current database together with any sidecars, run the full offline suite, and start one process to perform the integrity preflight. For rollback, stop every database user first and restore a validated snapshot or the complete preserved database set; never replace only the main file while WAL users are active.


---
Back to the [README](../../README.md) · [Operations index](README.md)
