## Purpose

Defines reliable creation, retention, failure reporting, and restoration of SQLite recovery points while the web interface may continue using the live database.

## ADDED Requirements

### Requirement: Backups are transactionally consistent SQLite snapshots

The system SHALL create each upload source with SQLite's online backup API rather than reading or copying the live main database file. The snapshot SHALL contain one committed database state, including committed changes that resided in WAL at snapshot time and excluding uncommitted changes, without requiring the web application to stop. The system SHALL validate the completed snapshot with SQLite integrity and foreign-key checks before any upload and SHALL remove temporary local snapshot artifacts after the attempt.

#### Scenario: Backup overlaps a committed web write in WAL mode

- **WHEN** a web write commits while the daily job creates a snapshot
- **THEN** the snapshot contains a complete committed state from before or after that write
- **AND** it does not contain a partial transaction

#### Scenario: Committed rows reside in WAL

- **WHEN** committed database changes have not been checkpointed into the main database file
- **THEN** the snapshot still contains those committed changes
- **AND** no raw main-file copy or separately uploaded `-wal` or `-shm` file is used as the backup

#### Scenario: Snapshot validation fails

- **WHEN** the snapshot fails its SQLite integrity check or reports a foreign-key violation
- **THEN** no Dropbox latest or dated backup is updated from that snapshot
- **AND** the backup attempt is reported as failed

#### Scenario: Backup cannot finish within its bounded wait

- **WHEN** SQLite remains too busy for the online backup to complete within the configured backup deadline
- **THEN** the attempt terminates unsuccessfully instead of waiting indefinitely
- **AND** the live database remains available to its existing users

### Requirement: Dropbox provides latest and bounded dated recovery points

For each successful snapshot, the system SHALL upload a dated recovery point and then update a stable latest path. The date SHALL use the application's effective date, and another successful run on the same date SHALL replace that date's recovery point. The system SHALL retain the configured positive number of dated recovery points, defaulting to 14, and SHALL prune only older files that match this application's dated-backup naming convention in the configured Dropbox folder. The stable latest object SHALL not count toward the dated limit.

#### Scenario: Successful daily backup

- **WHEN** snapshot creation and validation succeed
- **THEN** Dropbox contains that effective date's recovery point
- **AND** the stable latest path contains the same validated snapshot

#### Scenario: Same effective date runs again

- **WHEN** a second successful backup occurs on the same effective date
- **THEN** that date's recovery point and the stable latest object are replaced
- **AND** no duplicate dated recovery point is added

#### Scenario: Retention limit is exceeded

- **WHEN** more dated recovery points exist than the configured retention count
- **THEN** the oldest matching dated recovery points are deleted until the limit is met
- **AND** unrelated files, the stable latest object, and filenames that do not match the application's convention remain untouched

#### Scenario: Retention count is invalid

- **WHEN** the configured dated-retention count is absent
- **THEN** the system uses 14

#### Scenario: Non-positive retention count is configured

- **WHEN** the configured dated-retention count is zero, negative, or not an integer
- **THEN** the daily job fails configuration validation before creating or uploading a backup

### Requirement: Every backup-stage failure is visible

The system SHALL treat failures during snapshot creation, validation, local snapshot reading, Dropbox client creation, dated upload, latest upload, folder listing, pagination, retention deletion, or local cleanup as backup failures. Each failure SHALL be logged with its stage and error details without exposing credentials, and the daily job SHALL finish with an unsuccessful exit status. A backup failure SHALL NOT be converted into a warning-only success. Independent notification work SHALL still be attempted when the database remains usable, and the unsuccessful backup result SHALL be preserved until job exit.

#### Scenario: Dated upload fails

- **WHEN** Dropbox rejects the dated recovery-point upload
- **THEN** the stable latest object is not updated from that snapshot
- **AND** the failure is logged and the daily job exits unsuccessfully after safe independent work is attempted

#### Scenario: Latest upload fails after dated upload succeeds

- **WHEN** the dated recovery point is stored but updating latest fails
- **THEN** the partial outcome is logged accurately
- **AND** the daily job exits unsuccessfully

#### Scenario: Listing or pruning fails

- **WHEN** Dropbox listing, pagination, or deletion fails during retention enforcement
- **THEN** the already uploaded valid objects are not misreported as absent
- **AND** the retention failure is logged and the daily job exits unsuccessfully

#### Scenario: Failure is not a Dropbox API error

- **WHEN** a filesystem, SQLite, authentication, transport, configuration, or unexpected backup-stage error occurs
- **THEN** it receives the same fail-visible treatment as a Dropbox API error

### Requirement: A documented and verified restore procedure handles WAL safely

The project SHALL document a restore procedure that stops all web and daily database users, preserves the current database for rollback, downloads a selected backup to a temporary path, validates SQLite integrity and foreign keys, prevents stale live `-wal` or `-shm` sidecars from being paired with the restored main file, replaces the database atomically on the same filesystem, and verifies application startup before normal operation resumes. The procedure SHALL state that copying only a live WAL-mode main file is not a valid backup or restore method. An offline automated test SHALL exercise the material validation and replacement steps using synthetic databases.

#### Scenario: Operator restores a validated recovery point

- **WHEN** the operator follows the documented procedure with all database users stopped and a valid selected backup
- **THEN** the restored database passes integrity, foreign-key, and startup checks
- **AND** committed data represented by the selected recovery point is available
- **AND** stale sidecars from the replaced database cannot alter the restored state

#### Scenario: Downloaded recovery point is invalid

- **WHEN** the downloaded file fails integrity or foreign-key validation
- **THEN** the current database is not replaced
- **AND** the procedure directs the operator to retain the current database and choose another recovery point

#### Scenario: Restore rehearsal

- **WHEN** the offline restore test runs against a synthetic WAL-mode database
- **THEN** it restores from a backup snapshot without relying on source sidecars
- **AND** it demonstrates that validation occurs before replacement
