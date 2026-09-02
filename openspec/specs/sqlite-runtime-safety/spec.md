# sqlite-runtime-safety Specification

## Purpose

Defines the SQLite operating guarantees that keep one local web process and the daily job reliable and consistent while they share the Raspberry Pi database.

## Requirements

### Requirement: Every database connection uses verified safety settings

The system SHALL enable foreign-key enforcement and a 5-second busy timeout on every application-created SQLite connection. Before either the web application accepts requests or the daily job performs work, the system SHALL verify that foreign-key enforcement is active, the database reports no existing foreign-key violations, and the journal mode is WAL. Startup SHALL fail with an actionable error if these invariants cannot be established or verified.

#### Scenario: Foreign-key enforcement is active

- **WHEN** an application connection attempts to store a row that references a missing parent
- **THEN** SQLite rejects the write with a foreign-key constraint failure

#### Scenario: Existing relationship violation is detected

- **WHEN** startup checks find an existing foreign-key violation
- **THEN** startup fails before serving requests or running the daily workflow
- **AND** the error identifies the failed database integrity check

#### Scenario: Runtime settings cannot be established

- **WHEN** the configured database cannot enter or report WAL mode, or foreign-key enforcement cannot be verified
- **THEN** startup fails rather than continuing with weaker settings

### Requirement: The supported SQLite deployment has bounded contention

The supported deployment SHALL use a database on a local Raspberry Pi filesystem, one threaded Flask web process with one worker, and at most one daily-job process. Concurrent readers and short writers SHALL be allowed to share the database, and a connection encountering a transient lock SHALL wait for up to 5 seconds before reporting failure. The system SHALL keep write transactions short and SHALL NOT hold them open during scraping, notification delivery, Dropbox calls, or other network I/O.

#### Scenario: Web read overlaps a daily write

- **WHEN** the web application reads committed schedule data while the daily job commits a game-plan update
- **THEN** both operations complete without exposing a partially committed update

#### Scenario: Transient writer lock clears within the timeout

- **WHEN** a write encounters another writer and that lock clears within 5 seconds
- **THEN** the waiting write proceeds against the current committed state

#### Scenario: Lock remains beyond the timeout

- **WHEN** database contention remains beyond 5 seconds
- **THEN** the web operation returns a temporary-unavailability response or the command exits unsuccessfully, as appropriate
- **AND** the failure is logged with enough context to identify SQLite lock contention

#### Scenario: Unsupported worker topology is not advertised

- **WHEN** an operator follows the deployment documentation
- **THEN** the documented supported topology specifies exactly one web worker
- **AND** multi-worker operation is identified as unverified and unsupported by this change

### Requirement: Assignment compare-and-swap semantics survive database contention

The system SHALL preserve per-slot compare-and-swap behavior and assignment audit atomicity under the supported concurrent topology. Waiting for a database lock SHALL NOT turn a stale claim or release into an overwrite: after access is obtained, the operation SHALL still apply only if its expected occupant matches current committed state. A successful assignment mutation and its audit entry SHALL commit together, while a refused or failed mutation SHALL commit neither.

#### Scenario: Concurrent claims retain one winner

- **WHEN** two independent connections claim the same empty assignment slot
- **THEN** exactly one claim and its audit entry are committed
- **AND** the other caller receives a conflict containing the current occupant

#### Scenario: Stale release waits behind another writer

- **WHEN** a release waits for a concurrent replacement and its expected occupant is stale when the lock clears
- **THEN** the release is refused as a conflict
- **AND** the replacement and its audit remain unchanged
- **AND** no audit entry is written for the refused release

#### Scenario: Lock timeout during assignment mutation

- **WHEN** an assignment mutation cannot obtain the database lock within 5 seconds
- **THEN** no assignment or audit change from that mutation is committed
- **AND** the failure is reported as temporary unavailability rather than as a successful claim or a stale-state conflict

### Requirement: Daily runs targeting one database do not overlap

The system SHALL acquire an exclusive, process-lifetime daily-run lock derived from the resolved database path before initialization, scraping, synchronization, backup, or notification delivery. A second daily invocation for the same database SHALL fail fast and visibly, while invocations for different database paths SHALL use independent locks. Process termination SHALL release the lock without requiring manual lock-file deletion.

#### Scenario: A second daily run starts

- **WHEN** a daily job is already running for a resolved database path and another invocation targets the same path
- **THEN** the second invocation exits unsuccessfully before any scrape, database mutation, backup, or notification
- **AND** it logs that another daily run owns the lock

#### Scenario: Stale lock-file path remains after a crash

- **WHEN** a previous daily process terminated while its lock file remains on disk but no process holds the operating-system lock
- **THEN** the next daily invocation acquires the lock and runs normally

#### Scenario: Separate databases are targeted

- **WHEN** two daily invocations target different resolved database paths
- **THEN** one invocation's run lock does not block the other
