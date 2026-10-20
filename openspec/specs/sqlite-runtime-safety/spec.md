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

The system SHALL preserve per-slot compare-and-swap behavior and assignment/request/lock audit atomicity under the supported concurrent topology. Waiting for a database lock SHALL NOT turn a stale claim, replacement-request change or takeover into an overwrite: after access is obtained, the operation SHALL still apply only if its expected occupant and applicable request state match current committed state. Authorization, active-person status, offeredness/capacity, one-duty rules, current eligibility including the resulting Verkauf group, and the day-before boundary SHALL be revalidated under the same write serialization and after retries. The fixed boundary SHALL be 09:00 Europe/Berlin on the day before the duty, regardless of reminder-job or delivery success. Every assignment write SHALL recheck the aware local clock and saved lock state after obtaining access and after retries, persist due cohort marker/lock/request-closure transitions idempotently under serialization, and refuse post-deadline replacement even if no batch ran. Due reminder processing SHALL commit those transitions before dispatch without advancing the deadline on an early invocation. A successful assignment, request, takeover or lock transition and its audit entries SHALL commit together, while a refused or failed requested mutation SHALL commit neither its state change nor an audit claiming success. A takeover SHALL transfer responsibility without a committed intermediate vacancy. Any separately committed reminder-batch lock transition SHALL remain distinguishable from a refused takeover. A successful claim of an empty duty after the boundary SHALL commit its occupied and locked state together. Network delivery SHALL NOT occur while this serialization is held.

#### Scenario: Concurrent claims retain one winner

- **WHEN** two independent connections claim the same empty assignment slot
- **THEN** exactly one claim and its audit entry are committed
- **AND** the other caller receives a conflict containing the current occupant

#### Scenario: Stale takeover waits behind another writer

- **WHEN** a takeover waits for another successful takeover and its expected occupant or request state is stale when the lock clears
- **THEN** the waiting takeover is refused as a conflict
- **AND** the winner and its audit remain unchanged
- **AND** no successful-transfer audit is written for the refused takeover

#### Scenario: Stale release waits behind another writer

- **WHEN** a legacy ordinary-release request waits behind a successful replacement and then obtains write access with a stale expected occupant
- **THEN** it cannot overwrite or release the saved assignment
- **AND** the replacement and its audit remain unchanged
- **AND** neither stale state nor waiting for a writer creates an ordinary-release permission

#### Scenario: Lock timeout during assignment mutation

- **WHEN** an assignment mutation cannot obtain the database lock within 5 seconds
- **THEN** no assignment or audit change from that mutation is committed
- **AND** the failure is reported as temporary unavailability rather than as a successful claim or a stale-state conflict

#### Scenario: Two replacements contend for one request

- **WHEN** two otherwise valid takeovers contend for the same occupied slot and replacement request
- **THEN** exactly one responsibility transfer and its audit commit
- **AND** the other request receives the current authoritative state without changing the winner

#### Scenario: Cutoff batch wins while a takeover waits

- **WHEN** a takeover waits until 09:00 Europe/Berlin on the day before the duty or until a due cutoff transition has committed before it obtains write access
- **THEN** the saved cutoff is revalidated and the takeover is refused
- **AND** the original helper remains responsible and the committed lock has closed the request

#### Scenario: Late empty-slot claim races another claim

- **WHEN** two callers claim the same truly empty offered position after the day-before boundary
- **THEN** exactly one claim and its immediately locked state commit with their audit
- **AND** the loser receives a conflict without being able to replace the winner

#### Scenario: Cake configuration overlaps a takeover

- **WHEN** a cake configuration update and a takeover contend for the same block
- **THEN** both use current saved capacity and binding protection under shared serialization
- **AND** neither can silently remove an occupied position or transfer responsibility into a removed position

#### Scenario: Takeover commits before the cutoff batch

- **WHEN** a valid takeover commits before the fixed 09:00 Europe/Berlin day-before cutoff
- **THEN** the batch locks the new responsible helper and closes any remaining request against current authoritative state
- **AND** reminders are dispatched only after the cutoff transaction commits

#### Scenario: Dispatch fails after the cutoff commit

- **WHEN** the dated cutoff and all occupied-duty locks have committed and mail or SMS dispatch then fails or is skipped
- **THEN** the saved cutoff and locks remain durable
- **AND** contactless helpers, process restarts and retries cannot reopen the assignments

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
