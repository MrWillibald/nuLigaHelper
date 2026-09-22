## Purpose

Defines how an operator promotes reviewed GitHub revisions to the production server while preserving the existing SQLite data, controlling scheduled work, and retaining a verifiable recovery path.

## ADDED Requirements

### Requirement: Production releases come from protected master by operator action

The production deployment workflow SHALL accept only an exact commit that is reachable from the fetched protected `master` branch. Fetching or merging code SHALL NOT by itself activate a release; activation SHALL require an operator command. The workflow SHALL record the selected commit and the previously active commit without relying on a mutable branch name after selection.

#### Scenario: Operator selects an eligible release

- **WHEN** the operator requests deployment after a reviewed commit has reached `master`
- **THEN** the workflow pins that commit and identifies it in the preparation and activation record
- **AND** a later movement of `master` does not change the selected release

#### Scenario: Commit is outside master

- **WHEN** the operator supplies a commit that is not reachable from the fetched `master`
- **THEN** deployment refuses it before changing the active application or database

#### Scenario: GitHub receives a new commit

- **WHEN** `master` advances without an operator deployment command
- **THEN** the production application continues running its previously activated commit

### Requirement: Candidate preparation leaves production state unchanged

The workflow SHALL prepare each candidate in a separate release location from the exact selected source revision and a source-controlled production dependency definition. It SHALL validate the candidate before service interruption. Preparation SHALL NOT modify the active release, production database, secrets, or site configuration. A failed preparation SHALL leave the current service and timers unaffected.

#### Scenario: Candidate passes preparation

- **WHEN** source retrieval, dependency installation, and required offline validation succeed
- **THEN** the candidate is available for an explicit activation step
- **AND** the currently running application and its data remain unchanged

#### Scenario: Candidate preparation fails

- **WHEN** retrieval, dependency installation, or validation fails
- **THEN** the workflow reports the failed stage and does not interrupt production services

### Requirement: Deployment preserves the live SQLite database and private configuration

Application releases SHALL be separate from the existing production SQLite database at `/var/lib/nuligahelper/nuliga_helper.db` and from private environment and configuration files. Deployment SHALL preserve the database, its associated live SQLite state, account records, assignments, and audits; it SHALL NOT initialize a replacement database or copy only a live WAL-mode main file. Reusable release assets SHALL contain no production credentials or personal data.

#### Scenario: New application release is prepared

- **WHEN** a candidate release is installed
- **THEN** its source and dependencies are placed outside the database and private configuration paths
- **AND** the existing database and private settings remain the inputs used by the activated services

#### Scenario: Existing database is absent or unexpected

- **WHEN** the configured production database is missing, empty, or differs from the approved target
- **THEN** activation stops without creating or adopting a replacement database

### Requirement: All database writers are quiescent for cutover

Before changing the active release or running a schema migration, the deployment workflow SHALL prevent new scheduled daily and cleanup runs, wait for or refuse an active job rather than interrupting it silently, remove public write traffic, and stop the web database user. It SHALL verify that no known production database writer remains active. Timer restoration SHALL account for systemd persistent catch-up behavior and SHALL not cause an unreviewed immediate notification or cleanup run.

#### Scenario: Scheduled job is already running

- **WHEN** the operator starts activation while the daily or cleanup service is active
- **THEN** activation waits under a documented bound or refuses with an actionable status
- **AND** it does not terminate the job mid-notification or mid-backup without an explicit separate operator decision

#### Scenario: Maintenance spans a timer firing time

- **WHEN** a persistent timer could fire immediately upon resumption
- **THEN** the workflow identifies the pending run and requires an explicit operator choice before resuming that timer

### Requirement: Cutover has a validated recovery point and explicit schema gate

After writers are quiescent and before activation or migration, the workflow SHALL create and validate a complete SQLite recovery snapshot and record its protected location. It SHALL inspect the database revision against the candidate release. A recognized older revision SHALL require the existing guarded migration with stopped-services confirmation; an unknown, divergent, corrupt, or newer revision SHALL block activation. Normal deployment SHALL NOT run a schema migration implicitly.

#### Scenario: Database already matches the candidate

- **WHEN** the validated database is at the candidate's required schema revision
- **THEN** activation may continue without changing its schema or domain rows

#### Scenario: Recognized schema upgrade is needed

- **WHEN** the database is behind a candidate with a reviewed migration path
- **THEN** activation pauses until the operator explicitly runs the guarded migration while all writers are stopped
- **AND** the migration and its postflight must succeed before activation continues

#### Scenario: Recovery point or schema preflight fails

- **WHEN** snapshot validation fails or the source schema is unsafe to interpret
- **THEN** activation stops before switching the active release or reopening traffic

### Requirement: Activation verifies service readiness before resuming normal work

The workflow SHALL switch the application entry path to the prepared candidate as one recoverable operation, start the web service, and verify local readiness and the configured public HTTPS endpoint before accepting the release. It SHALL verify that every application service references the same selected release and configured database. Scheduled work SHALL resume only after the operator accepts the health and timer checks. A deployment record SHALL include the selected commit, previous commit, database revision, recovery snapshot path, checks, and outcome without secret values.

#### Scenario: Candidate passes cutover checks

- **WHEN** the selected release starts and its local and public checks succeed
- **THEN** the workflow reports it as active with the verified commit and database revision
- **AND** the operator can resume the daily and cleanup timers under the documented catch-up decision

#### Scenario: Candidate fails before public reopening

- **WHEN** the candidate cannot start or fails readiness while public writes remain blocked
- **THEN** the workflow keeps scheduled work paused and offers the previous release and recovery snapshot for the documented rollback path
- **AND** it does not claim deployment success

### Requirement: Rollback respects database compatibility and accepted writes

The workflow SHALL retain the previously active release until the new release is accepted. It SHALL permit a code-only rollback when the previous release can use the current database revision. If a migration or accepted production writes make that compatibility uncertain, it SHALL refuse an automatic code-only rollback and require an explicit stopped-writer recovery decision. Database restoration SHALL use the existing validated restore procedure and account for data and privacy obligations after the snapshot time.

#### Scenario: Failed release has not changed the database schema

- **WHEN** cutover fails before public writes resume and the previous release accepts the current database revision
- **THEN** the operator can reactivate the previous release without replacing the database

#### Scenario: Previous release is incompatible with migrated data

- **WHEN** the database has moved to a revision the previous release cannot use
- **THEN** code-only rollback is refused
- **AND** recovery requires a deliberate validated database restore or another reviewed compatible release

#### Scenario: Rollback is considered after public writes

- **WHEN** the new release has accepted production changes
- **THEN** the workflow does not automatically restore an older snapshot and erase those changes
