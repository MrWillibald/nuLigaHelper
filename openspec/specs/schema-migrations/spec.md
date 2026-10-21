# Schema Migrations Specification

## Purpose

Defines controlled database schema versioning and backup-first upgrades so deployed SQLite data can evolve without silent mutation, incompatible startup, or unreviewed data loss.

## Requirements

### Requirement: Every initialized database records its schema revision

The system SHALL record the application schema revision in every newly initialized or adopted database. Normal web and daily-job startup SHALL verify that the recorded revision equals the revision required by the running application and SHALL NOT create, remove, or transform schema objects in an existing database.

#### Scenario: Current database starts

- **WHEN** the web application or daily job opens a database at the required schema revision
- **THEN** schema compatibility verification succeeds before normal work begins

#### Scenario: Database is behind the application

- **WHEN** normal startup opens a recognized database whose schema revision is older than the running application requires
- **THEN** startup fails before serving requests, scraping, synchronizing, backing up, or notifying
- **AND** the error gives the guarded schema-migration command to run while services are stopped

#### Scenario: Database is newer or divergent

- **WHEN** normal startup opens a database with an unknown, newer, or divergent revision
- **THEN** startup fails with a diagnostic
- **AND** no schema or domain data is changed

### Requirement: Fresh initialization is explicit and current

The database initialization command SHALL create the complete current schema only when its target is absent or demonstrably empty, record the current schema revision, seed required system data, and validate integrity before succeeding. It SHALL refuse to treat a populated unversioned database as fresh.

#### Scenario: Empty database is initialized

- **WHEN** an operator runs initialization against an absent or empty target
- **THEN** the current schema and required Supporter team are created
- **AND** the database is recorded at the current schema revision

#### Scenario: Initialization targets existing data

- **WHEN** initialization is run against an unversioned database containing application tables or data
- **THEN** it refuses to create or stamp anything
- **AND** it directs the operator to the schema preflight or migration path

### Requirement: Schema upgrades are guarded offline operations

The system SHALL expose schema upgrades through an explicit command that requires confirmation that all web and daily database users are stopped. Before its first write, the command SHALL perform a read-only source-state preflight and create a validated SQLite backup. It SHALL then apply every pending reviewed revision in order and finish with schema-revision, integrity, foreign-key, and change-specific data checks.

#### Scenario: Recognized database is upgraded

- **WHEN** the operator runs the guarded command against a recognized older revision with stopped services
- **THEN** a validated backup is created and reported
- **AND** all pending revisions are applied in order
- **AND** the upgraded database passes all post-migration checks before success is reported

#### Scenario: Stopped-services confirmation is absent

- **WHEN** the schema-migration command is invoked without the explicit stopped-services confirmation
- **THEN** it refuses before inspecting or changing the target database

#### Scenario: Preflight or backup fails

- **WHEN** the source schema is unsafe to interpret or a valid backup cannot be created
- **THEN** no revision is stamped and no schema or domain data is changed

#### Scenario: Upgrade or postflight fails

- **WHEN** a revision or required post-migration validation fails after the backup was created
- **THEN** the command reports the failed stage and backup path
- **AND** it directs the operator to keep services stopped and restore the validated backup rather than claiming success

### Requirement: Existing unversioned databases are adopted only from an exact baseline

The guarded schema-migration command SHALL adopt an unversioned database only when its schema and required invariants exactly match the reviewed baseline. It SHALL record that baseline without replaying historical creation and then apply later revisions. A database with the legacy game-identity schema SHALL be directed through the existing specialized migration first; every other unknown unversioned shape SHALL be refused.

#### Scenario: Current unversioned database matches the baseline

- **WHEN** preflight finds the exact supported pre-versioning schema and data invariants
- **THEN** the command records the baseline revision
- **AND** applies pending revisions without recreating existing application data

#### Scenario: Legacy game identity is still present

- **WHEN** preflight finds the historical game source-key schema
- **THEN** it makes no change
- **AND** directs the operator to complete the specialized game-identity migration before retrying

#### Scenario: Unversioned schema is not recognized

- **WHEN** preflight finds an unversioned database that differs from every accepted source schema
- **THEN** adoption is refused with diagnostic differences
- **AND** the database remains unversioned and unchanged

### Requirement: The membership migration preserves existing relationships

The first versioned data migration SHALL create exactly one membership for every legacy active or inactive person with a non-null team reference. For a pending registration without an active team, it SHALL convert the legacy desired-team reference into its selected inactive membership. It SHALL create no invented membership when both legacy references are empty and SHALL remove both obsolete single-team fields only after verifying the copied relationships. Existing person IDs, accounts, contacts, assignments, audit records, and MV references SHALL remain attached to the same identities.

#### Scenario: Legacy person has one team

- **WHEN** the membership migration processes a person whose legacy team reference is set
- **THEN** that person has exactly one membership for the referenced team afterward
- **AND** all other relationships retain the same person identity

#### Scenario: Legacy person has no team

- **WHEN** the membership migration processes a person whose legacy active-team and desired-team references are both empty
- **THEN** the person has no memberships afterward
- **AND** the person record is not discarded

#### Scenario: Pending registration has a desired team

- **WHEN** the membership migration processes an unapproved registration
- **THEN** its legacy desired-team reference becomes a selected inactive membership
- **AND** the registration remains hidden and unassignable until admin approval

#### Scenario: Legacy records contain both team references

- **WHEN** a legacy person contains both an active team and a desired-team reference
- **THEN** the migration creates the deterministic union of those referenced memberships
- **AND** creates no duplicate person/team row

#### Scenario: Copied membership cannot be verified

- **WHEN** copied membership counts, foreign keys, MV invariants, or retained relationship checks do not match the preflight plan
- **THEN** the migration does not report success
- **AND** recovery uses the validated pre-migration backup

### Requirement: Migration history is reviewed and drift is detected

Every application schema change SHALL have an ordered, source-controlled migration revision or an explicit determination that no schema revision is needed. Automated verification SHALL detect model changes that are not represented by the migration head, while data transformations and generated operations SHALL remain subject to manual review.

#### Scenario: Model changes without a revision

- **WHEN** automated verification compares the current data model with a database at migration head and finds upgrade operations
- **THEN** verification fails and identifies the unrepresented schema drift

#### Scenario: Generated migration contains a data transformation

- **WHEN** a schema change requires existing domain data to be transformed
- **THEN** the reviewed revision contains an explicit deterministic transformation and validation
- **AND** generated schema operations alone are not accepted as proof of correctness

### Requirement: Birth-date migration preserves existing roster and assignment data

The birth-date schema migration SHALL run through the existing explicit, stopped-services, backup-first upgrade procedure. It SHALL add birth-date storage while preserving person identities, contacts, memberships, account status, MV appointments, assignments, and audit records. Existing people without an independently supplied birth date SHALL retain an unknown date; the migration SHALL NOT infer dates from teams, names, assignment roles, or account state. Missing dates SHALL not by themselves make an otherwise recognized source database ineligible for migration.

#### Scenario: Existing roster is upgraded

- **WHEN** the guarded migration upgrades a recognized database
- **THEN** every existing person retains their identity and relationships
- **AND** their birth date remains unknown unless it was explicitly present in recognized source data

#### Scenario: Existing assignments involve unknown dates

- **WHEN** the migration processes people who already hold assignments and whose dates are unknown
- **THEN** assignments and audit snapshots remain unchanged
- **AND** application eligibility evaluation can report unresolved current/future duties after upgrade

#### Scenario: Birth date could be guessed from a team

- **WHEN** a person's team appears to imply an age group
- **THEN** the migration does not manufacture a date or age from that membership

#### Scenario: Migration verification fails

- **WHEN** revision, integrity, foreign-key, or retained-data checks fail
- **THEN** the migration does not report success
- **AND** recovery follows the existing validated-snapshot procedure

### Requirement: Binding-assignment migration preserves responsibility and history

Durable binding, replacement-request and cutoff state SHALL be introduced only through the existing explicit stopped-services, backup-first schema migration path and a single reviewed migration head. Existing assignment identities, occupants, container identities, slot numbers, relationships and audit snapshots SHALL remain intact. Migration SHALL NOT fabricate historical helper acknowledgment or replacement requests. Existing commitments at or after the fixed 09:00 Europe/Berlin day-before cutoff SHALL be protected on authoritative reads and writes without waiting for a cron marker or allowing a temporary release window. Clock enforcement SHALL NOT fabricate an already-dispatched reminder or historical acknowledgment. Startup SHALL continue to refuse an incompatible revision. Recovery SHALL restore the validated pre-migration backup with writers stopped.

#### Scenario: Existing assignments upgraded

- **WHEN** a recognized database is upgraded
- **THEN** all saved game/block occupants and historical audit snapshots retain their identities and contents
- **AND** no historical confirmation or replacement request is invented

#### Scenario: Due assignment after migration

- **WHEN** a saved assignment is first processed after migration and the fixed 09:00 Europe/Berlin day-before cutoff has already passed
- **THEN** it receives binding protection before any attempted change

#### Scenario: Migration fails verification

- **WHEN** retained-data, revision, integrity or foreign-key verification fails
- **THEN** the migration reports failure and recovery uses the validated pre-migration snapshot

### Requirement: Wischer migration preserves first assignments and removes second assignments

The current per-game Reinigung role SHALL become Wischer only through the explicit stopped-services, backup-first migration workflow and a single reviewed successor revision. Position 0 SHALL retain its assignment id, game, person, slot number, binding lock and replacement state under Wischer. Every existing current Reinigung position 1 assignment SHALL be deleted across all dates, seasons and categories, including past games, due or saved locks and open replacement requests. Its responsibility and replacement state SHALL end with the deletion; its occupant SHALL NOT be moved into position 0. This SHALL be a one-time, explicitly requested exception to binding preservation and SHALL NOT grant a runtime lock override or ordinary release permission.

The migration SHALL preserve every existing audit record and snapshot verbatim and SHALL append one distinguishable system migration-removal entry for each deleted assignment, identifying its original role, slot, game and helper without full birth dates or exact ages. Removal evidence SHALL commit atomically with the deletion and SHALL NOT identify the helper as having voluntarily withdrawn. The rename itself SHALL NOT fabricate confirmations, claims or reminders. The offline migration SHALL send no provider notifications. Other assignments, persons, relationships and day blocks SHALL remain unchanged. Historical Reinigung assignments preceding day blocks SHALL still follow the existing Reinigung-to-Unterstützung-to-Kasse chain and SHALL NOT be deleted or become Wischer.

Preflight SHALL identify the exact rows to rename and delete and refuse unexpected target-role rows, collisions or unsupported source positions before domain-data changes. Verification SHALL permit only the planned first-position renames, exact second-position deletions, corresponding appended audit evidence and revision change, alongside separately reviewed transformations of earlier pending revisions. It SHALL compare every pre-existing audit row unchanged rather than exclude the whole audit table from verification. Repeating the guarded upgrade at the current revision SHALL NOT add removal entries or change unrelated assignments. Normal startup SHALL refuse an older revision without mutation. Recovery SHALL restore the validated backup and matching prior application with all writers stopped.

#### Scenario: Current two-position cleaning assignments are upgraded

- **WHEN** a recognized source contains current Reinigung occupants in positions 0 and 1
- **THEN** position 0 is renamed to Wischer with unchanged identity, occupant, slot, lock and replacement state
- **AND** position 1 is deleted and receives exactly one appended system migration-removal entry
- **AND** existing audits, other assignments, accounts, memberships, dates and day blocks remain unchanged

#### Scenario: Locked or offered second duty is deleted

- **WHEN** current Reinigung position 1 is occupied with a saved or due lock or an open replacement request
- **THEN** the one-time migration deletes the assignment and its request state regardless of cutoff
- **AND** evidence identifies an explicit migration removal rather than voluntary withdrawal
- **AND** stale request and takeover attempts cannot recreate it

#### Scenario: Past and non-adult second assignments are included

- **WHEN** current Reinigung position 1 exists in a past game or as a retained youth or unresolved-category duty
- **THEN** that assignment is deleted and audited by the same one-time migration
- **AND** current Reinigung position 0 and all unrelated obsolete duties are preserved

#### Scenario: Only the second position is occupied

- **WHEN** only the second current Reinigung position has an occupant
- **THEN** it is deleted with removal evidence
- **AND** position 0 stays empty without inventing an assignment or transferring responsibility

#### Scenario: Older Reinigung meant the predecessor of Kasse

- **WHEN** a supported pre-day-block database traverses the complete migration chain
- **THEN** its original Reinigung assignments become Kasse through the reviewed historical revisions
- **AND** those rows are not deleted or turned into Wischer and historical audits remain unchanged

#### Scenario: Unexpected target data is present

- **WHEN** an older source already contains Wischer rows, conflicting role/slot rows or unsupported current Reinigung positions
- **THEN** preflight refuses before domain-data changes
- **AND** it does not broaden the deletion set to make invalid data fit

#### Scenario: Empty current source and fresh initialization

- **WHEN** an empty recognized source is upgraded or a fresh database is explicitly initialized
- **THEN** it reaches the new single head without invented Wischer assignments or removal audits
- **AND** adult games offer one Wischer place through the current staffing policy

#### Scenario: Prior revision is started without migration

- **WHEN** the new web application or daily job opens a database at the preceding revision
- **THEN** startup refuses before serving, syncing or notifying and directs the operator to the guarded migration
- **AND** no automatic rename or assignment deletion occurs

#### Scenario: Upgrade is repeated

- **WHEN** the guarded command runs again on the new current revision
- **THEN** it makes no additional assignment changes or removal audit entries

#### Scenario: Verification failure requires recovery

- **WHEN** revision, integrity, foreign-key or retained-data verification fails
- **THEN** migration reports failure and the validated backup path
- **AND** recovery restores the stopped-writers snapshot, including the deleted second assignments
