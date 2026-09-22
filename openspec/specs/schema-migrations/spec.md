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
