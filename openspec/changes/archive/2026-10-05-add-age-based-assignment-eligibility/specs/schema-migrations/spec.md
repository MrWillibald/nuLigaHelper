# Spec Delta

## ADDED Requirements

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
