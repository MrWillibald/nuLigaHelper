# Schema Migrations Delta

## ADDED Requirements

### Requirement: Visibility-setting migration begins unrestricted and preserves domain records

Durable shared date-visibility storage SHALL be introduced through a single reviewed successor of the current migration head using the existing explicit stopped-services, backup-first workflow. Upgrading a recognized database or explicitly initializing an empty database SHALL yield an unrestricted initial setting. This revision SHALL NOT change any pre-existing game, day block, person, membership, assignment, lock, replacement request or historical audit row, and SHALL NOT send notifications or fabricate assignment evidence. Separately reviewed pending revisions SHALL retain their existing transformations and verification requirements.

Exact source-schema checks SHALL recognize the previously current revision without permitting unreviewed drift, and head validation SHALL include the new setting storage and its valid singleton state. Normal startup with an older revision SHALL fail before serving, syncing or notifying and SHALL NOT create or seed settings implicitly. Repeating migration at the new head SHALL preserve an already configured boundary. Recovery SHALL use the validated pre-migration snapshot and matching prior application with all writers stopped.

#### Scenario: Current deployment is upgraded

- **WHEN** the guarded stopped-services migration upgrades a recognized database from the previously current revision
- **THEN** it reports a validated backup and installs unrestricted visibility storage at the new single head
- **AND** all pre-existing domain rows, assignment state and historical audits remain unchanged
- **AND** postflight verifies both retained data and the new setting state

#### Scenario: Fresh database is initialized

- **WHEN** explicit initialization runs against an absent or empty database
- **THEN** it creates the new current schema and unrestricted setting along with existing required system data

#### Scenario: Old database is started without upgrade

- **WHEN** the new application starts against the preceding schema revision
- **THEN** startup refuses and directs the operator to the guarded migration
- **AND** it does not create setting storage or modify domain rows

#### Scenario: Earlier supported history is upgraded

- **WHEN** a recognized older database traverses earlier pending revisions and the visibility-setting revision
- **THEN** the earlier reviewed transformations run under their own retention rules
- **AND** the visibility-setting revision adds no extra assignment removal, rename, release, lock or audit transition

#### Scenario: Upgrade is repeated after configuration

- **WHEN** an administrator has saved a boundary and the guarded migration is rerun at the new head
- **THEN** the boundary and all other existing data remain unchanged

#### Scenario: Source drift or postflight failure

- **WHEN** source preflight detects an unreviewed schema difference or postflight cannot verify the expected setting and retained data
- **THEN** migration does not report success or silently accept the difference
- **AND** recovery follows the existing stopped-writers validated-snapshot procedure
