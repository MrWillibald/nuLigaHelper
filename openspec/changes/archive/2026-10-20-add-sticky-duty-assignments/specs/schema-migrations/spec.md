# Spec Delta


## ADDED Requirements

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
