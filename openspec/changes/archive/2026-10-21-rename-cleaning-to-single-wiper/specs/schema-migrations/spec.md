# Spec Delta

## ADDED Requirements

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
