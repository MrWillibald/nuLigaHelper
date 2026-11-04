# Spec Delta

## MODIFIED Requirements

### Requirement: Removed duties retain their occupants with release-only maintenance

Except for the explicitly authorized one-time migration deletion of current Reinigung position 1 across all dates and categories, an existing assignment outside the game's offered set SHALL be preserved with its identity and history and shown under current privacy rules. It SHALL offer no new claim or replacement candidate. It SHALL continue to enforce one task per person per game, contribute to personal statistics and receive ordinary helper reminders without contributing to required staffing.

Existing authorized release-only maintenance SHALL remain available before the fixed day-before cutoff solely to retire an obsolete duty. After cutoff ordinary release SHALL be refused for every tier, including administrators. No claim, replacement request or takeover SHALL be offered for a removed position. Wischer position 1 SHALL NOT exist after the one-time upgrade; that migration exception SHALL NOT introduce any runtime release or lock override. A permitted pre-cutoff retirement SHALL append a readable audit, close any stale replacement state, remove only the retired field and leave required staffing unchanged. Historical snapshots SHALL remain unchanged. Destructive account and game/date/block maintenance that would cancel a binding commitment SHALL be refused atomically, except removal of block duties of a cancelled home-game date by synchronization; it SHALL NOT authorize ordinary reassignment.

#### Scenario: Migrated youth Kasse retained

- **WHEN** migration renames a youth Unterstützung assignment to Kasse
- **THEN** it remains visible as an existing assignment with the same id, person, game and position
- **AND** no new youth Kasse claim or replacement candidate is offered
- **AND** authorized release-only retirement remains possible before cutoff and ordinary release is refused afterward

#### Scenario: Only the former second Wischer position is occupied

- **WHEN** the upgrade finds only current Reinigung position 1 occupied and position 0 empty
- **THEN** it deletes position 1 and leaves Wischer position 0 empty as a required adult vacancy
- **AND** no compaction, automatic transfer or new assignment occurs
- **AND** the removed helper is no longer blocked from another duty in that game by the deleted assignment

#### Scenario: Stale takeover of the former second position

- **WHEN** a takeover attempts to use a replacement token saved for former Reinigung position 1 before the upgrade
- **THEN** it is refused because the assignment was deleted and the position is not offered
- **AND** no assignment is recreated and no successful takeover audit or notification is produced

#### Scenario: Adult game becomes youth

- **WHEN** a staffed adult game changes to a youth category
- **THEN** Kasse and Wischer occupants remain visible without new candidates
- **AND** required progress and reporting use the youth baseline
- **AND** category change does not silently release their responsibility or bypass binding protection

#### Scenario: Retained duty is locked

- **WHEN** any tier attempts an ordinary release of a retained removed-duty assignment after cutoff
- **THEN** release is refused and the occupant remains responsible
- **AND** the assignment remains visible and excluded from required-duty counts
- **AND** no replacement candidate or new claim is created

#### Scenario: Permitted retirement preserves reporting and history

- **WHEN** an authorized caller explicitly retires a retained removed-duty assignment before cutoff
- **THEN** retirement appends an audit identifying the original duty and person
- **AND** its retained field disappears without creating a required vacancy or changing historical snapshots
