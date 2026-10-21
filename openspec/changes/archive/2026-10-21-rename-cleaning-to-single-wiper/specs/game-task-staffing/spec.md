# Spec Delta

## MODIFIED Requirements

### Requirement: Games expose their offered task position set

Adult games SHALL offer one Zeitnehmer, one Sekretär, two Verkauf, exactly one independent Ordnungsdienst, one Kasse and exactly one Wischer position (position 0). Youth games and unrecognized categories SHALL offer only Zeitnehmer, Sekretär, two Verkauf and exactly one Ordnungsdienst. Numbers SHALL identify positions of repeated semantic roles and the original position of a retained removed duty. Kasse SHALL additionally perform the ordering/security duties planned for a second Ordner without creating another assignment or relaxing one task per person per game. Per-game Wischer SHALL remain distinct from day-wide Aufräumen.

#### Scenario: Adult expanded game duties

- **WHEN** an adult game's assignments are opened
- **THEN** seven offered positions are visible or assignable according to existing rights
- **AND** Wischer is shown as one unnumbered offered position and Ordnungsdienst has exactly one

#### Scenario: Youth offered duties

- **WHEN** a youth game including SPF is opened
- **THEN** only the baseline five positions are offered
- **AND** empty Kasse and Wischer fields and optional markers are absent
- **AND** new Kasse and Wischer claims are refused for every tier and CLI/system caller

#### Scenario: Independent repeated duty positions

- **WHEN** the guarded upgrade processes occupants of current Reinigung positions 0 and 1
- **THEN** position 0 retains its assignment identity, occupant and binding state under Wischer
- **AND** position 1 is deleted with an appended system migration-removal audit, even when locked or seeking a replacement
- **AND** only the single Wischer position remains visible, without moving the second occupant to position 0

#### Scenario: Second position is empty

- **WHEN** an adult game has no assignment in Wischer position 1
- **THEN** that position is absent from the schedule and candidate controls
- **AND** every tier and CLI/system caller is refused a new position 1 claim

#### Scenario: Current name is consistent

- **WHEN** a current Wischer duty is presented in the schedule, confirmation, candidate response, CLI, personal statistics, new audit transition or ordinary notification
- **THEN** its semantic role is Wischer and no current role is offered as Reinigung
- **AND** historical audit labels and snapshots retain their recorded text

#### Scenario: Stale old role request

- **WHEN** a caller submits an assignment mutation using the former current role name Reinigung
- **THEN** it is refused without an assignment change, successful mutation audit or notification

### Requirement: Offered positions are required by category

Every offered position SHALL be required. Adult M/F classes SHALL have seven required positions. Youth mA–mE/wA–wE and SPF classes SHALL have five. GE, missing, unsupported and conflicting labels SHALL retain the five-position baseline and expose unresolved classification. The shared explicit classifier SHALL be used with the existing age rules. New claims and takeovers SHALL revalidate offeredness against saved game state under write serialization and after retries. A replacement request SHALL NOT turn an occupied position into a free claim target or authorize a takeover of a position outside the offered set.

#### Scenario: Adult additional gaps

- **WHEN** an adult game's baseline five positions are occupied and its added duties are empty
- **THEN** two required positions remain missing: Kasse and Wischer

#### Scenario: Youth complete baseline

- **WHEN** a youth game's baseline five positions are occupied
- **THEN** there are no required vacancies
- **AND** removed Kasse and Wischer duties do not generate gaps

#### Scenario: Youth Ordnungsdienst is required

- **WHEN** a youth game's Ordnungsdienst is empty
- **THEN** that required position remains missing even if a retained removed duty is occupied

#### Scenario: Classification is unresolved

- **WHEN** the class is GE or cannot be recognized
- **THEN** no adult-only duties are offered or required
- **AND** the baseline five and unresolved-classification status are retained

#### Scenario: Saved category changed before claim

- **WHEN** a caller holds stale adult game state and the saved game has become youth
- **THEN** a new Kasse or Wischer claim is refused using the saved category
- **AND** no assignment or claim audit is written

#### Scenario: Saved category changed before takeover

- **WHEN** an adult-only position has an open replacement request but the saved game has become youth before takeover
- **THEN** the takeover is refused because that position is no longer offered
- **AND** the original assignment remains visible with its history and without changing youth required progress

### Requirement: Removed duties retain their occupants with release-only maintenance

Except for the explicitly authorized one-time migration deletion of current Reinigung position 1 across all dates and categories, an existing assignment outside the game's offered set SHALL be preserved with its identity and history and shown under current privacy rules. It SHALL offer no new claim or replacement candidate. It SHALL continue to enforce one task per person per game, contribute to personal statistics and receive ordinary helper reminders without contributing to required staffing.

Existing authorized release-only maintenance SHALL remain available before the fixed day-before cutoff solely to retire an obsolete duty. After cutoff ordinary release SHALL be refused for every tier, including administrators. No claim, replacement request or takeover SHALL be offered for a removed position. Wischer position 1 SHALL NOT exist after the one-time upgrade; that migration exception SHALL NOT introduce any runtime release or lock override. A permitted pre-cutoff retirement SHALL append a readable audit, close any stale replacement state, remove only the retired field and leave required staffing unchanged. Historical snapshots SHALL remain unchanged. Destructive account and game/date/block maintenance that would cancel a binding commitment SHALL be refused atomically; it SHALL NOT authorize ordinary reassignment.

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

### Requirement: Staffing reports share required positions

Progress, open-duty statistics, completeness and MV missing-task reminders SHALL use the same offered/required positions and retain independent age-deficiency reporting. Removed empty duties SHALL never create gaps or MV reminders. The second Reinigung assignments deleted by the one-time migration SHALL no longer count in personal statistics or receive ordinary helper reminders, including when their games are in the past; existing historical audits SHALL remain unchanged. Every remaining saved assignment SHALL count in personal statistics and receive ordinary helper reminders, including retained removed duties. Repeated-role recipients SHALL be resolved from their actual assignments without duplication or dropped occupants. An assignment seeking a replacement SHALL remain physically occupied by the responsible helper and SHALL remain included in personal statistics and reminders. Opening or closing a request, locking an assignment, or atomically transferring its responsibility SHALL NOT create a physical vacancy or change required-position counts; current age deficiencies SHALL still be evaluated independently.

#### Scenario: Removed youth duties do not create gaps

- **WHEN** a youth game's offered positions are occupied and age-eligible
- **THEN** Kasse and Wischer do not create a missing-duty entry or MV reminder

#### Scenario: Retained removed youth duty counted and notified

- **WHEN** a person holds a retained removed youth Wischer assignment
- **THEN** it remains in their duty statistics and ordinary helper reminders
- **AND** required progress is unaffected

#### Scenario: Replacement request does not create a gap

- **WHEN** an offered position is marked `Ersatz gesucht`
- **THEN** its current helper remains the occupant for staffing progress, personal duty statistics and ordinary reminders
- **AND** the request alone does not create a missing-duty entry or remove any independently reported age deficiency

#### Scenario: Atomic takeover preserves occupancy

- **WHEN** an eligible replacement takes over an occupied offered position
- **THEN** the physical occupancy and required-position totals remain unchanged
- **AND** personal duty statistics and subsequent reminders refer to the new responsible helper

#### Scenario: Fully staffed adult game after second-assignment deletion

- **WHEN** all seven offered adult positions are occupied and age-eligible after the former second Reinigung assignment was deleted by migration
- **THEN** required progress is 7/7 (100%) with no required vacancy or MV missing-task reminder
- **AND** the deleted assignment no longer counts in personal duty statistics or ordinary reminders
- **AND** its historical audit and migration-removal evidence remain readable

#### Scenario: Baseline-only adult staffing

- **WHEN** only the five baseline adult positions are occupied and age-eligible
- **THEN** progress is 5/7 and missing-task reporting identifies Kasse and one Wischer
