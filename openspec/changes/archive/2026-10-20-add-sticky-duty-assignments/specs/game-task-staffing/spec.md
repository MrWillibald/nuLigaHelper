# Spec Delta

## MODIFIED Requirements

### Requirement: Games expose their offered task position set

Adult games SHALL offer one Zeitnehmer, one Sekretär, two Verkauf, exactly one independent Ordnungsdienst, one Kasse and two Reinigung positions. Youth games and unrecognized categories SHALL offer only Zeitnehmer, Sekretär, two Verkauf and exactly one Ordnungsdienst. Numbers SHALL identify positions of a semantic role. Kasse SHALL additionally perform the ordering/security duties planned for a second Ordner without creating another assignment or relaxing one task per person per game. Per-game Reinigung SHALL remain distinct from day-wide Aufräumen.

#### Scenario: Adult expanded game duties

- **WHEN** an adult game's assignments are opened
- **THEN** eight positions are visible or assignable according to existing rights
- **AND** Reinigung has two numbered positions and Ordnungsdienst has exactly one

#### Scenario: Youth offered duties

- **WHEN** a youth game including SPF is opened
- **THEN** only the baseline five positions are offered
- **AND** empty Kasse and Reinigung fields and optional markers are absent
- **AND** new Kasse and Reinigung claims are refused for every tier and CLI/system caller

#### Scenario: Independent repeated duty positions

- **WHEN** two different eligible people hold adult Reinigung positions 0 and 1
- **THEN** both occupants remain visible as independent assignments subject to binding responsibility, replacement-request and day-before lock rules
- **AND** each assignment identifies its original position

### Requirement: Offered positions are required by category

Every offered position SHALL be required. Adult M/F classes SHALL have eight required positions. Youth mA–mE/wA–wE and SPF classes SHALL have five. GE, missing, unsupported and conflicting labels SHALL retain the five-position baseline and expose unresolved classification. The shared explicit classifier SHALL be used with the existing age rules. New claims and takeovers SHALL revalidate offeredness against saved game state under write serialization and after retries. A replacement request SHALL NOT turn an occupied position into a free claim target or authorize a takeover of a position outside the offered set.

#### Scenario: Adult additional gaps

- **WHEN** an adult game's baseline five positions are occupied and its added duties are empty
- **THEN** three required positions remain missing: Kasse and both Reinigung places

#### Scenario: Youth complete baseline

- **WHEN** a youth game's baseline five positions are occupied
- **THEN** there are no required vacancies
- **AND** removed Kasse and Reinigung duties do not generate gaps

#### Scenario: Youth Ordnungsdienst is required

- **WHEN** a youth game's Ordnungsdienst is empty
- **THEN** that required position remains missing even if a retained removed duty is occupied

#### Scenario: Classification is unresolved

- **WHEN** the class is GE or cannot be recognized
- **THEN** no adult-only duties are offered or required
- **AND** the baseline five and unresolved-classification status are retained

#### Scenario: Saved category changed before claim

- **WHEN** a caller holds stale adult game state and the saved game has become youth
- **THEN** a new Kasse or Reinigung claim is refused using the saved category
- **AND** no assignment or claim audit is written

#### Scenario: Saved category changed before takeover

- **WHEN** an adult-only position has an open replacement request but the saved game has become youth before takeover
- **THEN** the takeover is refused because that position is no longer offered
- **AND** the original assignment remains visible with its history and without changing youth required progress

### Requirement: Removed duties retain their occupants with release-only maintenance

An existing assignment outside the game's offered set SHALL be preserved with its identity and history and shown under current privacy rules. It SHALL offer no new claim or replacement candidate. It SHALL continue to enforce one task per person per game, contribute to personal statistics and receive ordinary helper reminders without contributing to required staffing.

Existing authorized release-only maintenance SHALL remain available before the fixed day-before cutoff solely to retire an obsolete duty. After cutoff ordinary release SHALL be refused for every tier, including administrators. No claim, replacement request or takeover SHALL be offered for the removed role. A permitted pre-cutoff retirement SHALL append a readable audit, close any stale replacement state, remove only the retired field and leave required staffing unchanged. Historical snapshots SHALL remain unchanged. Destructive account and game/date/block maintenance that would cancel a binding commitment SHALL be refused atomically; it SHALL NOT authorize ordinary reassignment.

#### Scenario: Migrated youth Kasse retained

- **WHEN** migration renames a youth Unterstützung assignment to Kasse
- **THEN** it remains visible as an existing assignment with the same id, person, game and position
- **AND** no new youth Kasse claim or replacement candidate is offered
- **AND** authorized release-only retirement remains possible before cutoff and ordinary release is refused afterward

#### Scenario: Adult game becomes youth

- **WHEN** a staffed adult game changes to a youth category
- **THEN** Kasse and Reinigung occupants remain visible without new candidates
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

Progress, open-duty statistics, completeness and MV missing-task reminders SHALL use the same offered/required positions and retain independent age-deficiency reporting. Removed empty duties SHALL never create gaps or MV reminders. Every saved assignment SHALL count in personal statistics and receive ordinary helper reminders, including retained removed duties. Repeated-role recipients SHALL be resolved from their actual assignments without duplication or dropped occupants. An assignment seeking a replacement SHALL remain physically occupied by the responsible helper and SHALL remain included in personal statistics and reminders. Opening or closing a request, locking an assignment, or atomically transferring its responsibility SHALL NOT create a physical vacancy or change required-position counts; current age deficiencies SHALL still be evaluated independently.

#### Scenario: Removed youth duties do not create gaps

- **WHEN** a youth game's offered positions are occupied and age-eligible
- **THEN** Kasse and Reinigung do not create a missing-duty entry or MV reminder

#### Scenario: Retained removed youth duty counted and notified

- **WHEN** a person holds a retained removed youth Reinigung assignment
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
