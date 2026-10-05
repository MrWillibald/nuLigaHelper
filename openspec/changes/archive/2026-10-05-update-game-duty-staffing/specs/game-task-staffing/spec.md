# Spec Delta

## Purpose

Defines category-dependent game duty positions and preserves removed-duty assignments without offering new claims.

## ADDED Requirements

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
- **THEN** both occupants remain visible and independently releasable
- **AND** each assignment identifies its original position

### Requirement: Offered positions are required by category

Every offered position SHALL be required. Adult M/F classes SHALL have eight required positions. Youth mA–mE/wA–wE and SPF classes SHALL have five. GE, missing, unsupported and conflicting labels SHALL retain the five-position baseline and expose unresolved classification. The shared explicit classifier SHALL be used with the existing age rules. New claims SHALL revalidate offeredness against saved game state under write serialization and after retries.

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

### Requirement: Removed duties retain their occupants with release-only maintenance

An existing assignment outside the game's offered set SHALL be preserved with its identity and history. It SHALL be shown as an existing assignment under the current privacy rules and SHALL remain releasable through existing authority and CAS semantics. It SHALL offer no new claim or replacement candidate. A successful release SHALL append an audit and remove the retained field. It SHALL continue to enforce one task per person per game, contribute to personal statistics and receive ordinary helper reminders without contributing to required staffing.

#### Scenario: Migrated youth Kasse retained

- **WHEN** migration renames a youth Unterstützung assignment to Kasse
- **THEN** it remains visible as an existing assignment with the same id, person, game and position
- **AND** authorized users may release it without offering a new youth Kasse claim

#### Scenario: Adult game becomes youth

- **WHEN** a staffed adult game changes to a youth category
- **THEN** Kasse and Reinigung occupants remain visible and releasable without new candidates
- **AND** required progress and reporting use the youth baseline

### Requirement: Staffing reports share required positions

Progress, open-duty statistics, completeness and MV missing-task reminders SHALL use the same offered/required positions and retain independent age-deficiency reporting. Removed empty duties SHALL never create gaps or MV reminders. Every saved assignment SHALL count in personal statistics and receive ordinary helper reminders, including retained removed duties. Repeated-role recipients SHALL be resolved from their actual assignments without duplication or dropped occupants.

#### Scenario: Removed youth duties do not create gaps

- **WHEN** a youth game's offered positions are occupied and age-eligible
- **THEN** Kasse and Reinigung do not create a missing-duty entry or MV reminder

#### Scenario: Retained removed youth duty counted and notified

- **WHEN** a person holds a retained removed youth Reinigung assignment
- **THEN** it remains in their duty statistics and ordinary helper reminders
- **AND** required progress is unaffected
