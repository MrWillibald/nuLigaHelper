# Spec Delta

## MODIFIED Requirements

### Requirement: Games expose their offered task position set

Adult games SHALL offer one Zeitnehmer, one Sekretär, two Verkauf, exactly one independent Ordner, one Kasse and exactly one Wischer position (position 0). Youth games and unrecognized categories SHALL offer only Zeitnehmer, Sekretär, two Verkauf and exactly one Ordner. Numbers SHALL identify positions of repeated semantic roles and the original position of a retained removed duty. Kasse SHALL additionally perform the ordering/security duties planned for a second Ordner without creating another assignment or relaxing one task per person per game. Per-game Wischer SHALL remain distinct from day-wide Aufräumen.

#### Scenario: Adult expanded game duties

- **WHEN** an adult game's assignments are opened
- **THEN** seven offered positions are visible or assignable according to existing rights
- **AND** Wischer is shown as one unnumbered offered position and Ordner has exactly one

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

- **WHEN** a youth game's Ordner is empty
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

## ADDED Requirements

### Requirement: Ordner is the consistent current security-duty label

The independent security duty formerly displayed as Ordnungsdienst SHALL be presented as the singular, unnumbered Ordner in current schedule fields, accessible task names, candidate guidance, JavaScript and non-JavaScript confirmation flows, replacement controls, eligibility messages, statistics, CLI output and newly generated helper, rescheduling, MV and team-vacancy notifications. The label SHALL refer to the same single position; it SHALL NOT introduce an additional role or change staffing counts, eligibility, authorization, responsibility, replacement generations or cutoff rules.

Stored assignment identifiers and existing JSON mutation and candidate-key contracts SHALL remain compatible with the role identifier Ordnungsdienst. The CLI SHALL accept Ordner as the current human-facing role argument and retain Ordnungsdienst as a compatibility alias for the same duty. Raw role identifiers in transport fields, form values and data attributes SHALL NOT be used as visible labels. New audit transitions SHALL capture the label Ordner; existing audit entries SHALL retain their recorded role names and snapshots.

#### Scenario: Existing assignment receives the current label

- **WHEN** an existing Ordnungsdienst assignment is displayed after the rename, including one on a past game, a locked one or one seeking replacement
- **THEN** its current assignment label is Ordner with the same assignment identity, occupant, slot, lock and replacement state
- **AND** its statistics and required-position contribution remain unchanged
- **AND** the rename alone writes no assignment mutation, audit transition or notification

#### Scenario: Compatible role identifiers keep one duty

- **WHEN** an authorized client uses the existing Ordnungsdienst role identifier in a claim, replacement, closure or takeover request
- **THEN** the request targets the same single duty under all existing checks
- **AND** returned human-readable task labels use Ordner while candidate lookup and mutation identifiers retain their existing contract
- **AND** an Ordner display label does not create a second assignable API role

#### Scenario: CLI supports the current name and the compatibility alias

- **WHEN** an operator invokes a role-taking CLI command using Ordner or the retained Ordnungsdienst alias
- **THEN** both resolve to the same saved role and apply the same existing authorization, eligibility, occupancy and cutoff rules
- **AND** user-facing task output and supported-role guidance show Ordner as the current name

#### Scenario: Newly generated notifications use Ordner

- **WHEN** an e-mail or SMS mentions the security duty as an assigned helper task, missing task, eligibility deficiency or advertised team vacancy
- **THEN** the rendered message uses Ordner
- **AND** message contracts, recipient selection, timing, routing and dispatch counts retain their existing behavior

#### Scenario: Audit history preserves the name in force

- **WHEN** an administrator reviews a historical Ordnungsdienst entry and a newly recorded Ordner transition
- **THEN** the old entry still displays its recorded Ordnungsdienst label and the new entry displays Ordner
- **AND** both remain readable for the same game and numbered position without rewriting earlier evidence

#### Scenario: Confirmation cancellation remains free of writes

- **WHEN** a helper opens a claim or takeover confirmation labeled Ordner and then cancels or dismisses it, with or without JavaScript
- **THEN** no assignment, replacement request, successful mutation audit or notification is written
- **AND** the saved occupant and responsibility remain authoritative
