# Game Task Staffing Specification

## Purpose

Defines category-dependent game duty positions and preserves removed-duty assignments without offering new claims.

## Requirements

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

### Requirement: Responsible-team members receive a seven-day vacancy invitation

When the daily job processes the date exactly seven calendar days after its effective today, the system SHALL send a personal invitation for each game on that date with at least one physically vacant required position and a saved responsible team. The invitation SHALL use that team's current active memberships and SHALL NOT depend on an appointed MV. The current MV SHALL receive the invitation only when they meet the same recipient rules as every other member. Pending, rejected and inactive people and people outside the responsible team SHALL NOT receive it. Each qualifying person SHALL receive at most one invitation per game per invocation, identified by internal person identity rather than name or membership count. Invitations SHALL prefer e-mail, fall back to SMS when no usable e-mail is present, and skip contactless recipients with the existing warning and zero sent-message count.

The trigger and advertised vacancies SHALL use only the game's category-dependent required positions. A position seeking replacement SHALL remain occupied. Age or classification deficiencies without a physical vacancy SHALL NOT trigger this invitation. Team-independent day-block vacancies SHALL NOT trigger a responsible-team invitation. Existing assigned-helper reminders and day-before MV follow-up SHALL retain their existing conditions. Sending an invitation SHALL NOT assign a duty, change responsibility, alter replacement state or lock an assignment.

#### Scenario: Active responsible-team members are invited a week ahead

- **WHEN** the daily job runs seven days before a game with an assigned responsible team and a vacant required position
- **THEN** each active, otherwise qualifying member of that team receives one personal invitation for that game
- **AND** pending, rejected and inactive people and non-members receive none

#### Scenario: A team needs no appointed MV for recruitment

- **WHEN** a game seven days away has required vacancies and a responsible team without an MV
- **THEN** its qualifying active team members still receive invitations
- **AND** no fallback team or administrator recipient is invented

#### Scenario: A qualifying MV receives one member invitation

- **WHEN** the responsible team's active MV holds no assignment in the game
- **THEN** the MV receives one invitation under the ordinary member rules
- **AND** the MV appointment does not create a second invitation

#### Scenario: Games outside the seven-day date are not recruited for

- **WHEN** the daily job runs and games have required vacancies six or eight days later
- **THEN** those games receive no invitation from this seven-day flow

#### Scenario: Missing responsible team gives no recruitment recipients

- **WHEN** a game seven days away has vacancies but no responsible team
- **THEN** no team vacancy invitation is sent for that game

#### Scenario: Required vacancies follow the game category

- **WHEN** an adult game has its five baseline positions filled but Kasse and Wischer are empty
- **THEN** its invitation lists one Kasse and one Wischer vacancy
- **AND** a youth or Spielfest game with all five required positions filled sends no invitation for empty Kasse or Wischer duties

#### Scenario: A replacement search or age deficiency is not a vacancy

- **WHEN** every required position is occupied but a replacement search or an age or classification deficiency exists
- **THEN** no team vacancy invitation is sent
- **AND** current occupants and existing follow-up behavior remain intact

#### Scenario: Team-independent block gaps do not recruit a team

- **WHEN** a date has an empty preparation, cleanup or cake position and its games have no required vacancies
- **THEN** no responsible-team vacancy invitation is generated for those blocks

#### Scenario: Contact routing and skipped counts stay consistent

- **WHEN** qualifying recipients include one person with both contacts, one with phone only and one with neither
- **THEN** the first receives only e-mail and the second receives SMS
- **AND** the contactless person is skipped with a warning and contributes zero to the sent count

#### Scenario: Multiple memberships and identical display names preserve identity

- **WHEN** a qualifying member belongs to several teams and another qualifying person has the same display name
- **THEN** each internal person identity receives one invitation for the game
- **AND** neither extra memberships nor a shared name duplicates or suppresses a recipient

### Requirement: Same-game assignments exclude vacancy-invitation recipients

The system SHALL exclude a responsible-team member from a game's vacancy invitation if that person holds any saved assignment in that game, including retained removed-duty assignments and assignments seeking a replacement. Exclusion SHALL use the internal game and person identities and SHALL apply regardless of task, offeredness or MV/admin status. An assignment in another game or in a day block SHALL NOT by itself exclude the member. No additional individual age filter SHALL limit this team invitation; actual claims SHALL continue to enforce the existing eligibility and one-duty rules.

#### Scenario: An already assigned member keeps only their helper reminder

- **WHEN** an active member of the responsible team already holds any duty in the game with vacancies
- **THEN** that person receives no vacancy invitation for that game
- **AND** their ordinary weekly assigned-helper reminder remains available

#### Scenario: Retained or offered replacement duties also exclude the helper

- **WHEN** a member holds a retained removed-duty assignment or a duty marked Ersatz gesucht in the game
- **THEN** that member receives no invitation to take another duty in that game
- **AND** the existing assignment and replacement state remain unchanged

#### Scenario: Duties in another container do not exclude membership

- **WHEN** an active responsible-team member holds a duty in another game or a day block but no duty in the game being recruited for
- **THEN** the member remains a recipient for that game's invitation

#### Scenario: The invitation is not an eligibility authorization

- **WHEN** an active, unassigned member with an unknown birth date belongs to the responsible team
- **THEN** the member remains a recipient
- **AND** any subsequent claim is checked under the existing task-specific eligibility rules

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
