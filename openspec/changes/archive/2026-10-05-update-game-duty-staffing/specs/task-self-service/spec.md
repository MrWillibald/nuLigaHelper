# Spec Delta

## RENAMED Requirements

- FROM: `### Requirement: Unterstützung is an optional per-game task`
- TO: `### Requirement: Kasse replaces the per-game Unterstützung task`

## MODIFIED Requirements

### Requirement: No release cutoff

The system SHALL allow a release at any time before the game, including after
notifications for that game have been sent. A freed required position SHALL be reported
as missing again by the statistics and by the notification that chases open slots.
Releasing a retained removed-duty assignment SHALL not create a required vacancy or
change required-duty progress or MV follow-up.

#### Scenario: Release after the reminder went out

- **WHEN** a member releases a required position after the notification for that game has been sent
- **THEN** the release succeeds
- **AND** the game appears again among the games with missing assignments

#### Scenario: Retained removed duty released after the reminder

- **WHEN** a member releases their retained removed-duty assignment before the game after its reminder has been sent
- **THEN** the release succeeds and appends an audit
- **AND** required-duty progress, missing-duty statistics and MV follow-up are unchanged

### Requirement: Kasse replaces the per-game Unterstützung task

The current per-game role named Unterstützung SHALL be named Kasse. It SHALL remain assignable, auditable, included in assigned-person statistics and included in helper notifications when occupied. It SHALL also carry the additional ordering/security duties of the rejected second Ordner while Ordnungsdienst remains a separate required singleton. Kasse SHALL be offered and required only for adult games. Youth and unknown-category games SHALL not offer new Kasse claims. Retained Kasse assignments SHALL stay visible and releasable under existing rights, remain in personal statistics and helper reminders, and SHALL not affect required progress, open duties or MV follow-up. No optional game duties SHALL be offered.

The migration SHALL preserve every current Unterstützung assignment's game, person and position under Kasse. It SHALL retain historical audit snapshots with the role text recorded at the time. The newly introduced Reinigung positions SHALL not receive those migrated assignments.

#### Scenario: Empty Unterstützung slot

- **WHEN** every required position of an upcoming youth game is occupied and there is no retained Kasse assignment
- **THEN** the game is treated as completely staffed
- **AND** statistics and MV reminders do not report that removed position as missing

#### Scenario: Empty required Kasse position

- **WHEN** an adult game's Kasse position is empty
- **THEN** that position is reported as a missing required duty

#### Scenario: Occupied Unterstützung slot

- **WHEN** a person is assigned to Kasse
- **THEN** the assignment appears on the schedule and in that person's statistics
- **AND** the person receives the applicable per-game reminders

#### Scenario: Existing Unterstützung assignment migrated

- **WHEN** migration encounters a current Unterstützung assignment
- **THEN** it preserves the same game, position and person under Kasse
- **AND** historical audit snapshots retain their recorded role text
- **AND** the new Reinigung positions begin empty

#### Scenario: Role collision is refused

- **WHEN** migration would overwrite a pre-existing Kasse assignment
- **THEN** migration refuses the collision without silently discarding either assignment

#### Scenario: Existing Reinigung assignment migrated

- **WHEN** a recognized older database still contains original Reinigung assignments and is upgraded through the reviewed migration chain
- **THEN** the existing Reinigung-to-Unterstützung revision and the new Unterstützung-to-Kasse revision preserve the same game, position and person under Kasse
- **AND** historical audit snapshots retain their recorded role text
- **AND** the newly introduced Reinigung positions do not inherit those legacy assignments

#### Scenario: Migrated youth Kasse is release-only

- **WHEN** a youth game has a migrated Kasse assignment
- **THEN** its occupant is shown as an existing assignment and may be released according to existing rights
- **AND** no new candidate, replacement claim or empty Kasse field is offered
- **AND** release appends an audit without changing historical snapshots
