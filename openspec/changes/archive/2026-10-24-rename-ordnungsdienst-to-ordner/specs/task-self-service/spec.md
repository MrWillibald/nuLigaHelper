# Spec Delta

## MODIFIED Requirements

### Requirement: Kasse replaces the per-game Unterstützung task

The current per-game role named Unterstützung SHALL be named Kasse. It SHALL remain assignable, auditable, included in assigned-person statistics and included in helper notifications when occupied. It SHALL also carry the additional ordering/security duties of the rejected second Ordner while Ordner remains a separate required singleton. Kasse SHALL be offered and required only for adult games. Youth and unknown-category games SHALL not offer new Kasse claims. Retained Kasse assignments SHALL stay visible under existing rights, remain in personal statistics and helper reminders, and SHALL not affect required progress, open duties or MV follow-up. They SHALL offer existing authorized release-only retirement before the fixed cutoff, SHALL refuse ordinary release after cutoff and SHALL never offer a new claim, replacement request or takeover. Destructive account and game/date/block maintenance that would cancel a binding commitment SHALL be refused atomically. No optional game duties SHALL be offered.

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
- **THEN** its occupant is shown as an existing assignment subject to binding protection
- **AND** authorized release-only retirement is available before the fixed cutoff and ordinary release is refused afterward
- **AND** no new candidate, replacement claim or empty Kasse field is offered
- **AND** any explicitly authorized retirement appends an audit without changing historical snapshots
