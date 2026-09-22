## ADDED Requirements

### Requirement: Day-block slots use compare-and-swap assignment

The system SHALL provide claim and release operations for one named slot of one
preparation or cleanup block. Each operation SHALL carry the occupant the caller expects,
and SHALL be refused without overwriting stored data when the current occupant differs.

#### Scenario: Concurrent block claim

- **WHEN** two callers claim the same empty block slot and the second request is processed
  after the first claim is stored
- **THEN** the second request is refused with a conflict and the current occupant
- **AND** the first claim remains stored

#### Scenario: Stale block release

- **WHEN** a caller releases a block slot whose occupant differs from the expected person
- **THEN** the release is refused with a conflict
- **AND** the current assignment remains stored

### Requirement: Block assignment authority has no team scope

An active member, including an MV, SHALL be allowed to claim an empty block slot only for
themselves and release only a block slot they hold. An admin SHALL be allowed to claim or
release any block slot for any active person. A block SHALL have no responsible team, and
MV status or team membership SHALL neither broaden nor restrict block assignment rights.
Non-admin changes SHALL be refused after the block's home-game date has passed, while an
admin SHALL remain able to correct past block assignments.

#### Scenario: Member claims own preparation slot

- **WHEN** an active member claims an empty preparation slot for themselves on a current
  or future game date
- **THEN** the assignment is recorded

#### Scenario: MV tries to assign another person

- **WHEN** an MV who is not an admin attempts to claim a block slot for another person
- **THEN** the request is refused regardless of either person's team memberships

#### Scenario: Admin staffs a cleanup block

- **WHEN** an admin claims an empty cleanup slot for an active person
- **THEN** the assignment is recorded without requiring a responsible team

#### Scenario: Member changes a past block

- **WHEN** a member or MV attempts to claim or release a block slot after its home-game
  date
- **THEN** the request is refused

### Requirement: One-task limits are scoped per assignment container

A person SHALL hold at most one task in a particular preparation block, cleanup block, or
game. Holding a task in one container SHALL not prevent that person from holding one task
in another container on the same date.

#### Scenario: Two tasks in one preparation block

- **WHEN** a person who already holds one preparation slot is claimed for a second slot
  in the same preparation block
- **THEN** the second claim is refused

#### Scenario: Preparation and game task on the same date

- **WHEN** a preparation assignee is claimed for a task in a game later that date
- **THEN** the game claim is permitted if its slot and other game-level rules allow it

#### Scenario: Preparation and cleanup tasks on the same date

- **WHEN** a preparation assignee is claimed for one cleanup slot on the same date
- **THEN** the cleanup claim is permitted

### Requirement: Unterstützung is an optional per-game task

The per-game role formerly named `Reinigung` SHALL be named `Unterstützung`. It SHALL
remain assignable, auditable, included in assigned-person statistics, and included in
helper notifications when occupied. An empty `Unterstützung` slot SHALL not make a game
incomplete, appear as an open duty, or cause an MV missing-task reminder.

#### Scenario: Empty Unterstützung slot

- **WHEN** every required task of an upcoming game is occupied but `Unterstützung` is
  empty
- **THEN** the game is treated as completely staffed
- **AND** statistics and MV reminders do not report the optional slot as missing

#### Scenario: Occupied Unterstützung slot

- **WHEN** a person is assigned to `Unterstützung`
- **THEN** the assignment appears on the schedule and in that person's statistics
- **AND** the person receives the applicable per-game reminders

#### Scenario: Existing Reinigung assignment migrated

- **WHEN** the schema migration encounters a current `Reinigung` assignment
- **THEN** it preserves the same game, slot, and person under `Unterstützung`
- **AND** historical audit snapshots retain their recorded role text

