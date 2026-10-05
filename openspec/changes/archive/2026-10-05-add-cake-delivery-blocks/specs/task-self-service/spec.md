# Spec Delta

## MODIFIED Requirements

### Requirement: Day-block slots use compare-and-swap assignment

The system SHALL provide claim and release operations for one named slot of one
preparation, cake-delivery or cleanup block. Each operation SHALL carry the occupant the caller expects,
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

### Requirement: One-task limits are scoped per assignment container

A person SHALL hold at most one task in a particular preparation block, cake-delivery block, cleanup block, or
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

#### Scenario: Two cakes in one block

- **WHEN** a person who already holds one cake position is claimed for another position in that same cake block
- **THEN** the second claim is refused

#### Scenario: Cake and game duties on the same date

- **WHEN** a cake volunteer claims a position in another block or game on that date
- **THEN** the claim is permitted if the other container and its rules allow it
