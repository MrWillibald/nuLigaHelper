## ADDED Requirements

### Requirement: On-demand candidates preserve assignment selection behavior

Loading candidates after card expansion SHALL preserve the applicable candidate ordering, complete team labels, playing-team and outside-team hints, and the assignment scope of each slot. A person already assigned to another task in the same game or day-task block SHALL be excluded; the current occupant SHALL remain visible in their own slot even if they are no longer an active candidate. Successful claims and releases SHALL keep sibling controls consistent without loading unrelated cards. Assignment writes SHALL retain their per-slot compare-and-swap checks.

#### Scenario: Expand a staffed game

- **WHEN** a viewer opens a game with one or more occupied slots
- **THEN** each occupant remains selected in their own slot
- **AND** no occupant is offered in another slot of that game
- **AND** available candidates retain the existing category order and hints

#### Scenario: Expand a staffed day-task block

- **WHEN** a viewer opens a preparation or cleanup block with an occupied slot
- **THEN** that occupant remains selected in their own slot and is absent from its other slots
- **AND** a task in another game or block does not exclude that person

#### Scenario: Change an assignment on an open card

- **WHEN** a claim or release succeeds
- **THEN** the other controls in that same card update their candidate availability
- **AND** the stored occupant remains authoritative if a later write detects a stale expectation
