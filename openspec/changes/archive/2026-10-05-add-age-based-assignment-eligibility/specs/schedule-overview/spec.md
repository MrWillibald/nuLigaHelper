# Spec Delta

## ADDED Requirements

### Requirement: Age compliance is distinct from physical staffing progress

Game cards SHALL distinguish physical required-slot occupancy from unresolved or invalid age eligibility. An age deficiency SHALL NOT invent an empty slot, conceal an assigned occupant, or alter the physical occupancy count. A game with an outstanding age deficiency SHALL not be presented as having complete valid staffing solely because its occupancy bar is full.

Current/future game cards SHALL provide a readable German eligibility indication identifying the affected duty or adult-coverage requirement without exposing a full birth date or exact personal age. Saved assignment changes and refreshed game data SHALL update the affected indication.

#### Scenario: Full game lacks an adult seller

- **WHEN** every required slot of a current/future game is occupied but Verkauf has no qualifying adult seller
- **THEN** the occupancy count remains accurate
- **AND** the card visibly reports outstanding adult coverage instead of complete valid staffing

#### Scenario: Timing eligibility is unresolved

- **WHEN** a current/future timing assignment cannot be verified because of missing birth-date or game information
- **THEN** the assigned helper remains visible
- **AND** the duty carries an unresolved-eligibility indication

#### Scenario: Stored data changes compliance

- **WHEN** a successful correction or assignment mutation changes a game's eligibility status
- **THEN** the card reflects saved occupancy and eligibility state
- **AND** a rejected mutation does not display an uncommitted improvement

#### Scenario: Guest sees eligibility status

- **WHEN** a guest views a deficient game's public card
- **THEN** the card can explain the outstanding duty requirement
- **AND** its response contains no full birth date, exact personal age, private roster payload, or assignment controls
