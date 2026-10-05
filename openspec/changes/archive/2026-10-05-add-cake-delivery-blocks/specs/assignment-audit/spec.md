# Spec Delta

## MODIFIED Requirements

### Requirement: An entry describes the change completely

Each entry SHALL record the time of the change, the acting person, the tier the actor
acted as, the kind of change, the affected person, and the target task slot. A game entry
SHALL identify the game, role, and slot. A day-block entry SHALL identify the season,
home-game date, preparation, cake-delivery or cleanup phase, displayed task name, and slot.

#### Scenario: Entry content

- **WHEN** a game assignment change is recorded
- **THEN** the entry states when it happened, who made it, in which tier, what kind of
  change it was, whom it affected, and which task slot of which game was involved

#### Scenario: Day-block entry content

- **WHEN** a preparation, cake-delivery or cleanup assignment change is recorded
- **THEN** the entry states when it happened, who made it, in which tier, what kind of
  change it was, whom it affected, and which named slot of which dated block was involved

#### Scenario: Timestamps are precise

- **WHEN** entries are recorded
- **THEN** each carries a full date and time
- **AND** entries can be ordered by that time without relying on the German date format
  used for game and block dates

#### Scenario: Cake assignment history remains readable

- **WHEN** a cake position is assigned or explicitly released before a quantity reduction
- **THEN** the audit identifies the dated cake-delivery block and original position
- **AND** that description remains readable after the quantity changes or the block is deleted
