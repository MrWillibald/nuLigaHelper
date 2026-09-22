## MODIFIED Requirements

### Requirement: Overview filters games by team and assigned helper

The overview SHALL offer filters for playing team, responsible team, and assigned person.
Selected team filters SHALL apply to games, and a date SHALL remain visible only when at
least one game on that date satisfies every selected team filter. An unset responsible
team SHALL not match a selected responsible team.

The person filter SHALL match the names of people assigned either to a game or to a
day-level preparation or cleanup block. A matching game assignment SHALL retain that
game. A matching block assignment SHALL retain the date and all games on that date that
satisfy the selected team filters. Filtering SHALL not search the roster or empty slots.
It SHALL retain season ordering and SHALL show date and month headings only when they
contain matching content.

#### Scenario: Combine filters

- **WHEN** a visitor selects a playing team and responsible team and enters an assigned
  person's name
- **THEN** the overview shows only dates satisfying the person filter and games satisfying
  both team filters
- **AND** the remaining games stay in chronological order within their date groups

#### Scenario: Responsible team is open

- **WHEN** a visitor selects a responsible team
- **THEN** games with no responsible team are excluded

#### Scenario: Duplicate names

- **WHEN** multiple people with the same displayed name hold assignments on different
  games or day blocks
- **THEN** a name search can match each assignment without treating the name as a unique
  person identity

#### Scenario: Person matches a day block

- **WHEN** the person filter matches a preparation or cleanup assignee
- **THEN** that date remains visible
- **AND** its games remain subject to any selected team filters

## ADDED Requirements

### Requirement: Day task blocks bookend each visible game date

The overview SHALL render the preparation block before the first displayed game of its
date and the cleanup block after the last displayed game of its date. Each block SHALL
show its calculated time, numbered display slots, and assigned helper names. The numbers
SHALL distinguish UI positions without creating separate task roles. It SHALL not show or ask
for a responsible team. A guest response SHALL expose no roster payload, person IDs,
contact data, or assignment controls for the blocks. Block cards SHALL use the same plain
surface and border treatment as game cards, without a phase-colored background fade or side
bar. A `#a0e656` `↗` icon SHALL identify preparation, and a `#ffb752` `↘` icon SHALL
identify cleanup.

#### Scenario: Unfiltered date presentation

- **WHEN** a visitor opens a date containing several games
- **THEN** the preparation block appears before all game cards
- **AND** the cleanup block appears after all game cards

#### Scenario: Filter leaves only a middle game

- **WHEN** filtering retains a date but hides its original first or last game
- **THEN** both blocks still bookend the displayed games
- **AND** their times remain derived from the actual first and last games of the full date

#### Scenario: Guest views block assignments

- **WHEN** an unauthenticated visitor views a game date
- **THEN** assigned block helper names are visible
- **AND** the response contains no block person IDs, roster payload, contact data, or
  assignment controls
