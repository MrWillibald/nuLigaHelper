# Schedule Overview Specification

## Purpose

Defines how visitors find relevant home games and distinguish completed dates from upcoming dates on the public game overview.

## Requirements

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

### Requirement: Filters are usable by all schedule viewers

The overview SHALL allow guests and signed-in viewers to apply and clear filters. The person filter SHALL be a case-insensitive name search over assigned people visible on game cards. The guest response SHALL contain no roster payload, person IDs, or contact data. When no games match, the overview SHALL show a clear German empty result message and a way to clear the filters; when the season has no games at all, it SHALL retain a distinct no-games message.

#### Scenario: Guest filters by assigned name

- **WHEN** a guest searches for an assigned person's name
- **THEN** matching game cards are shown without a roster, person IDs, or contact data in the response

#### Scenario: No matching games

- **WHEN** filters exclude every game in the season
- **THEN** the overview explains that no games match and offers a clear-filter action

#### Scenario: Clear filters

- **WHEN** a visitor clears the active filters
- **THEN** all games for the current season appear in their normal upcoming and past sections

### Requirement: Past games remain available in a muted expandable section

Games whose date is before the application's effective current day SHALL appear in a separate past-games section that is collapsed by default and can be expanded and collapsed by the visitor. Games on the current day SHALL remain in the upcoming section. Past games SHALL have a visually muted but readable treatment consistent with the rest of the overview. Filtering SHALL apply to both sections, and the past section's label SHALL reflect the number of matching past game days.

#### Scenario: Open and close past games

- **WHEN** a visitor opens the overview with past games
- **THEN** past game cards are initially hidden behind a labeled expand control
- **AND** activating the control reveals them and activating it again hides them

#### Scenario: Effective date boundary

- **WHEN** a game's date is the effective current day
- **THEN** that game is shown with upcoming games

#### Scenario: Filter past games

- **WHEN** a visitor applies a filter that matches only past games
- **THEN** the collapsed past section indicates the number of matching past game days
- **AND** expanding it reveals those matching games without unrelated dates or month headings
