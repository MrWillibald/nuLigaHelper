## ADDED Requirements

### Requirement: Overview filters by a selected game date

The overview SHALL offer a single-date dropdown containing the dates of all home games
in the current season in chronological order and an all-dates option selected by default.
The date filter SHALL combine with the playing-team, responsible-team, and assigned-person
filters. When a date is selected, only games and preparation or cleanup blocks on that
date that satisfy the other active filters SHALL be visible. A nonempty date value that
does not identify a current-season game date SHALL match no cards rather than show all dates.
The date filter SHALL work for guests and signed-in viewers without JavaScript. Clearing
filters SHALL restore the all-dates option.

#### Scenario: Select a date

- **WHEN** a visitor selects a date with multiple home games
- **THEN** only that date's games and its preparation and cleanup blocks are shown
- **AND** the selected date remains selected after the page loads

#### Scenario: Combine date and existing filters

- **WHEN** a visitor selects a date and a playing team
- **THEN** only games on that date matching the playing team are shown
- **AND** the date's blocks bookend those matching games

#### Scenario: Date with no remaining match

- **WHEN** the selected date has no games or blocks matching the other active filters, or the date
  value is not a current-season game date
- **THEN** the ordinary no-results message and clear-filter action are shown
- **AND** no other date or block is shown

### Requirement: Every game and day-task card starts compact

Every game, preparation block, and cleanup block SHALL start collapsed on page load,
regardless of whether a game has a responsible team or occupied slots. A collapsed game
card SHALL show its time, matchup or Spielfest identity, existing game metadata, and a
staffing progress bar. A collapsed block card SHALL show its calculated time, block
label, and staffing progress bar. Each card SHALL have a labeled control that expands
and collapses that card independently. The expanded game SHALL reveal the responsible
team and all existing task fields, including optional `Unterstützung`; the expanded
block SHALL reveal all three numbered assignment fields. All viewers SHALL be able to
expand cards, while assignment controls and private data SHALL continue to depend on
their existing access rights. Expansion SHALL be usable with a keyboard.

#### Scenario: Open and close a game

- **WHEN** a visitor opens the overview containing a game without a responsible team
- **THEN** that game starts collapsed like every other game
- **AND** expanding it reveals the responsible-team field and every task field
- **AND** collapsing it hides those fields again

#### Scenario: Open a day-task block

- **WHEN** a visitor expands a preparation or cleanup card
- **THEN** all three numbered assignment fields and their visible occupants are shown
- **AND** other cards retain their current expansion states

### Requirement: Compact cards show accurate staffing progress

Each game progress bar SHALL count occupied slots among the five required game task
slots and SHALL exclude optional `Unterstützung` and responsible-team selection from
both numerator and denominator. Each preparation and cleanup progress bar SHALL count
occupied slots among that block's three slots. Every bar SHALL have a readable count
and whole-number percentage, including at zero and full coverage, so meaning does not
depend on color alone. A successful assignment or release SHALL update the affected
card's progress without requiring a page reload; an unsuccessful or stale operation
SHALL not show uncommitted progress.

#### Scenario: Optional support does not complete a game

- **WHEN** a game has three occupied required slots and an occupied `Unterstützung` slot
- **THEN** its compact card reports 3 of 5 required slots filled and 60 percent

#### Scenario: Block coverage

- **WHEN** one of a preparation block's three slots is occupied
- **THEN** its compact card reports 1 of 3 slots filled and 33 percent

#### Scenario: Assignment changes progress

- **WHEN** a viewer successfully claims or releases an assignment on an expanded card
- **THEN** that card's count, percentage, and bar length reflect the saved assignment
- **AND** a rejected assignment leaves the displayed progress at its saved value

## MODIFIED Requirements

### Requirement: Overview filters games by team and assigned helper

The overview SHALL offer filters for playing team, responsible team, and assigned person.
Selected team filters SHALL apply to games, and a date SHALL remain visible only when at
least one game on that date satisfies every selected team filter. An unset responsible
team SHALL not match a selected responsible team. A selected responsible team SHALL
hide all preparation and cleanup blocks because blocks have no responsible team.

The person filter SHALL match the names of people assigned to each game or day-level
preparation or cleanup block independently. It SHALL retain a game only when that game
has a matching assignee, and retain a block only when that block has a matching assignee.
A matching block SHALL be allowed to retain its date without any displayed game, provided
the date satisfies the selected team filters. Filtering SHALL not search the roster or
empty slots. It SHALL retain season ordering and SHALL show date and month headings only
when they contain matching content.

#### Scenario: Combine filters

- **WHEN** a visitor selects a playing team and responsible team and enters an assigned
  person's name
- **THEN** the overview shows only games satisfying both team filters and the person filter
- **AND** no day-task blocks are shown
- **AND** the remaining games stay in chronological order within their date groups

#### Scenario: Responsible team is open

- **WHEN** a visitor selects a responsible team
- **THEN** games with no responsible team are excluded
- **AND** preparation and cleanup blocks are excluded

#### Scenario: Duplicate names

- **WHEN** multiple people with the same displayed name hold assignments on different
  games or day blocks
- **THEN** a name search can match each assignment without treating the name as a unique
  person identity

#### Scenario: Person matches a day block

- **WHEN** the person filter matches a preparation or cleanup assignee but no game
  assignee on a qualifying date, and no responsible-team filter is selected
- **THEN** the date and only the matching block or blocks remain visible
- **AND** no game is shown merely because its date contains a matching block

#### Scenario: Person matches a game only

- **WHEN** the person filter matches a game's assignee but no day-block assignee
- **THEN** the matching game remains visible without either day-task block

### Requirement: Day task blocks bookend each visible game date

When visible, the overview SHALL render the preparation block before displayed games of
its date and the cleanup block after displayed games of its date. Filtering SHALL omit a
block that does not match and SHALL allow a matching block to appear without a displayed game.
Each block SHALL show its calculated time and staffing progress while collapsed, and SHALL show its
numbered display slots and assigned helper names when expanded. The numbers SHALL
distinguish UI positions without creating separate task roles. It SHALL not show or ask
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

- **WHEN** filtering retains only a middle game without a person or responsible-team filter
- **THEN** both blocks still bookend the displayed games
- **AND** their times remain derived from the actual first and last games of the full date

#### Scenario: Filter leaves only a day-task block

- **WHEN** a person filter matches only one block on a date with no matching game assignments
- **THEN** that block appears under the date heading without unrelated games or the other block
- **AND** its time remains derived from the full date's games

#### Scenario: Guest views block assignments

- **WHEN** an unauthenticated visitor expands a day-task block
- **THEN** assigned block helper names are visible
- **AND** the response contains no block person IDs, roster payload, contact data, or
  assignment controls

### Requirement: Filters are usable by all schedule viewers

The overview SHALL allow guests and signed-in viewers to apply and clear filters. The
person filter SHALL be a case-insensitive name search over assigned people visible when
game or day-task block cards are expanded. The guest response SHALL contain no roster
payload, person IDs, or contact data. When no cards match, the overview SHALL show a
clear German empty result message and a way to clear the filters; when the season has
no games at all, it SHALL retain a distinct no-games message.

#### Scenario: Guest filters by assigned name

- **WHEN** a guest searches for an assigned person's name
- **THEN** matching game or day-task block cards are shown without a roster, person IDs,
  or contact data in the response

#### Scenario: No matching games

- **WHEN** filters exclude every game and day-task block in the season
- **THEN** the overview explains that no results match and offers a clear-filter action

#### Scenario: Clear filters

- **WHEN** a visitor clears the active filters
- **THEN** all games and day-task blocks for the current season appear in their normal
  upcoming and past sections

### Requirement: Past games remain available in a muted expandable section

Games whose date is before the application's effective current day SHALL appear in a
separate past-games section that is collapsed by default and can be expanded and collapsed
by the visitor. When a specific past date is selected, that section SHALL start expanded
so the selected date is immediately visible. Individual game and block cards SHALL still
start collapsed. Games on the current day SHALL remain in the upcoming section. Past games
SHALL have a visually muted but readable treatment consistent with the rest of the
overview. Filtering SHALL apply to both sections, and the past section's label SHALL
reflect the number of matching past game days.

#### Scenario: Open and close past games

- **WHEN** a visitor opens the overview with past games and no date selected
- **THEN** past game cards are initially hidden behind a labeled expand control
- **AND** activating the control reveals them and activating it again hides them

#### Scenario: Effective date boundary

- **WHEN** a game's date is the effective current day
- **THEN** that game is shown with upcoming games

#### Scenario: Filter past games

- **WHEN** a visitor applies a filter other than a date selection that matches only past games
- **THEN** the collapsed past section indicates the number of matching past game days
- **AND** expanding it reveals those matching games without unrelated dates or month headings

#### Scenario: Select a past date

- **WHEN** a visitor selects a past date from the date dropdown
- **THEN** the past-games section starts expanded with only matching content from that date
- **AND** its game and block cards start collapsed
