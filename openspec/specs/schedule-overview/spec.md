# Schedule Overview Specification

## Purpose

Defines how visitors find relevant home games and distinguish completed dates from upcoming dates on the public game overview.

## Requirements

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

### Requirement: Every game and day-task card starts compact

Every game, preparation block, and cleanup block SHALL start collapsed on page load,
regardless of whether a game has a responsible team or occupied slots. A collapsed game
card SHALL show its time, matchup or Spielfest identity, existing game metadata, its
responsible-team name or an open placeholder when unset, and a staffing progress bar.
A collapsed block card SHALL show its calculated time, block
label, and staffing progress bar. Each card SHALL have a labeled control that expands
and collapses that card independently. The expanded game SHALL reveal the responsible
team and all existing task fields, according to the game's offered set, plus occupied removed duties marked as existing assignments with release-only maintenance; the expanded
preparation or cleanup block SHALL reveal all three numbered assignment fields. All viewers SHALL be able to
expand cards, while assignment controls and private data SHALL continue to depend on
their existing access rights. Expansion SHALL be usable with a keyboard.

#### Scenario: Open and close a game

- **WHEN** a visitor opens the overview containing a game without a responsible team
- **THEN** that game starts collapsed like every other game and shows the open responsible-team placeholder
- **AND** expanding it reveals the responsible-team field and every task field
- **AND** collapsing it hides those fields again

#### Scenario: Responsible team remains visible in a collapsed game

- **WHEN** a visitor opens the overview containing a game with a responsible team
- **THEN** its collapsed summary shows that team's name without opening the card
- **AND** expanding it reveals the responsible-team field and every task field

#### Scenario: Open a day-task block

- **WHEN** a visitor expands a preparation or cleanup card
- **THEN** all three numbered assignment fields and their visible occupants are shown
- **AND** other cards retain their current expansion states

#### Scenario: Youth duties and retained assignments

- **WHEN** a viewer expands a youth game
- **THEN** only the five offered baseline positions appear as game duty fields
- **AND** empty Kasse and Reinigung fields and optional markers are absent
- **AND** any retained removed-duty occupant appears separately as an existing assignment with release-only controls according to existing rights

### Requirement: Compact cards show accurate staffing progress

Each game progress bar SHALL count occupied positions in that game's offered/required-position set and SHALL exclude retained removed-duty positions and responsible-team selection from both numerator and denominator. Adult games SHALL have eight required positions; youth games SHALL have five. Each preparation and cleanup progress bar SHALL count
occupied slots among that block's three slots. Every bar SHALL have a readable count
and whole-number percentage, including at zero and full coverage, so meaning does not
depend on color alone. A successful assignment or release SHALL update the affected
card's progress without requiring a page reload; an unsuccessful or stale operation
SHALL not show uncommitted progress.

#### Scenario: Optional support does not complete a game

- **WHEN** a youth game has three occupied required slots and an occupied retained removed Kasse slot
- **THEN** its compact card reports 3 of 5 required slots filled and 60 percent

#### Scenario: Block coverage

- **WHEN** one of a preparation block's three slots is occupied
- **THEN** its compact card reports 1 of 3 slots filled and 33 percent

#### Scenario: Assignment changes progress

- **WHEN** a viewer successfully claims or releases an assignment on an expanded card
- **THEN** that card's count, percentage, and bar length reflect the saved assignment
- **AND** a rejected assignment leaves the displayed progress at its saved value

#### Scenario: Adult additional duties count

- **WHEN** an adult game has five of its eight required positions occupied
- **THEN** its compact card reports 5 of 8 required positions filled

#### Scenario: Retained removed duty released

- **WHEN** a retained Kasse or Reinigung assignment of a youth game is released
- **THEN** its required-duty progress is unchanged and the removed duty field disappears
- **AND** changing its sole Ordnungsdienst position updates progress

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

### Requirement: Editable cards load candidates when opened

The initial schedule response SHALL show the current assignment state without including lists of unassigned candidates in collapsed or expanded card markup. When a signed-in viewer opens an editable game or day-task card, its candidate controls SHALL become usable after the permitted candidates have loaded. Opening one card SHALL not load candidates for other cards. The summary and assigned helper names SHALL remain available before candidate loading completes.

#### Scenario: Initial overview with a growing roster

- **WHEN** active, unassigned people are added to the roster and the same schedule is requested
- **THEN** the initial response does not repeat those people across task controls
- **AND** its candidate payload remains absent regardless of roster size

#### Scenario: Open one editable game

- **WHEN** a signed-in viewer opens an editable game card
- **THEN** that card loads its permitted candidates and makes its task controls usable
- **AND** unopened cards do not load their candidates

#### Scenario: Open an editable day-task block

- **WHEN** a signed-in viewer opens a preparation or cleanup card with editable slots
- **THEN** that block loads its permitted candidates independently of games and other blocks

#### Scenario: Candidate loading fails

- **WHEN** a candidate request fails or the session has expired
- **THEN** the affected controls remain unable to submit an assignment based on an incomplete list
- **AND** the viewer sees a German error or sign-in message and can retry after recovery
- **AND** current assignments remain visible and unchanged

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

### Requirement: Cake delivery appears as a compact dated task card

The overview SHALL show the date's cake-delivery block alongside its existing preparation, game and cleanup cards. The cake card SHALL start collapsed and SHALL show its label, configured delivery time or setup-needed state, requested cake quantity and staffing progress. Expanding it SHALL reveal one numbered volunteer position per requested cake and each position's assigned helper name. Each position SHALL represent one cake contribution. For a configured quantity, progress SHALL count occupied positions out of that quantity and SHALL update after successful assignments or configuration edits. A refused or stale operation SHALL not show uncommitted progress. Configured zero SHALL show that no cakes are requested without division-by-zero or invented staffing percentages.

The cake card SHALL use the established plain card treatment, a small 🍰 emoji beside its label and a labeled keyboard-operable expansion control. It SHALL have no responsible-team field. Administrators SHALL be offered controls for the date's delivery time and quantity; delivery time entry SHALL consistently use 24-hour `HH:MM` regardless of the browser locale. Other viewers SHALL see those settings without edit controls. An unconfigured block SHALL visibly require administrator setup and SHALL not offer claim controls.

#### Scenario: Configured cake block is expanded

- **WHEN** a viewer expands a cake block configured for four cakes at 10:00
- **THEN** the card shows that delivery time and quantity
- **AND** exactly four independently displayed cake positions are shown

#### Scenario: Cake block needs setup

- **WHEN** a cake block has no delivery time or requested quantity
- **THEN** it shows that administrator setup is needed
- **AND** no invented time, quantity or claimable volunteer positions are shown

#### Scenario: Administrator adjusts delivery settings

- **WHEN** an administrator saves valid new delivery settings
- **THEN** the card shows the saved time and requested quantity
- **AND** its position count and progress reflect the saved configuration

### Requirement: Cake cards follow day-block filtering and privacy

Cake cards SHALL participate in date and assigned-person filtering under the same rules as other day-task cards. A playing-team filter SHALL retain a cake card only for a date with a qualifying game. A responsible-team filter SHALL hide the cake card because it has no responsible team. A person-name filter SHALL retain the cake card only when one of its own assignments matches, and SHALL allow it to retain a qualifying date without a matching game card. Clearing filters SHALL restore cake cards in their normal upcoming or past date groups.

Guests SHALL see assigned names without roster payloads, person identifiers, contact data or assignment/configuration controls. Editable cake cards SHALL load their permitted candidates only when opened, retaining the existing conflict and candidate-failure behavior.

#### Scenario: Find a cake volunteer

- **WHEN** a person-name filter matches only a cake volunteer on a qualifying date
- **THEN** the date and cake card remain visible
- **AND** unrelated games and other day blocks are not shown merely because they share that date

#### Scenario: Responsible-team filter is selected

- **WHEN** a viewer selects a responsible-team filter
- **THEN** cake cards are hidden

#### Scenario: Guest views cake assignments

- **WHEN** a guest expands a configured cake card
- **THEN** the assigned helper names are visible
- **AND** the response contains no roster payload, person identifiers, contacts or editing controls

#### Scenario: Open an editable cake card

- **WHEN** a signed-in viewer opens an editable cake card
- **THEN** only that card loads its permitted candidates
- **AND** a failed request leaves its controls unable to submit claims based on an incomplete candidate list
