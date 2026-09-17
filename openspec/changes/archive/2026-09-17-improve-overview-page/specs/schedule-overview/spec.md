## Purpose

Defines how visitors find relevant home games and distinguish completed dates from upcoming dates on the public game overview.

## ADDED Requirements

### Requirement: Overview filters games by team and assigned helper

The overview SHALL offer filters for playing team, responsible team, and assigned person. Selected filters SHALL combine so a game appears only when it satisfies every active filter. An unset responsible team SHALL not match a selected responsible team. The person filter SHALL match the names of people assigned to a game, without searching the roster or empty slots. Filtering SHALL retain the season's game ordering and SHALL show date and month headings only when they contain matching games.

#### Scenario: Combine filters

- **WHEN** a visitor selects a playing team and responsible team and enters an assigned person's name
- **THEN** the overview shows only games satisfying all three criteria
- **AND** the remaining games stay in chronological order within their date groups

#### Scenario: Responsible team is open

- **WHEN** a visitor selects a responsible team
- **THEN** games with no responsible team are excluded

#### Scenario: Duplicate names

- **WHEN** multiple people with the same displayed name hold assignments on different games
- **THEN** a name search can match each of those games without treating the name as a unique person identity

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
