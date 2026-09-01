# Game Identity Specification

## Purpose

Defines stable identity and unambiguous selection for scraped games so valid meetings
with repeated display numbers coexist without losing assignments or lifecycle events.

## Requirements

### Requirement: Game number is display data rather than identity

The system SHALL allow multiple games in the same season to carry the same game number
when they represent distinct nuLiga meetings. Each such game SHALL retain its own
responsible team, task assignments, audit history and notification state.

#### Scenario: Tournament games share a number

- **WHEN** a scrape contains two distinct home games in the same season with the same
  game number
- **THEN** both games are stored and displayed
- **AND** changes to one game's assignments or responsible team do not affect the other

#### Scenario: Duplicate number persists across syncs

- **WHEN** the same two duplicate-number games appear in a later scrape
- **THEN** each scraped game updates the same stored game it represented previously
- **AND** neither game is reported as newly added merely because its number is duplicated

### Requirement: Scraped games have stable source identity

The system SHALL associate every accepted scraped game with a source identity that is
stable across changes to scheduling fields. Date, time, hall and score changes SHALL NOT
create a new game identity.

#### Scenario: Game date or time shifts

- **WHEN** a known game appears with a changed date or time but the same source identity
- **THEN** the stored game is updated rather than replaced
- **AND** its responsible team, assignments and audit references remain attached

#### Scenario: Score or hall changes

- **WHEN** a known game appears with a changed score or hall but the same source identity
- **THEN** the changed fields are synchronized onto the existing game

### Requirement: Ambiguous source identity fails safely

The system SHALL NOT merge distinct scraped rows when it cannot establish which stored
game each row represents. An unresolved source-identity collision SHALL stop the sync
with a diagnostic that identifies the conflicting rows, without committing a partial
sync.

#### Scenario: Fallback identity collides

- **WHEN** two distinct rows in one scrape produce the same fallback source identity and
  no unique nuLiga identifier distinguishes them
- **THEN** the sync is refused with a collision error
- **AND** neither row overwrites or absorbs the other

### Requirement: Game lifecycle events retain exact identity

The system SHALL carry the exact game identity through new-game, shift, missing-referee
and removed-game events. Event handling SHALL NOT resolve a game from its game number
alone.

#### Scenario: One duplicate-number game shifts

- **WHEN** only one of two games sharing a number changes date or time
- **THEN** the shift event refers to that exact game
- **AND** notifications are sent only to helpers assigned to that game

#### Scenario: One duplicate-number game disappears

- **WHEN** one of two games sharing a number is absent from a later complete scrape
- **THEN** only the absent game is reported as removed

#### Scenario: One duplicate-number game lacks a referee

- **WHEN** one of two games sharing a number transitions to a missing-referee state
- **THEN** the alert resolves the affected game and its responsible parties unambiguously

### Requirement: Existing games are selected unambiguously

Administrative interfaces SHALL identify an existing game by its stable internal
identifier rather than by game number. Lists and filters SHALL display enough scheduling
and matchup context to distinguish games that share a number, while the internal
identifier itself need not be exposed on public pages.

#### Scenario: CLI selects a duplicate-number game

- **WHEN** an administrator lists or searches games before assigning a person or setting
  the responsible team
- **THEN** each result includes its internal ID, game number, date, time, age class and
  matchup
- **AND** the mutation command accepts the selected internal ID

#### Scenario: Audit filter lists duplicate-number games

- **WHEN** an admin opens a game picker containing repeated game numbers
- **THEN** each option includes enough date, time, age-class and matchup context to select
  the intended game
