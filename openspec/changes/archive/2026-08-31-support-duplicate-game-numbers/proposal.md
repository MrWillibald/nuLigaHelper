## Why

nuLiga reuses game numbers for valid lower-age-group tournaments, while nuLigaHelper
currently treats `(season_year, game_nr)` as a unique game identity. A real scrape can
therefore fail its database sync or merge two distinct games, losing schedule entries and
making notifications or assignments refer to the wrong game.

## What Changes

- Introduce a stable source identity for each scraped nuLiga meeting, separate from the
  human-facing game number.
- Permit multiple games in one season to carry the same game number while preserving each
  game's assignments, responsible team, audit history and notification behavior.
- Match shifts, referee alerts, removals and newly discovered games by stable game
  identity instead of by game number.
- Keep the game number as display data everywhere it is useful, supplemented by date,
  time, age class and teams where users need to distinguish duplicate numbers.
- **BREAKING**: Change CLI commands that select an existing game to use its internal game
  ID; add an ID-bearing game listing/search workflow instead of accepting an ambiguous
  game number.
- Continue the project's no-migration policy: schema changes apply to newly created
  databases, and the operator recreates the current database when adopting the change.

## Capabilities

### New Capabilities

- `game-identity`: Stable scraped-game identity, duplicate game-number handling, sync
  continuity and unambiguous game selection.

### Modified Capabilities

<!-- None. Existing account, authorization, assignment and audit behavior is unchanged;
     those capabilities consume the corrected game identity internally. -->

## Impact

- `scraper.py` must retain or derive a stable identity from each nuLiga result row rather
  than returning only display fields.
- `db.py` changes the game uniqueness constraint and all sync/event lookup paths that are
  currently keyed by `game_nr`.
- `main.py` and `notifier.py` must carry stable game references through new-game, shift
  and referee-alert processing.
- `manage_db.py` must list and select games by internal ID while presenting enough context
  to distinguish duplicate numbers.
- Web and audit displays remain ID-backed but need duplicate-number disambiguation in game
  filters and labels.
- Offline scraper, database, notification, CLI and web tests gain duplicate-number and
  identity-continuity scenarios. No new runtime dependency is expected.
