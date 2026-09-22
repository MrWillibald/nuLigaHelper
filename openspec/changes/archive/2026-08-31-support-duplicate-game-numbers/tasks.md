## 1. Scraped Source Identity

- [x] 1.1 Add `source_key` to the scraper output contract and extract a stable nuLiga
      meeting identifier from row/link metadata when available; verify an offline HTML
      fixture with duplicate game numbers produces two distinct `meeting:` keys
- [x] 1.2 Add the normalized fallback key over game number, age class, home and guest for
      rows without an upstream identifier; verify date, time, hall and score changes do
      not change that fallback key
- [x] 1.3 Validate source-key uniqueness before returning scraped games and raise a
      diagnostic containing both conflicting rows; verify an offline fixture with an
      unresolved fallback collision fails without returning partial data

## 2. Database Identity And Sync

- [x] 2.1 Add non-null `Game.source_key`, replace the `(season_year, game_nr)` uniqueness
      constraint with `(season_year, source_key)`, and update synthetic game fixtures;
      verify two games in one season can share a number but not a source key
- [x] 2.2 Rework `sync_games()` to validate and index incoming/stored games by source key;
      verify duplicate-number games are both inserted and matched to the same local rows
      on a repeated scrape
- [x] 2.3 Preserve assignments and responsible teams when mutable fields change on one
      duplicate-number game; verify a date/time shift updates only that game and retains
      its related records
- [x] 2.4 Change new, removed, shift and referee event records to carry local `game_id`
      plus display `game_nr`, and compare removals by source key; verify only the affected
      duplicate-number game appears in each event scenario
- [x] 2.5 Make sync collision validation happen before ORM mutation/commit; verify a
      rejected duplicate source key leaves the previously stored game set unchanged

## 3. Event Consumers

- [x] 3.1 Update notifier shift and referee handlers to resolve the exact game by event
      `game_id`; verify duplicate-number games notify only the helpers and responsible
      parties attached to the affected game
- [x] 3.2 Update `main.py` new-game filtering and reporting so records are associated by
      source/local identity instead of a `game_nr` dictionary; verify two new games with
      one number are independently included or excluded by their own age class
- [x] 3.3 Remove or narrow remaining domain lookups that accept game number as identity;
      verify a repository search leaves `game_nr` only in display, filtering and readable
      event fields rather than exact-game lookup paths

## 4. Unambiguous Administration

- [x] 4.1 Print internal game IDs with number, date, time, age class and matchup in
      `manage_db.py list-games`, and optionally filter the list by display number; verify
      two same-number games are both shown distinguishably
- [x] 4.2 Change CLI `assign`, `unassign` and `set-jteam` to accept `game_id`; verify each
      command mutates only the selected duplicate-number game and update CLI help examples
- [x] 4.3 Expand admin game-picker labels, including the audit filter, with date, time,
      age class and matchup while continuing to submit internal IDs; verify duplicate
      numbers render as distinct selectable options without exposing IDs on the public
      schedule

## 5. Regression Coverage And Documentation

- [x] 5.1 Add a database scenario covering insert, repeat sync, one-game shift, one-game
      referee transition and one-game removal for duplicate numbers; verify the test passes
      standalone and under pytest
- [x] 5.2 Extend notifier, CLI and web scenarios for duplicate-number identity isolation;
      verify every affected test file passes standalone
- [x] 5.3 Update `AGENTS.md` and `README.MD` to state that `Game.id`/`source_key` are identity,
      `game_nr` may repeat, and CLI mutations take game IDs; verify documented commands
      match `manage_db.py --help`
- [x] 5.4 Run `test/run_tests.sh` and `pytest test/ -v`, confirm all fixtures remain offline
      and synthetic, and verify no test reads `config.json` or the real database
