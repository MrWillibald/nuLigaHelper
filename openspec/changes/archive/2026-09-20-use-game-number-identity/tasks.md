## 1. Scraper identity and Spielfest normalization

- [x] 1.1 Replace the scraper output contract's `source_key` with a canonical textual
  game number and remove meeting-link/fingerprint identity extraction; verify ordinary
  rows retain their scraped decimal number in offline parser tests.
- [x] 1.2 Add case-insensitive SPF recognition, full age-group normalization and
  deterministic `SPF:<ISO-date>:<normalized-age-group>` pseudo-number generation;
  verify repeated parsing produces identical pseudo numbers.
- [x] 1.3 Collapse SPF rows by date and full normalized age group, deriving the earliest
  time and common event fields; verify seven individual matches become one plan game and
  reused match numbers on another date produce a second plan game.
- [x] 1.4 Validate SPF dates, times and shared hall data before returning scraper output;
  verify conflicting or malformed groups fail with row context and no partial result.

## 2. Database model and synchronization

- [x] 2.1 Change `Game.game_nr` to bounded text semantics, remove `source_key`, and define
  `(season_year, game_nr)` uniqueness; verify fresh-database schema inspection reports
  the new columns and constraint.
- [x] 2.2 Rewrite sync preflight, lookup, update and removal detection around canonical
  season/game-number identity; verify source metadata and matchup changes update one
  ordinary row while an unchanged game number preserves its ID and relationships.
- [x] 2.3 Adapt sync event payloads, game sorting and number filters to textual ordinary
  and SPF values while retaining local `game_id` lookups; verify numeric ordinary order,
  Spielfest ordering and exact notification targets in database/notifier tests.
- [x] 2.4 Replace duplicate-source-key tests and fixtures with tests for canonical-number
  uniqueness, ordinary shift continuity, collapsed SPF lifecycle behavior and distinct
  reuse across seasons; verify the affected offline test modules pass.

## 3. Controlled SQLite migration

- [x] 3.1 Add read-only migration preflight and old-schema detection that reports
  duplicate ordinary season/number rows and incompatible SPF teams or assignments;
  verify synthetic conflict databases remain byte-for-byte logically unchanged.
- [x] 3.2 Implement deterministic SPF survivor selection, pseudo-number conversion and
  audit-reference reconciliation while preserving safe assignments, responsible teams
  and `Game.id`; verify synthetic migration fixtures cover empty, singly populated and
  conflicting groups.
- [x] 3.3 Implement the transactional `games` table rebuild, schema-version recording,
  foreign-key and row/relationship-count checks; verify old-schema fixtures migrate to
  the exact new schema without losing unaffected records.
- [x] 3.4 Add the explicit migration command with stopped-writer checks, a dated SQLite
  backup created through the backup API, actionable output and retry-safe failure;
  verify successful, preflight-failed and interrupted/rolled-back command tests.
- [x] 3.5 Make normal daily-job and web startup reject the legacy schema with migration
  instructions rather than changing it implicitly; verify both entry points fail closed
  on an old fixture and start normally after migration.
- [x] 3.6 Exercise the migration against a copy of the current local database and verify
  14 existing `SPF Mini` rows become two Spielfest games, all four unrelated assignments
  and four audit rows remain, and integrity checks pass; never mutate the original file.

## 4. Plan, administration and notifications

- [x] 4.1 Render each synthetic SPF game as one clearly labelled Spielfest card without
  an invented home-versus-guest matchup; verify guest and authenticated schedule views
  show one entry and one shared task set per date/age group.
- [x] 4.2 Update administrative game lists, audit filters and CLI number parsing for
  canonical textual values while mutations continue to use `Game.id`; verify ordinary
  and Spielfest selections target the intended row.
- [x] 4.3 Update helper, MV, shift and missing-slot notification formatting so a
  Spielfest message describes the aggregate event and earliest start time; verify format
  placeholder contracts and dispatch-count tests remain satisfied.
- [x] 4.4 Verify SPF aggregates do not produce individual-match referee alerts or task
  reminders and a date change is handled as removal plus addition without assignment
  transfer.

## 5. Documentation and release verification

- [x] 5.1 Update `README.MD` with the canonical game-number model, Spielfest aggregation,
  migration/backup command, downtime requirement, verification and rollback procedure;
  verify all documented commands match CLI help.
- [x] 5.2 Update `AGENTS.md` to replace the source-key domain rule, document textual
  pseudo numbers and the approved migration exception, and remove obsolete identity
  guidance; verify a repository search finds no contradictory documentation.
- [x] 5.3 Run `openspec validate use-game-number-identity --strict`, the complete offline
  `test/run_tests.sh` suite and `git diff --check`; resolve every failure and confirm both
  debug switches remain `False`.
