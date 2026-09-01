## Context

See `proposal.md` — Why. `Game.id` already provides correct local relational identity,
and web assignment endpoints already use it. The defect is at the scrape/sync boundary:
`Game` is unique on `(season_year, game_nr)`, `sync_games()` indexes both stored and
scraped rows by `game_nr`, and notification events later look games up by that same
display number. The scraper currently returns only visible table cells and discards any
row-link identity present in the nuLiga HTML.

Dates and times are mutable and deliberately remain German strings. They cannot be part
of stable identity because shifts must preserve assignments. There is no migration layer;
the deployment procedure may recreate the database after this schema change.

## Goals / Non-Goals

**Goals:**

- Give every accepted scrape row a deterministic source key that distinguishes valid
  duplicate game numbers.
- Preserve local `Game.id` and every relationship attached to it across ordinary nuLiga
  field changes.
- Make all sync events and administrative selection paths unambiguous.
- Fail before commit when source data cannot be distinguished safely.

**Non-Goals:**

- Deduplicating genuinely repeated HTML rows that nuLiga itself cannot distinguish.
- Introducing database migrations or preserving the current development database.
- Changing public URLs to expose source or internal game identifiers.
- Changing assignment, authorization or notification timing rules.

## Decisions

### D1 — Add a stable `source_key`; keep `game_nr` as display data

`Game` receives a non-null textual `source_key`, unique together with `season_year`.
The existing `(season_year, game_nr)` unique constraint is removed. `game_nr` remains an
integer because it is useful in schedules, logs and messages, but no lookup or event uses
it as the sole key.

The scraper adds `source_key` to its output contract. The preferred key is nuLiga's own
meeting identifier extracted from the row or its links and namespaced, for example
`meeting:<value>`. If no upstream identifier is present, the scraper emits a namespaced,
deterministic fingerprint over normalized fields expected to survive rescheduling:
`game_nr`, `ak`, `home` and `guest`. Date, time, hall and score are excluded.

Before returning, the scraper validates that all source keys in the filtered home-game
set are unique. A fallback collision raises a diagnostic containing both rows; adding an
occurrence counter would appear to work but would attach assignments to a different game
when row ordering changes.

*Alternative considered:* use `(game_nr, date, time)` as identity. Rejected because the
normal shift workflow would create a new game and orphan its assignments.

*Alternative considered:* use `(game_nr, ak, home, guest)` directly as the database
constraint. Rejected because a namespaced source key can use a stronger upstream ID when
available and keeps source-specific logic in the scraper.

### D2 — Sync and events carry source/local identity end to end

`sync_games()` builds its existing-game map by `source_key`, validates all incoming keys
before mutating ORM objects, and reports event records with the affected local `game_id`
and display `game_nr`. Removed-game detection compares source-key sets.

Notifier lookups use the event's `game_id`; `get_game()` becomes an ID-based helper or is
replaced by `session.get()`. Keeping `game_nr` in event payloads preserves readable logs
without making it an identity key. `main.py` associates newly discovered records by
source key or local ID rather than a dictionary keyed by game number.

*Alternative considered:* have notifier events carry only `source_key`. Rejected because
the local primary key is already the canonical reference for assignments and avoids
repeating season/source lookup logic.

### D3 — CLI mutations use internal game IDs

`list-games` prints `ID <id>` plus number, date, time, age class and matchup. Commands that
mutate an existing game (`assign`, `unassign`, `set-jteam`) accept `game_id`, matching the
existing person-ID convention. An optional game search/list filter may still accept a
display number because returning multiple contextualized results is not ambiguous.

Web JSON routes already use `Game.id`. Audit and other admin selectors continue posting
IDs but expand labels so repeated numbers are distinguishable. Public schedule cards do
not need to expose either ID as user-facing text.

### D4 — Source identity is immutable after insertion

Sync uses the incoming key only to locate the existing row and never rewrites a stored
`source_key`. If nuLiga changes an upstream meeting identifier, the row is treated as a
new game and the old one as removed, producing a visible diagnostic rather than guessing
that assignments should move. A future reconciliation tool can address a demonstrated
upstream-ID churn pattern if one appears.

## Risks / Trade-offs

- **nuLiga table parsing may not expose a meeting identifier through `pandas.read_html`**
  → inspect row links from the raw HTML; retain the deterministic fingerprint fallback.
- **A team-name correction changes a fallback fingerprint** → report a removed and new
  game rather than silently transferring assignments; prefer upstream IDs whenever
  available.
- **Two fallback rows may share number, age class and matchup** → abort before commit and
  log both rows so the source-key extractor can be extended from real evidence.
- **CLI command syntax changes** → update help, README and tests together; IDs are visible
  in `list-games` and game search output.
- **Existing database lacks `source_key`** → recreate it according to the project's
  explicit no-migration policy.

## Migration Plan

1. Deploy code and documentation together.
2. Delete the existing SQLite database after taking any desired backup; no migration is
   introduced.
3. Run the real scrape once and confirm duplicate-number tournament games are stored as
   separate rows with distinct source keys.
4. Re-enter responsible teams and assignments as needed, using IDs from `list-games`.
5. Roll back by reverting code and recreating the database with the prior schema.
