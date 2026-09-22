## Why

The current source-key identity can create duplicate database rows when nuLiga changes
or omits its meeting identifier even though the visible game number still denotes the
same game. Spielfeste are the sole valid exception to unique scraped game numbers: their
individual SPF matches reuse small numbers across dates but require only one shared set
of helper tasks per date and age group.

## What Changes

- Identify ordinary games within a season exclusively by their scraped game number;
  nuLiga meeting IDs, matchup data and derived source fingerprints no longer determine
  identity.
- Collapse all SPF rows with the same date and full age group into one task-relevant
  Spielfest entry in the plan.
- Give each collapsed Spielfest a deterministic, human-readable pseudo game number
  derived from its date and age group. A moved Spielfest intentionally receives a new
  identity and does not retain assignments automatically.
- Derive the Spielfest plan entry from the grouped rows, using the earliest start time
  and their common scheduling data, and reject internally inconsistent groups instead
  of selecting arbitrary values.
- **BREAKING**: Store canonical game numbers as text so ordinary numeric values and SPF
  pseudo numbers use the same identity field, and enforce uniqueness by season and
  canonical game number instead of by source key.
- Add a reviewed, backup-first SQLite migration that preserves existing people,
  ordinary-game assignments and audit history, collapses existing SPF rows, and aborts
  rather than guessing when duplicate ordinary games or SPF task data conflict.
- Update plan presentation, administrative selection, notifications, documentation and
  offline tests for the unified ordinary-game/Spielfest model.

## Capabilities

### New Capabilities

<!-- None. -->

### Modified Capabilities

- `game-identity`: Replace source-key identity and duplicate-number coexistence with
  season-scoped canonical game-number identity and SPF aggregation into one
  task-relevant plan entry.

## Impact

- `scraper.py` changes its output normalization and collapses SPF rows before sync.
- `db.py` changes the game schema, uniqueness constraint, synchronization keys and
  migration/reconciliation behavior.
- `main.py`, `notifier.py`, `manage_db.py`, `webapp.py`, templates and sorting/filtering
  paths must accept textual canonical game numbers and synthetic Spielfest entries.
- Existing SQLite installations require a controlled schema and data migration rather
  than database recreation.
- Scraper, database, notifier, CLI and web tests must replace duplicate-number fixtures
  with ordinary identity-continuity and SPF aggregation/migration scenarios.
- `README.MD` and `AGENTS.md` must document the new identity rule and migration process.
