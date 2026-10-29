# Proposal

## Why

The one-time migration from the historical `games.source_key` identity to season-scoped canonical game numbers (`game_identity_migration.py`, `manage_db.py migrate-game-identity`) has been run on every deployed database. It still costs 276 lines of production code, a dedicated test file, a CLI command, README instructions and a special case in schema adoption, and it has to be carried through the planned module restructuring. Removing it now shrinks the codebase before files are moved.

## What Changes

- **BREAKING (operators only):** remove `game_identity_migration.py`, the `manage_db.py migrate-game-identity` command and `test/test_game_identity_migration.py`.
- Keep fail-closed detection of the legacy `games.source_key` schema in `schema_migrations.inspect_schema()` and `db.verify_db()`, but change the diagnostic: the database is refused unchanged and the operator is told that this release no longer converts legacy game identity and that an earlier release containing `migrate-game-identity` must be used first.
- Remove the README section "Migrate legacy game identity" and the corresponding step in the schema-migration procedure; update `AGENTS.md` accordingly.
- Canonical game-number identity, Spielfest aggregation and all other game-identity rules are unchanged.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

- `game-identity`: Remove the requirement for a backup-first source-key migration.
- `schema-migrations`: A database with the legacy game-identity schema is refused with a diagnostic pointing to an earlier release instead of to a command of this release.

## Impact

- Deleted: `game_identity_migration.py`, `test/test_game_identity_migration.py`.
- Changed: `manage_db.py` (command and import removed), `schema_migrations.py` and `db.py` (diagnostic text only), `test/test_schema_migrations.py` and `test/test_cli.py` (expected message / command list), `README.MD`, `AGENTS.md`.
- No schema revision, no data change, no deployment step. Rollback is the previous release.
- Must land before `restructure-project-layout` so the module is not moved first.
