# Tasks

## 1. Remove the converter

- [ ] 1.1 Delete `game_identity_migration.py` and `test/test_game_identity_migration.py`; remove the import, `cmd_migrate_game_identity` and the `migrate-game-identity` sub-parser from `manage_db.py`. Verify `python manage_db.py --help` no longer lists the command and `grep -rn game_identity_migration` finds no remaining reference outside `openspec/changes/archive`.
- [ ] 1.2 Keep `SchemaState("legacy_game_identity")` detection and the refusal paths in `schema_migrations.migrate_to_head()` and `db.verify_db()`; replace both diagnostics with the "use an earlier release" wording (German for `db.verify_db()` as today). Verify the database file is byte-identical after a refused run.

## 2. Tests and documentation

- [ ] 2.1 Update `test/test_schema_migrations.py` (legacy-schema refusal no longer expects `migrate-game-identity`, asserts the new hint and an unchanged file) and `test/test_cli.py` where the command list is asserted.
- [ ] 2.2 Remove the README section "Migrate legacy game identity" and step 2 of the schema-initialization procedure that references it; update the schema-migration rule in `AGENTS.md`.
- [ ] 2.3 Run `test/run_tests.sh` and `openspec validate remove-legacy-game-identity-migration --strict`.
