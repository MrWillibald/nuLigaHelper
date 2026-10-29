# Implementation verification

Verified on 2026-10-29 using synthetic databases only.

- Deleted `game_identity_migration.py` and `test/test_game_identity_migration.py`;
  removed the import, `cmd_migrate_game_identity` and the `migrate-game-identity`
  sub-parser from `manage_db.py`. `manage_db.py --help` no longer lists the command.
  `grep -rn game_identity_migration` finds no reference outside the archive and this
  change. `test/test_cli.py` asserts no command list, so it needed no change.
- `schema_migrations.inspect_schema()` still classifies `games.source_key` databases as
  `legacy_game_identity`. `migrate_to_head()` and `db.verify_db()` refuse them with a
  hint to use an earlier release and no longer name the removed command.
- `test/test_schema_migrations.py`: the CLI refusal test now asserts the release hint,
  the absence of the removed command and a byte-identical database; the new
  `test_legacy_game_identity_startup_refuses_unchanged_with_release_hint` replaces the
  startup test from the deleted file (German hint, no command name, unchanged file).
- Documentation: the legacy procedure in `docs/betrieb/schema-migration.md` became a
  short "Legacy game identity" note and the upgrade steps were renumbered; `AGENTS.md`
  module map and schema rule, `docs/betrieb/README.md` and
  `docs/REFACTORING-VORSCHLAG.md` were updated.
- `test/run_tests.sh`: 480 passing checks (486 before, minus the 7 tests of the deleted
  file, plus 1 new startup test). The only failure is the known host-disk dependency in
  `test_monitor_detects_staleness_missing_timers_clock_disk_and_tls`, unchanged and
  addressed by `tidy-maintenance-helpers`.
- `openspec validate remove-legacy-game-identity-migration --strict`: valid.
