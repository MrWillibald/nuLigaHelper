# Tasks

## 1. Agent and contributor guidance

- [ ] 1.1 Correct `CLAUDE.md`: replace the "no migration layer" invariant with the Alembic/`migrate-schema --confirm-stopped` rule, list the current pages and JSON endpoints from the `webapp.py` route table, and verify each named route exists with `grep -n "@app\.\(route\|get\|post\)" webapp.py`.
- [ ] 1.2 Complete the `AGENTS.md` module map with all 25 top-level modules (one-line responsibility each) and update the project summary to mention the production deployment; verify no module is missing with `ls *.py`.
- [ ] 1.3 Reduce `CLAUDE.md` to a short orientation plus `@AGENTS.md` import; move any fact that exists only in `CLAUDE.md` (architecture overview) into `AGENTS.md` first; verify no rule is lost by diffing the bullet lists.
- [ ] 1.4 Replace the per-file table in `test/README.md` with a description of the test groups; keep the run and browser-check instructions.
- [ ] 1.5 Fix the `run_webapp.sh` deployment hint and the `manage_db.py` header comment; verify `bash -n run_webapp.sh` and that the CLI help output is unchanged.

## 2. README split

- [ ] 2.1 Create `docs/betrieb/` with one file per operator topic (SQLite profile and backups, restore, schema initialization/migration, notification-wording migration, private deployment) by moving the README sections verbatim; add them to `docs/README.md`.
- [ ] 2.2 Remove the English behavior chapters from `README.MD` (web interface, home-game-day visibility, task descriptions, game duty positions, binding duties, age eligibility, cake delivery); before deleting, check each statement is covered by `openspec/specs/` or the user manuals and move operator-only facts to `docs/betrieb/`.
- [ ] 2.3 Reduce `README.MD` to overview, feature list, quick start, testing and a documentation index; keep the German user-manual links and verify every relative link (README, `docs/`, `DEPLOYMENT.md`, `DATA-INVENTORY.md`, `release-assets/README.md`) resolves to an existing file/anchor.
- [ ] 2.4 Ensure each documentation file uses a single language (German for user manuals, English for developer/operator documents).

## 3. Verification

- [ ] 3.1 Run `test/run_tests.sh` (documentation tests such as deployment-asset checks must stay green) and `openspec validate refresh-developer-documentation`.
