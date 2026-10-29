# Tasks

## 1. Agent and contributor guidance

- [ ] 1.1 Correct `CLAUDE.md`: replace the "no migration layer" invariant with the Alembic/`migrate-schema --confirm-stopped` rule, list the current pages and JSON endpoints from the `webapp.py` route table, and verify each named route exists with `grep -n "@app\.\(route\|get\|post\)" webapp.py`.
- [ ] 1.2 Complete the `AGENTS.md` module map with all 25 top-level modules (one-line responsibility each) and update the project summary to mention the production deployment; verify no module is missing with `ls *.py`.
- [ ] 1.3 Fix the `run_webapp.sh` deployment hint and the `manage_db.py` header comment; verify `bash -n run_webapp.sh` and that the CLI help output is unchanged.

## 2. README split

- [ ] 2.1 Create `docs/betrieb/` with one file per operator topic (SQLite profile and backups, restore, schema initialization/migration, notification-wording migration, private deployment) by moving the README sections verbatim; add them to `docs/README.md`.
- [ ] 2.2 Reduce `README.MD` to overview, feature list, quick start, testing and a documentation index; keep the German user-manual links and verify every relative link (README, `docs/`, `DEPLOYMENT.md`, `release-assets/README.md`) resolves to an existing file/anchor.
- [ ] 2.3 Ensure each documentation file uses a single language (German for user manuals, English for developer/operator documents).

## 3. Verification

- [ ] 3.1 Run `test/run_tests.sh` (documentation tests such as deployment-asset checks must stay green) and `openspec validate refresh-developer-documentation`.
