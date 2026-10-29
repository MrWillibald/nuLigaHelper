# Tasks

## 1. Code moves (one commit per area, full suite green after each)

- [ ] 1.1 Create `nuligahelper/__init__.py`; move `ops` modules (`backup`, `production` → `settings`, `production_logging` → `journal_logging`, `operations` → `monitoring`, `monitor_watchdog` → `watchdog`, `runtime_check`, `privacy`) and update imports.
- [ ] 1.2 Move `jobs` modules (`main` → `daily`, `daily_lock`, `scraper`, `notifier`, `messages`, `message_config`) and `common` → `nuligahelper/config.py`.
- [ ] 1.3 Move `db/` plus `age_eligibility`, `contact_validation` → `contacts`, `task_labels` (fold into `roles`), `schema_migrations` → `schema_upgrade` and `migrations/`; update `alembic.ini` and `migrations/env.py`; verify schema-drift and migration tests and `alembic heads`.
- [ ] 1.4 Move `webapp/` → `nuligahelper/web/` with `auth_abuse`, `templates/`, `static/`; add `__main__.py`; drop the temporary re-exports from `split-webapp-into-modules`.
- [ ] 1.5 `manage_db.py` → `nuligahelper/cli.py`; verify `python -m nuligahelper.cli --help` matches the previous help text.

## 2. Entry points and deployment assets

- [ ] 2.1 Update `run.sh`, `run_webapp.sh`, `release-assets/systemd/*.service`, `release-assets/nuligahelper-watchdog.cron` and the Gunicorn target to `-m` / `nuligahelper.web:app`.
- [ ] 2.2 Add the root compatibility shims for every path referenced by currently installed units and for `schema_migrations` (imported by the installed deploy command's schema-head probe); extend `test_deployment_assets.py` to execute each shim in a synthetic release tree.
- [ ] 2.3 Update `release-assets/nuligahelper-deploy.py` (`verify_installed_units`/`unit_paths_safe` accept old and new unit forms; schema-head probe imports `nuligahelper.db.schema_upgrade`) and `test_release_prepare.py`/`test_host_operations.py`.
- [ ] 2.4 Document the one-time unit update and the shim removal plan in `docs/betrieb/deployment.md`.

## 3. Tests and documentation

- [ ] 3.1 Move tests to `tests/{db,web,jobs,ops,migrations,js}/`, update `helpers.py` path setup, all patch targets and `run_tests.sh` (recursive discovery) and the CI workflow; add the guard test against stale `patch("common.`/`patch("db.` targets.
- [ ] 3.2 Reorganize `docs/` into `benutzer/` (incl. `pdf/` from `output/pdf/`), `betrieb/` (incl. `DEPLOYMENT.md`, `DATA-INVENTORY.md`) and `entwicklung/`; update `build_handbook.py`, README and manual links; rename `README.MD` → `README.md`.
- [ ] 3.3 Update `AGENTS.md`, `CLAUDE.md`, `test/README.md` (→ `tests/README.md`) and `docs/REFACTORING-VORSCHLAG.md` with the final layout.

## 4. Verification

- [ ] 4.1 Run `tests/run_tests.sh`, `gunicorn --check-config nuligahelper.web:app`, `python -m nuligahelper.cli --help`, and each root shim with `--help`/dry invocation; verify the root directory listing matches the target layout.
- [ ] 4.2 Run `openspec validate restructure-project-layout` and review that no file content changed beyond imports, paths and documentation.
