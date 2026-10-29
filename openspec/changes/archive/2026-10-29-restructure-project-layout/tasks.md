# Tasks

## 1. Code moves (one commit for all areas, full suite green, see design implementation notes)

- [x] 1.1 Create `nuligahelper/__init__.py`; move `ops` modules (`backup`, `production` → `settings`, `production_logging` → `journal_logging`, `operations` → `monitoring`, `monitor_watchdog` → `watchdog`, `runtime_check`, `privacy`) and update imports.
- [x] 1.2 Move `jobs` modules (`main` → `daily`, `daily_lock`, `scraper`, `notifier`, `messages`, `message_config`) and `common` → `nuligahelper/common.py` (name kept, see design).
- [x] 1.3 Move `db/` plus `age_eligibility`, `contact_validation` → `contacts`, `task_labels` (kept separate to avoid an import cycle), `schema_upgrade` ← `schema_migrations` and `migrations/`; update `alembic.ini` and `migrations/env.py`; verify schema-drift and migration tests and `alembic heads`.
- [x] 1.4 Move `webapp/` → `nuligahelper/web/` with `templates/`, `static/`; `auth_abuse` → `nuligahelper/auth_abuse.py`; add `__main__.py`; drop the temporary re-exports from `split-webapp-into-modules`.
- [x] 1.5 `manage_db.py` → `nuligahelper/cli.py`; verify `python -m nuligahelper.cli --help` matches the previous help text (apart from the program name).

## 2. Entry points and deployment assets

- [x] 2.1 Update `run.sh`, `run_webapp.sh`, `release-assets/systemd/*.service`, `release-assets/nuligahelper-watchdog.cron` and the Gunicorn target to `-m` / `nuligahelper.web:app`.
- [x] 2.2 Add the root compatibility shims for every path referenced by currently installed units, for `schema_migrations` (imported by the installed deploy command's schema-head probe) and `test/run_tests.sh` (run by the installed deploy command); extend `test_deployment_assets.py` to execute each shim in a synthetic release tree.
- [x] 2.3 Update `release-assets/nuligahelper-deploy.py` (`unit_paths_safe` accepts old and new unit forms; schema-head probe imports `nuligahelper.db.schema_upgrade` with a fallback for older releases; prepare runs `tests/run_tests.sh`, falling back to `test/run_tests.sh`) and `test_release_prepare.py`/`test_host_operations.py`.
- [x] 2.4 Document the one-time unit update and the shim removal plan in `docs/betrieb/deployment.md`.

## 3. Tests and documentation

- [x] 3.1 Move tests to the `tests/{db,web,jobs,ops,migrations,js}/` package, update imports, all patch targets, `run_tests.sh` and the CI workflow; add guard tests against imports or patch targets that name former top-level modules.
- [x] 3.2 Reorganize `docs/` into `benutzer/` (incl. `pdf/` from `output/pdf/`), `betrieb/` (incl. `deployment.md`, `data-inventory.md`) and `entwicklung/` with a `docs/README.md` index; update `build_handbook.py`, README and manual links; rename `README.MD` → `README.md`.
- [x] 3.3 Update `AGENTS.md`, `CLAUDE.md` and `tests/README.md` with the final layout; remove `docs/REFACTORING-VORSCHLAG.md`, whose plan is complete.

## 4. Verification

- [x] 4.1 Run `tests/run_tests.sh`, `gunicorn --check-config nuligahelper.web:app`, `python -m nuligahelper.cli --help`, and each root shim with `--help`/dry invocation; verify the root directory listing matches the target layout.
- [x] 4.2 Run `openspec validate restructure-project-layout --strict` and review that no file content changed beyond imports, paths and documentation.
