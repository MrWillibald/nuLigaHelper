# Proposal

## Why

The repository root holds 25 Python modules side by side with shell scripts, requirement files, Alembic configuration and five documentation files. Nothing tells a reader which files belong to the web interface, the daily job, production operations or the domain layer, and `test/` holds 80 files in one flat directory. After `split-webapp-into-modules` and `split-domain-module` the code is already divided by topic; this change groups it into one application package with four clearly named areas and mirrors that structure in tests and documentation (phases 3 and 4 of `docs/REFACTORING-VORSCHLAG.md`).

## What Changes

- Create the package `nuligahelper/` and move code with `git mv` (history preserved):
  - `nuligahelper/config.py` ← `common.py`; `nuligahelper/cli.py` ← `manage_db.py`.
  - `nuligahelper/db/` ← `db/` package plus `age_eligibility.py`, `contact_validation.py` (as `contacts.py`), `task_labels.py` (folded into `roles.py`), `schema_migrations.py` (as `schema_upgrade.py`) and `migrations/`.
  - `nuligahelper/web/` ← `webapp/` package plus `auth_abuse.py`, `templates/`, `static/`.
  - `nuligahelper/jobs/` ← `main.py` (as `daily.py`), `daily_lock.py`, `scraper.py`, `notifier.py`, `messages.py`, `message_config.py`.
  - `nuligahelper/ops/` ← `backup.py`, `production.py` (as `settings.py`), `production_logging.py` (as `journal_logging.py`), `operations.py` (as `monitoring.py`), `monitor_watchdog.py` (as `watchdog.py`), `runtime_check.py`, `privacy.py`.
- Entry points become module invocations: `python -m nuligahelper.jobs.daily`, `python -m nuligahelper.web` (local server), `gunicorn nuligahelper.web:app`, `python -m nuligahelper.cli`, `python -m nuligahelper.ops.monitoring …`, etc. `run.sh`, `run_webapp.sh`, `alembic.ini` (`script_location`), systemd units, the watchdog cron entry and the deploy script's unit verification are updated.
- **Transition release:** thin root shims (`main.py`, `webapp.py`, `manage_db.py`, `operations.py`, `privacy.py`, `runtime_check.py`, `monitor_watchdog.py`, `schema_migrations.py`) forward to the new modules so already installed systemd units — and the already installed deploy command, whose schema-head probe runs `import schema_migrations` inside the candidate release — keep working until the operator installs the updated units and deploy script. A follow-up change removes the shims.
- Tests move to `tests/` with sub-directories `db/`, `web/`, `jobs/`, `ops/`, `migrations/`, `js/`; `run_tests.sh` and CI discover tests recursively.
- Documentation: `docs/benutzer/` (manuals, screenshots, `pdf/` ← `output/pdf/`), `docs/betrieb/` (`DEPLOYMENT.md`, `DATA-INVENTORY.md`, operator topics), `docs/entwicklung/` (architecture, tests); `README.MD` → `README.md`.
- `release-assets/` keeps its name and path (installed hosts, deploy script and recovery tooling reference it; `deploy/` is already reserved for private, gitignored host material).

## Capabilities

### New Capabilities

None.

### Modified Capabilities

None. File locations and invocation commands change; application behavior, data, schema and deployment workflow guarantees stay the same.

## Impact

- Every Python import (`import db` → `from nuligahelper import db`, etc.), all test imports and patch targets (`patch("common.effective_today")` → `patch("nuligahelper.config.effective_today")`, …).
- `release-assets/systemd/*.service`, `release-assets/nuligahelper-watchdog.cron`, `release-assets/gunicorn.conf.py` reference, `release-assets/nuligahelper-deploy.py` (unit path checks, schema-head probe `import schema_migrations`), `test/test_deployment_assets.py`, `test/test_release_prepare.py`, `test/test_host_operations.py`.
- `docs/build_handbook.py` output path, README/manual links, `AGENTS.md`, `CLAUDE.md`, `test/README.md`, `.github/workflows/offline-tests.yml`.
- Operators must install the updated systemd units once (documented in `docs/betrieb/deployment.md`); until then the shims keep the old units working.
- Depends on `remove-legacy-game-identity-migration`, `split-webapp-into-modules` and `split-domain-module`.
