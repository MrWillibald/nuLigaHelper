# Tasks

## 1. Deploy command refusals (before removing anything)

- [x] 1.1 Extend `unit_paths_safe()` and `verify_installed_units()` in `release-assets/nuligahelper-deploy.py` to require `-B -m nuligahelper.` for every release-interpreter `ExecStart`/`ExecStartPre` line and `nuligahelper.web:app` for the web unit, as in design.md decision 2. Verify in `tests/ops/test_release_prepare.py` that every reviewed unit in `release-assets/systemd/` passes and that variants running `main.py`, `runtime_check.py` or `webapp:app` are refused, both in file content and in loaded properties.
- [x] 1.2 Add the watchdog entry check for `/etc/cron.d/nuligahelper-watchdog` (decision 3) to `activate` and `continue`, and report it read-only from `inspect`. Verify with synthetic files in `tests/ops/test_release_prepare.py`: the reviewed `release-assets/nuligahelper-watchdog.cron` passes; missing, symlinked, group-writable, `monitor_watchdog.py` and duplicated entries are refused before `enter_maintenance()` runs.
- [x] 1.3 Drop the pre-package fallbacks (decision 4): `prepare()` requires `tests/run_tests.sh` and `nuligahelper/__init__.py`, `previous_schema_head()` imports only `nuligahelper.db.schema_upgrade`, `verify_previous_web_compatibility()` requires `nuligahelper/web/__init__.py`. Move the synthetic release fixture from `test/run_tests.sh` to `tests/run_tests.sh` and add a package marker; verify a pre-package candidate and a pre-package previous release are refused.

## 2. Remove the shims

- [x] 2.1 Delete `main.py`, `webapp.py`, `manage_db.py`, `operations.py`, `privacy.py`, `runtime_check.py`, `monitor_watchdog.py`, `schema_migrations.py` and `test/run_tests.sh` with `git rm`.
- [x] 2.2 In `tests/ops/test_deployment_assets.py`, remove `SHIMS` and `test_transition_shims_work_from_a_release_tree`, start Gunicorn only with `nuligahelper.web:app` and stop copying `webapp.py` into the synthetic release. Verify the real-Gunicorn test still serves `/healthz` through the redacting formatter.
- [x] 2.3 In `tests/ops/test_project_layout.py`, replace `ROOT_SHIMS` with an assertion that the repository root holds no `*.py` file and that `test/` does not exist; keep `OLD_MODULES` for the import and patch guards. Verify with `grep -rnE "schema_migrations|webapp:app|test/run_tests|monitor_watchdog\.py" --include=*.py --include=*.sh --include=*.md --include=*.yml . | grep -v openspec/changes/archive` that only intended historical mentions remain.

## 3. Documentation

- [x] 3.1 Replace "One-time transition to the package layout" in `docs/betrieb/deployment.md` with a short prerequisite: the host runs the reviewed units and watchdog entry, how to check it (`systemctl show nuligahelper-web.service nuligahelper-daily.service -p ExecStart`, the cron file) and that the updated deploy command refuses otherwise; keep the hint to install the updated deploy command after activating this release.
- [x] 3.2 In `release-assets/README.md`, remove the legacy `deploy/gunicorn.conf.py` compatibility paragraph and name `/etc/cron.d/nuligahelper-watchdog` as the install path of `nuligahelper-watchdog.cron`.
- [x] 3.3 Remove the shim notes from `AGENTS.md` (module map introduction), `README.md` (repository layout) and `tests/README.md`.

## 4. Verification

- [x] 4.1 Run `tests/run_tests.sh` and verify the whole suite is green; run `openspec validate remove-transition-shims --strict` and verify it passes.
- [x] 4.2 Verify from a clean export of the branch that `python -B -m nuligahelper.cli --help`, `python -B -m nuligahelper.ops.runtime_check` (expects the secret error) and `gunicorn --check-config nuligahelper.web:app` work and that no root `*.py` file remains.
