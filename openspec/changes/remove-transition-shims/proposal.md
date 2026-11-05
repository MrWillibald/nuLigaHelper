# Proposal

## Why

`restructure-project-layout` moved the application into the `nuligahelper/` package and
kept thin transition shims so that already installed systemd units, the watchdog cron
entry and the installed deploy command kept working: eight root modules (`main.py`,
`webapp.py`, `manage_db.py`, `operations.py`, `privacy.py`, `runtime_check.py`,
`monitor_watchdog.py`, `schema_migrations.py`) and `test/run_tests.sh`. Its design
planned a follow-up release that removes them once every host runs the updated units.

The production host now runs the package-layout release with the reviewed units from
`release-assets/systemd/`, the updated watchdog cron entry and the updated deploy
command, and the monitor runs cleanly. The shims are dead weight: they keep two entry
points per job, two Gunicorn targets and fallback branches in the deploy command, and
they invite new code or docs to use the old paths again.

## What Changes

- **BREAKING (operations):** delete the eight root shims and `test/run_tests.sh` (and the
  then empty `test/` directory). Units, cron entries or commands that still call a root
  script or `webapp:app` stop working with this release.
- The deploy command refuses activation (and `continue`) while an installed application
  unit still starts a root script or Gunicorn with `webapp:app`, instead of switching to a
  release that cannot start under those units. It also refuses while the watchdog cron
  entry at `/etc/cron.d/nuligahelper-watchdog` is missing or still calls a root script.
- The deploy command drops its pre-package fallbacks: `prepare` requires
  `tests/run_tests.sh` and the `nuligahelper/` package in the candidate; the previous
  release's schema head is read only through `nuligahelper.db.schema_upgrade`; rollback
  compatibility requires the previous release to contain the package, because the
  installed units cannot start a release without it.
- Tests: remove the shim execution test and the `webapp:app` Gunicorn target, turn the
  layout guard into "no Python modules at the repository root, no `test/` directory",
  move the synthetic release fixtures to `tests/run_tests.sh` and add refusal tests for
  old-form units, a missing or old watchdog entry, a pre-package candidate and a
  pre-package previous release.
- Documentation: replace the one-time transition section in `docs/betrieb/deployment.md`
  with the prerequisite and the refusal; drop the legacy Gunicorn compatibility paragraph
  and name the watchdog install path in `release-assets/README.md`; remove the shim notes
  from `AGENTS.md`, `README.md` and `tests/README.md`.

Non-goals: no change to unit names, users, timers, hardening or `ExecStart` targets of the
reviewed units; no change to application behaviour, schema or data; no removal of the
first-cutover legacy Git checkout identity in `prior_commit()` (a separate cleanup); no
automatic installation of host files by the deploy command.

## Capabilities

### New Capabilities

_None._

### Modified Capabilities

- `production-deployment`: activation and rollback require the package entry points in the
  installed units, the watchdog entry and the retained previous release.

## Impact

- Deleted: `main.py`, `webapp.py`, `manage_db.py`, `operations.py`, `privacy.py`,
  `runtime_check.py`, `monitor_watchdog.py`, `schema_migrations.py`, `test/run_tests.sh`.
- `release-assets/nuligahelper-deploy.py`: `unit_paths_safe()`, `verify_installed_units()`,
  new watchdog entry check, `prepare()`, `previous_schema_head()`,
  `verify_previous_web_compatibility()`.
- `tests/ops/test_deployment_assets.py`, `tests/ops/test_project_layout.py`,
  `tests/ops/test_release_prepare.py`, `tests/ops/test_host_operations.py`.
- Docs: `docs/betrieb/deployment.md`, `release-assets/README.md`, `AGENTS.md`, `README.md`,
  `tests/README.md`.
- Operators: no host change is required to deploy this release on a host that already runs
  the updated units, cron entry and deploy command. The deploy command installed from the
  previous release can activate it; installing the updated deploy command afterwards adds
  the new refusals. A release from before the package layout can no longer be selected for
  rollback.
- Depends on the archived `restructure-project-layout` change and on every host having
  completed its one-time unit transition.
