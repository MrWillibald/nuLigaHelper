# Design

## Context

Release N (`restructure-project-layout`) shipped the package together with root shims.
The host has since completed the documented one-time transition: the units in
`/etc/systemd/system/` start `python -B -m nuligahelper.…` and Gunicorn with
`nuligahelper.web:app`, `/etc/cron.d/nuligahelper-watchdog` runs
`python -m nuligahelper.ops.watchdog` from `/opt/nuligahelper/current`, and
`/usr/local/sbin/nuligahelper-deploy` is the release-N copy. That copy already prefers
`tests/run_tests.sh` and `nuligahelper.db.schema_upgrade`, so it can prepare and activate a
release without shims.

What protects a host that missed the transition today is only the shims. Once they are
gone, `unit_paths_safe()` would still accept an old unit such as
`ExecStart=/opt/nuligahelper/current/venv/bin/python -B /opt/nuligahelper/current/main.py`
because it only checks that every path lies under `current/`. Activation would switch
`current`, the web unit would fail to start and the daily job would fail at 09:00.

## Goals / Non-Goals

**Goals:** a repository root without Python modules; one entry point per job; a deploy
command that refuses a host whose units or watchdog entry would break, before any service
is stopped; no behaviour change for a correctly transitioned host.

**Non-Goals:** installing host files from a release; supporting rollback to a pre-package
release; changing the reviewed unit contents; removing the first-cutover legacy checkout
handling in `prior_commit()`.

## Decisions

### 1. Remove all shims in one release

The shims only matter together: the units, the cron entry and the deploy command were
switched in one maintenance window. Removing them piecemeal would keep the transition
documentation alive for no benefit. `test/run_tests.sh` goes with them; the installed
deploy command no longer reaches it because `tests/run_tests.sh` exists.

### 2. Refuse old units by their command form, not by file names

`unit_paths_safe()` additionally requires, for every `ExecStart`/`ExecStartPre` line that
runs the release interpreter, the form `<current>/venv/bin/python -B -m nuligahelper.<module>`,
and for the web unit `nuligahelper.web:app` as the Gunicorn target. A line that runs a
`.py` file or `webapp:app` is refused with "installed application unit still uses
transition entry points". `verify_installed_units()` applies the same rule to the loaded
`ExecStart`/`ExecStartPre` properties, so a stale `daemon-reload` is caught as well.
The check runs where installed units are already verified today: at the start of
`activate` and `continue`, before `enter_maintenance()` stops anything. `inspect` stays
read-only and reports the same finding without failing the host plan.

Alternative considered: compare installed units byte-for-byte with `release-assets/systemd/`.
Rejected because hosts may legitimately carry reviewed local hardening lines; the command
form is what decides whether a release can start.

### 3. Check the watchdog cron entry at its documented path

The watchdog is the independent safety net for a disabled monitor schedule; a cron entry
that still calls `monitor_watchdog.py` fails silently every 15 minutes. The deploy command
reads `/etc/cron.d/nuligahelper-watchdog` (root-owned, not group/world-writable, not a
symlink) and requires exactly one active entry containing
`-m nuligahelper.ops.watchdog`. `release-assets/README.md` names this install path.
Hosts that installed the entry elsewhere move it there before deploying this release.

### 4. Drop pre-package fallbacks in the deploy command

- `prepare()` refuses a candidate without `tests/run_tests.sh` or `nuligahelper/__init__.py`
  with "selected commit lacks required release assets" and never runs `test/run_tests.sh`.
- `previous_schema_head()` imports only `nuligahelper.db.schema_upgrade`.
- `verify_previous_web_compatibility()` additionally requires
  `<previous>/nuligahelper/web/__init__.py`; a previous release without the package cannot
  be started by the installed units, so code-only rollback to it must not be offered.

The release-N deploy command keeps working with this release, so installing the updated
command is a follow-up host step rather than a precondition.

## Risks / Trade-offs

- **A host without the transition breaks** → the updated deploy command refuses activation
  before stopping services; the release-N deploy command does not, so the deployment guide
  keeps the prerequisite check (`systemctl show … -p ExecStart`, cron entry) before
  activating this release with an older deploy command.
- **Rollback to a pre-package release is impossible** → accepted; such a release cannot
  start under the installed units anyway, and the transition is complete. Recovery to such
  a release would need the previous unit files, which the private runbook keeps.
- **Fixed cron path** → a host that keeps the entry elsewhere is refused with an actionable
  message instead of silently losing the watchdog.
