# Spec Delta

## ADDED Requirements

### Requirement: Activation requires package entry points on the host

Before stopping any writer or ingress for activation or continuation, the deployment workflow SHALL verify that every installed and loaded application unit starts the selected release through package module entry points and that the watchdog cron entry at its documented path runs the package watchdog module. It SHALL refuse with an actionable status, leaving services and timers unchanged, when a unit or the watchdog entry still invokes a former root script or Gunicorn target.

#### Scenario: Host runs the package units

- **WHEN** the operator activates a prepared release and every installed and loaded unit uses `python -B -m nuligahelper.…` or `nuligahelper.web:app`, and the watchdog entry runs `nuligahelper.ops.watchdog`
- **THEN** activation proceeds with the existing quiescence, snapshot and schema gates

#### Scenario: Installed unit still calls a root script

- **WHEN** an installed or loaded application unit still runs a root script such as `main.py` or starts Gunicorn with `webapp:app`
- **THEN** activation and continuation refuse before pausing timers or stopping the web service
- **AND** the refusal names the transition entry points without printing unit contents or environment values

#### Scenario: Watchdog entry is missing or outdated

- **WHEN** the watchdog cron entry is absent, unsafe to read or still calls `monitor_watchdog.py`
- **THEN** activation refuses before any service interruption

### Requirement: Releases contain the application package

The repository SHALL NOT ship root-level transition entry points for the former module layout. Candidate preparation SHALL refuse a commit without the application package or its offline test runner. Code-only rollback SHALL be offered only to a previous release that contains the application package.

#### Scenario: Candidate predates the package layout

- **WHEN** the operator prepares a commit without the `nuligahelper` package or `tests/run_tests.sh`
- **THEN** preparation refuses it without changing the active release, database or services

#### Scenario: Previous release predates the package layout

- **WHEN** the retained previous release does not contain the `nuligahelper` package
- **THEN** activation refuses because a code-only rollback to that release could not start under the installed units
