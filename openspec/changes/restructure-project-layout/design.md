# Design

## Context

The application runs from a checkout without packaging: `run.sh` executes `python main.py`, systemd units execute `/opt/nuligahelper/current/<module>.py` with `WorkingDirectory=/opt/nuligahelper/current`, Gunicorn loads `webapp:app`, Alembic finds `migrations/` via `alembic.ini`, and tests put the project root on `sys.path` through `test/helpers.py`. The deploy script verifies that installed units reference `/opt/nuligahelper/current/` and probes the schema head with `import schema_migrations`. Units are installed by the operator, not by the deploy command, so a release cannot change them on its own.

## Goals / Non-Goals

**Goals:** a root directory with ~15 entries; four obvious code areas (`db`, `web`, `jobs`, `ops`); tests and docs mirroring them; no behavior change; an upgrade path where the new release runs under old *and* new units.

**Non-Goals:** no `pyproject.toml`/installable package, no `src/` layout, no renaming of `release-assets/`, no change to systemd unit names, users, timers or hardening, no splitting of `static/app.js` (can follow separately).

## Decisions

### 1. Plain package in the checkout, run with `-m`

`nuligahelper/` sits at the repository root. With `WorkingDirectory` (systemd) or the shell's cwd (`run.sh`) at the checkout root, `python -m nuligahelper.jobs.daily` resolves without installation. A `src/` layout or packaging would require a build/install step in the deploy command for no benefit to a single-host application.

### 2. Names describe responsibility

`common` → `config`, `production` → `ops.settings`, `operations` → `ops.monitoring`, `monitor_watchdog` → `ops.watchdog`, `main` → `jobs.daily`, `manage_db` → `cli`, `schema_migrations` → `db.schema_upgrade`, `contact_validation` → `db.contacts`. `production_logging` becomes `ops.journal_logging` (not `ops.logging`, which would shadow the standard library inside the package for relative imports and confuse readers).

### 3. Two-step unit migration with shims

Release N (this change) contains the package **and** root shims, e.g. `main.py`:

```python
"""Compatibility entry point; installed units call this path until updated."""
from nuligahelper.jobs.daily import main
raise SystemExit(main())
```

`release-assets/systemd/` already contains the new `-m` units. The operator activates release N with the old units (shims), then installs the new units in a maintenance window following `docs/betrieb/deployment.md`. Release N+1 removes the shims after confirming the host runs the new units. `verify_installed_units()` accepts both forms in release N and only the new form in N+1.

### 4. Tests mirror the code

`tests/<area>/test_*.py`; shared `tests/helpers.py` and fixtures stay at `tests/`. Each test file keeps its standalone `if __name__ == "__main__"` runner, and `run_tests.sh` iterates `tests/**/test_*.py` (bash `globstar`). The JavaScript checks move to `tests/js/` with their path constants updated. Order-dependent webapp scenarios move as whole files.

### 5. Import style

Inside the package: absolute imports (`from nuligahelper import db, config`) and call-through-module for patchable functions (`config.effective_today()`, `db.clock.duty_now()`), so test patch targets are unambiguous: `nuligahelper.config.effective_today`, `nuligahelper.db.clock.duty_now`.

## Risks / Trade-offs

- **Old units break if shims are missing or wrong** → `test_deployment_assets.py` executes every shim path in a synthetic release directory and asserts it imports the intended module; the deploy rehearsal in `DEPLOYMENT.md` covers both unit forms.
- **Large diff hides mistakes** → one commit per area (`ops`, `jobs`, `db`, `web`, tests, docs), each green; `git mv` keeps blame; no content edits mixed into move commits beyond import lines.
- **Patch targets missed** (52 `common.effective_today` patches) → a guard test fails if any `patch("common.` or `patch("db.` string remains in `tests/`.
- **Alembic path** → `alembic.ini` `script_location = %(here)s/nuligahelper/db/migrations`; existing revision IDs and files are unchanged, so deployed databases need no migration.
