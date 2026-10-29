# Implementation verification

Verified on 2026-10-29 with synthetic databases only.

- Code move (`Move the application into the nuligahelper package`): all application
  modules, `migrations/`, `templates/` and `static/` moved with `git mv` into
  `nuligahelper/{db,web,jobs,ops}` plus `common.py`, `cli.py` and `auth_abuse.py`; imports and
  patch targets were rewritten mechanically. Content changes beyond imports are path
  anchoring through `PROJECT_DIR`, the Alembic `script_location`, journal logger names,
  `python -m …` hints in German error and operator messages, and the `prog` of the CLI.
- Regression fixes found during the move: `db.storage.resolve_db_path` resolves relative
  database paths against the repository root again (separate commit with test); browser
  fixtures patch `db.clock.duty_now` instead of `db.duty_now`, and the stale-patch guard
  scans `.mjs` files too.
- Tests (`Move the test suite into the tests package`): 80 files moved into
  `tests/{db,web,jobs,migrations,ops,js}`; shared helpers and fixtures stay in `tests/`. New
  guard tests reject imports and `patch(...)` targets naming former top-level modules,
  require that the root holds only the eight transition shims and that `test/` holds only
  `run_tests.sh`, execute every shim from a synthetic release tree, and check that the unit
  templates use `python -m` / `nuligahelper.web:app`.
- Runtime checks: `python -m nuligahelper.cli --db <synthetic> init` and `list-teams`;
  `alembic heads` reports the single head `0009_schedule_visibility`;
  `gunicorn --check-config --config release-assets/gunicorn.conf.py nuligahelper.web:app`
  exits 0; `manage_db.py`, `operations.py` and `privacy.py --help` exit 0 through the shims;
  `import webapp, schema_migrations` yields the app and the head revision. The help text of
  `python -m nuligahelper.cli --help` equals the previous `manage_db.py --help` except for the
  program name.
- Documentation: manuals, screenshots and PDFs are in `docs/benutzer/`, deployment and data
  inventory in `docs/betrieb/`, a new `docs/README.md` indexes all folders, and the README has
  a project layout table. `python docs/benutzer/build_handbook.py` regenerates both PDFs from
  the new paths (the committed PDFs were left unchanged). A local check of every relative
  Markdown link outside archived changes found no broken target. `openspec/specs/` names no
  former module or document path. `docs/betrieb/deployment.md` documents the one-time unit
  update, the rollback restriction after it and the shim removal.
- Root listing: `AGENTS.md`, `CLAUDE.md`, `README.md`, `alembic.ini`, `config_template.json`,
  three requirement files, `run.sh`, `run_webapp.sh`, the folders `docs/`, `nuligahelper/`,
  `openspec/`, `release-assets/`, `tests/`, `test/` (shim only) and the eight root shims.
- `tests/run_tests.sh`: exit 0, 504 passing checks, no failures. Optional browser checks
  (`tests/js/browser_*.mjs` with the preinstalled Playwright/Chromium): all four pass.
  `browser_binding_assignments.mjs` failed intermittently: after the confirmed stale claim
  returns 409, `app.js` reloads the page with feedback, and the check's next `page.goto`
  could be interrupted by that reload. The check now waits for the reload first; afterwards
  it passed 15 of 15 runs. One earlier failure, while waiting for the confirmation dialog
  right after a `reload()`, did not recur in about 40 later runs.
- `openspec validate restructure-project-layout --strict`: valid.
