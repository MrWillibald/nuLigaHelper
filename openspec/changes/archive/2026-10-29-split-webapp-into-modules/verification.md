# Implementation verification

Verified on 2026-10-29 with synthetic databases only.

- `webapp.py` → package `webapp/` (commit "Turn webapp.py into the webapp package"),
  suite green; then the topic split. Modules and sizes (lines): `__init__` 144 (`create_app()`
  72), `context` 155, `formatting` 57, `security` 326, `auth` 412, `challenges` 456,
  `schedule` 214, `views` 411, `assignments` 491, `persons` 416, `statistics` 202.
  No module exceeds 500 lines; no function closes over `app` (per-app resources are an
  `AppState` in `app.extensions["nuligahelper"]`).
- The split was generated from the AST of the previous `create_app()` (each statement
  moved verbatim, dedented, with `app.*` decorators mapped to a module `Registry` and
  closure state mapped to `state()`), then checked with pyflakes: no undefined or unused
  names in `webapp/`.
- Compared with the pre-split app: identical URL map (rules, endpoint names, methods),
  identical `before_request` order, error handlers and context processors.
  `test/test_webapp_structure.py` (4 tests) now pins the URL/endpoint set, the hook order,
  per-app state isolation and the request-independent `build_statistics()`.
- Before/after page comparison on identical synthetic data and a pinned date: 44
  responses (guest, member, MV, admin × `/`, `/statistik`, `/personen`, `/audit`, `/login`,
  `/registrieren`, game/block candidate APIs, `/dienst`, `/impressum` and one claim POST)
  were byte-identical after masking CSRF tokens, except the audit page's own
  `changed_at` timestamp of the claim made during the run.
- `gunicorn --check-config webapp:app` succeeds; `python -m webapp` (via `run_webapp.sh`)
  serves `/`, `/static/style.css` and `/healthz` with 200.
- Optional browser checks with the preinstalled Playwright/Chromium all exit 0:
  `browser_disclosures.mjs`, `browser_binding_assignments.mjs` (8 PASS),
  `browser_schedule_progress.mjs`, `browser_schedule_visibility.mjs` (8 PASS).
- `test/run_tests.sh`: exit 0, 492 passing checks (488 + 4 new), no failures.
- Deviations from the design are recorded under "Implementation notes" in `design.md`
  (registrar name, `challenges.py`, view builders reading `flask.g`, single split commit).
  The patch-target rule is documented in `AGENTS.md` (Architecture) rather than in each
  module header.
- `openspec validate split-webapp-into-modules --strict`: valid.
