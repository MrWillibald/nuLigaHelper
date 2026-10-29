# Proposal

## Why

`webapp.py` has 2,789 lines; about 2,550 of them live inside `create_app()`, where every route and helper is a nested closure sharing `get_session`, `tier_for`, `current_mv_team_ids` and similar functions. A reader looking for the statistics page has to scroll through authentication, CSRF and assignment code, and logic such as `build_schedule()` (~190 lines) or the statistics calculation (~125 lines) can only be exercised through HTTP requests. This is the largest readability problem in the repository (phase 1 of `docs/REFACTORING-VORSCHLAG.md`).

## What Changes

- Turn `webapp.py` into the package `webapp/` with the same import name. `webapp/__init__.py` keeps `create_app()` and the module-level `app`, so `gunicorn webapp:app`, all systemd units and the tests' `import webapp` keep working. The local development server moves from `python webapp.py` to `python -m webapp` (`webapp/__main__.py`); `run_webapp.sh` is updated accordingly.
- Split by page/topic into modules that each expose `register(app)`:
  - `security.py` – production proxy boundary, trusted hosts, body limit, response headers, CSRF, viewer loading and tier derivation, default-deny route guard.
  - `auth.py` – login, codes/challenges, registration, verification, approval decisions.
  - `schedule.py` – `/`, visibility setting, candidate APIs, due-lock preflight.
  - `assignments.py` – `/api/assignment/*`, `/api/block-assignment/*`, `/dienst`, cake settings, responsible-team API.
  - `persons.py` – `/personen*`, `/api/teams/<id>/mv`.
  - `statistics.py` – `/statistik`, `/audit`.
  - `views.py` – pure view-model builders (`build_schedule`, `block_view`, statistics aggregation, `_ordered_person_options`).
  - `formatting.py` – `MONATE`, `SLOT_LABELS`, `COUNTRY_CODES`, `ak_color`, `parse_date`, `display_time`.
- Shared per-request state (`get_session`, viewer, tier) becomes plain module functions using `flask.g`; no closures across modules.
- Endpoint names, URLs, templates, JSON payloads, German texts and HTTP behavior stay identical.

## Capabilities

### New Capabilities

None.

### Modified Capabilities

None. Pure refactoring; observable behavior is unchanged.

## Impact

- `webapp.py` → `webapp/` package (≈9 files). `templates/` and `static/` stay at the repository root for now and are found via an explicit `template_folder`/`static_folder`.
- Tests: no change expected except where they reach into internals (`webapp._ordered_person_options`, `webapp.common`, `webapp.db`, `webapp.notifier`), which remain re-exported from `webapp/__init__.py` during this change.
- No schema, deployment-asset or documentation behavior change; `AGENTS.md`/`CLAUDE.md` module descriptions are updated.
