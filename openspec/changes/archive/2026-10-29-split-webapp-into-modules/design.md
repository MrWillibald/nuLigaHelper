# Design

## Context

All routes are registered inside `create_app()` (`webapp.py` lines 234–2785). Helpers defined there close over `app` and over each other. Tests create fresh apps via `webapp.create_app()` (26 call sites) and patch module objects reachable as `webapp.common`, `webapp.db` and `webapp.notifier`. Templates call `url_for()` with unprefixed endpoint names (`schedule`, `persons`, `edit_person_teams`, `api_game_candidates`, …; ~40 call sites).

## Goals / Non-Goals

**Goals:** one file per topic, each readable on its own; view-model builders testable without a request; zero behavior change; no change for gunicorn, systemd or tests.

**Non-Goals:** no new framework or extension, no URL/endpoint renaming, no template restructuring, no move into the `nuligahelper/` package (that is `restructure-project-layout`), no business-rule changes (those stay in `db.py`).

## Decisions

### 1. `register(app)` modules instead of Flask Blueprints

Flask Blueprints always prefix endpoint names (`schedule.schedule`), which would require changing every `url_for()` in templates and tests. Each topic module instead exposes `register(app: Flask) -> None` and registers its routes with `@app.route`/`app.add_url_rule` exactly as today. `create_app()` calls `security.register(app)` first (its `before_request` hooks must keep their current order), then the page modules. Converting to Blueprints later remains possible but is not needed for readability.

### 2. Request-scoped helpers via `flask.g`

`get_session()`, the current viewer and its tier move to `webapp/context.py` (or `security.py`) as module functions reading/writing `flask.g` and `current_app.config`. The session factory is stored on `app.extensions["nuligahelper"]` in `create_app()`. Page modules import these helpers; nothing closes over `app`.

### 3. Hook order is part of the contract

Current `before_request` order: public readiness → trusted host → content length → load viewer → require authentication → CSRF → due-lock preflight (registered later inside the schedule section). The refactoring keeps this order explicitly in `create_app()` and adds a test asserting `app.before_request_funcs[None]` names in sequence, so a future reordering is a conscious change.

### 4. Pure view builders

`build_schedule()`, `block_view()`, `progress_data()` and the statistics aggregation become functions of `(session, viewer, today, params)` returning the same dicts the templates consume today. Routes only parse the request, call the builder and render. Existing HTTP tests stay as the safety net; new direct unit tests are optional.

### 5. Compatibility re-exports

`webapp/__init__.py` imports `common`, `db`, `notifier` and re-exports `_ordered_person_options` so existing tests and fixtures keep working. These re-exports are removed in `restructure-project-layout`, where tests are updated anyway.

## Implementation notes (2026-10-29)

- Routes and hooks are declared on a module-level `webapp.context.Registry` and replayed
  by `register_routes(app)`; the function is not called `register` because the
  `/registrieren` view already owns that endpoint name.
- Per-app resources are an immutable `AppState` in `app.extensions["nuligahelper"]`,
  read through `state()`; nothing closes over `app`.
- `auth.py` stayed above the size target, so contact validation, abuse limits, signed
  challenges and account messages moved to `webapp/challenges.py`.
- `build_statistics(session, season_year, today, hide_from)` is request-independent and
  lives in `webapp/statistics.py` next to its only caller. `build_schedule()`,
  `block_view()` and the duty-state helpers stay module functions in `webapp/views.py`
  that read the viewer from `flask.g`; making them fully viewer-explicit would change
  every helper signature and is not needed for readability.
- The topic split landed in one commit after the package skeleton commit: the closures
  shared state, so intermediate states would have needed throw-away glue. The suite,
  pyflakes and a before/after page comparison ran on the result.

## Risks / Trade-offs

- **Hidden closure coupling** (a helper silently using another route's local) → move one topic per commit, run the full suite after each move.
- **Patch targets**: tests that patch `webapp.common.effective_today` patch the shared `common` module object, so they keep working as long as page modules call `common.effective_today()` through the module, never via `from common import effective_today`. This rule is stated in the module headers.
- **Order-dependent webapp tests** (documented in `AGENTS.md`) → no test file is reorganized in this change.
