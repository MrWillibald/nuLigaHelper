# Tasks

## 1. Package skeleton

- [ ] 1.1 `git mv webapp.py webapp/__init__.py`; add `webapp/__main__.py` with the former `if __name__ == "__main__"` block and switch `run_webapp.sh` to `python -m webapp`; set explicit `template_folder`/`static_folder` to the repository-root `templates/` and `static/`; verify `test/run_tests.sh`, `gunicorn --check-config webapp:app` and `./run_webapp.sh` start unchanged.
- [ ] 1.2 Move request-scoped helpers (`get_session`, viewer, tier, MV-team lookup, `api_error`) into module functions backed by `flask.g`/`app.extensions`; add a test asserting the `before_request` hook order.

## 2. Topic modules (one commit each, full suite green after each)

- [ ] 2.1 `formatting.py` (`MONATE`, `SLOT_LABELS`, `COUNTRY_CODES`, `ak_color`, `parse_date`, `display_time`) and `security.py` (production boundary, host/body/header/CSRF hooks, authentication guard).
- [ ] 2.2 `auth.py` (login, challenges, registration, verification, approval).
- [ ] 2.3 `views.py` with pure `build_schedule`, `block_view`, `progress_data`, statistics aggregation and `_ordered_person_options`.
- [ ] 2.4 `schedule.py` (schedule page, visibility setting, candidate APIs, due-lock preflight).
- [ ] 2.5 `assignments.py` (game/block claim, release, replacement, takeover, `/dienst`, cake settings, responsible team).
- [ ] 2.6 `persons.py` and `statistics.py` (persons pages and MV API; statistics and audit).

## 3. Finish

- [ ] 3.1 Reduce `webapp/__init__.py` to `create_app()`, hook registration order, module registration and the temporary re-exports; verify no module exceeds ~600 lines and no function closes over `app`.
- [ ] 3.2 Update module descriptions in `AGENTS.md` and `CLAUDE.md`.
- [ ] 3.3 Run `test/run_tests.sh`, the optional browser checks in `test/README.md` if available, and `openspec validate split-webapp-into-modules`; diff rendered HTML of `/`, `/statistik`, `/personen` for a synthetic database before/after (expect identical output apart from CSRF tokens).
