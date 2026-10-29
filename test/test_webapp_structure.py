# ---------------------------------------------------------------
#                          nuLigaHelper – tests
# ---------------------------------------------------------------
# Structure of the webapp package: stable endpoint names and URLs,
# request-hook order and request-independent view builders.
# ---------------------------------------------------------------

import logging
from datetime import date

import helpers as h
from nuligahelper import db
from nuligahelper import web as webapp
from nuligahelper.web.statistics import build_statistics

EXPECTED_RULES = {
    ("/", "schedule"), ("/api/assignment/cancel-replacement", "api_assignment_cancel_replacement"),
    ("/api/assignment/claim", "api_assignment_claim"),
    ("/api/assignment/release", "api_assignment_release"),
    ("/api/assignment/request-replacement", "api_assignment_request_replacement"),
    ("/api/assignment/takeover", "api_assignment_takeover"),
    ("/api/block-assignment/cancel-replacement", "api_block_cancel_replacement"),
    ("/api/block-assignment/claim", "api_block_assignment_claim"),
    ("/api/block-assignment/release", "api_block_assignment_release"),
    ("/api/block-assignment/request-replacement", "api_block_request_replacement"),
    ("/api/block-assignment/takeover", "api_block_takeover"),
    ("/api/blocks/<int:block_id>/cake-settings", "api_cake_settings"),
    ("/api/blocks/<int:block_id>/candidates", "api_block_candidates"),
    ("/api/games/<int:game_id>/candidates", "api_game_candidates"),
    ("/api/games/<int:game_id>/team", "api_game_team"),
    ("/api/teams/<int:team_id>/mv", "api_team_mv"),
    ("/audit", "audit"), ("/datenschutz", "datenschutz"), ("/dienst", "assignment_form"),
    ("/healthz", "health"), ("/impressum", "impressum"), ("/login", "login"),
    ("/login/code", "login_code"), ("/login/token/<token>", "login_token"), ("/logout", "logout"),
    ("/personen", "persons"), ("/personen/<int:person_id>/deactivate", "deactivate_person"),
    ("/personen/<int:person_id>/delete", "delete_person"),
    ("/personen/<int:person_id>/edit", "edit_person"),
    ("/personen/<int:person_id>/reactivate", "reactivate_person"),
    ("/personen/<int:person_id>/teams", "edit_person_teams"),
    ("/personen/<int:person_id>/teams/<int:team_id>/<action>", "change_roster_membership"),
    ("/personen/add", "add_person"), ("/registrieren", "register"),
    ("/registrieren/code", "registration_code"),
    ("/registrieren/verifizieren/<token>", "verify_registration"),
    ("/registrierung/status", "registration_status"),
    ("/registrierungen/<int:person_id>/<decision>", "decide_registration"),
    ("/spielplan/sichtbarkeit", "schedule_visibility"), ("/static/<path:filename>", "static"),
    ("/statistik", "statistics"),
}


def test_endpoint_names_and_urls_are_unchanged():
    rules = {(rule.rule, rule.endpoint) for rule in webapp.app.url_map.iter_rules()}
    assert rules == EXPECTED_RULES, (
        "templates and static/app.js rely on these URLs and url_for() names: "
        f"missing {EXPECTED_RULES - rules}, unexpected {rules - EXPECTED_RULES}"
    )


def test_request_hooks_run_in_the_reviewed_order():
    names = [hook.__name__ for hook in webapp.app.before_request_funcs[None]]
    assert names == [
        "public_readiness", "enforce_trusted_host", "enforce_content_length", "load_viewer",
        "require_authentication", "csrf_protect", "persist_scheduled_duty_locks",
    ], "authentication and CSRF must run before the due-lock preflight writes anything"
    assert [hook.__name__ for hook in webapp.app.after_request_funcs[None]] == [
        "production_response_headers"]
    assert [hook.__name__ for hook in webapp.app.teardown_appcontext_funcs] == ["close_session"]


def test_each_app_keeps_its_own_state():
    first, second = webapp.app, webapp.create_app()
    assert first.extensions["nuligahelper"] is not second.extensions["nuligahelper"], (
        "per-app resources must not leak between create_app() calls"
    )
    assert second.extensions["nuligahelper"].abuse_service is second.extensions[
        "nuligahelper_auth_abuse"]


def test_statistics_builder_runs_without_a_request():
    engine = h.make_engine()
    with h.Session(engine) as session:
        db.sync_games(session, h.sample_games(), h.SEASON)
        session.commit()
        data = build_statistics(session, h.SEASON, date(2026, 9, 1))
        assert data["total_games"] == len(session.query(db.Game).all())
        assert {"team_stats", "person_stats", "gaps"} <= data.keys()
        assert data["gaps"], "unstaffed future games must appear as outstanding gaps"



def test_production_journal_classifies_flask_logs_as_web():
    from nuligahelper.ops.journal_logging import JournalFormatter
    record = logging.LogRecord(webapp.app.logger.name, logging.INFO, "", 0,
                               "web request method=GET path=/ status=200", (), None)
    assert "operation=web " in JournalFormatter("%(message)s").format(record), (
        "web request logs must keep their journal component after the package rename"
    )


if __name__ == "__main__":
    h.run_all(dict(globals()))
