"""Saved age eligibility agrees across candidates, cards, statistics and MV follow-up."""

import os
import tempfile
import shutil
import subprocess
from datetime import date
from unittest.mock import patch

from lxml import html
from sqlalchemy import select

import helpers as h
import db
import notifier
import webapp


def _site():
    path = os.path.join(h._TEST_DIR, f"age-reporting-{next(tempfile._get_candidate_names())}.db")
    engine = db.make_engine(path)
    db.initialize_db(engine)
    previous = os.environ["NULIGAHELPER_DB"]
    os.environ["NULIGAHELPER_DB"] = path
    try:
        app = webapp.create_app()
    finally:
        os.environ["NULIGAHELPER_DB"] = previous
    with h.Session(engine) as session:
        db.sync_games(session, [{
            "day": "So", "date": "01.11.2026", "time": "14:00", "hall": 1,
            "game_nr": "7001", "ak": "BL M", "home": "TuS", "guest": "Test",
            "score": "",
        }], h.SEASON)
        team = session.scalar(select(db.Team).where(db.Team.name == "BL M"))
        persons = {
            "admin": db.Person(name="Administrator", birth_date=date(1990, 1, 1), is_admin=True),
            "mv": db.Person(name="Verantwortlicher", birth_date=date(1991, 1, 1),
                            email="mv@fixture.invalid", teams=[team]),
            "adult": db.Person(name="Erwachsene", birth_date=date(1992, 1, 1), teams=[team]),
            "young": db.Person(name="Junge Person", birth_date=date(2012, 1, 1), teams=[team]),
            "unknown": db.Person(name="Altdaten", birth_date=None, teams=[team]),
        }
        session.add_all(persons.values())
        session.flush()
        team.mv_person_id = persons["mv"].id
        game = session.scalar(select(db.Game).where(db.Game.game_nr == "7001"))
        game.team = team
        session.commit()
        ids = {key: person.id for key, person in persons.items()}
        ids["game"] = game.id
    return app, engine, ids


def _client(app, person_id):
    client = app.test_client()
    h.sign_in(client, person_id)
    return client


def _claim(client, ids, role, person, slot=0):
    return client.post("/api/assignment/claim", json={
        "game_id": ids["game"], "role": role, "slot": slot,
        "expected_person_id": None, "person_id": ids[person],
    }, headers=h.csrf_headers())


def _candidates(client, ids):
    return client.get(f'/api/games/{ids["game"]}/candidates').get_json()


def test_browser_rejected_replacement_preserves_the_saved_release():
    node = shutil.which("node")
    if node is None:
        return
    result = subprocess.run([node, "test/js_age_replacement.mjs"], cwd=h.PROJECT_DIR,
                            capture_output=True, text=True)
    assert result.returncode == 0, result.stderr


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_sale_candidates_follow_saved_collective_state_and_keep_private_dates(_today):
    app, engine, ids = _site()
    admin = _client(app, ids["admin"])
    initial = _candidates(admin, ids)
    assert ids["unknown"] not in initial["slots"]["Zeitnehmer:0"]["candidate_ids"]
    assert ids["young"] not in initial["slots"]["Sekretär:0"]["candidate_ids"]
    assert ids["young"] in initial["slots"]["Verkauf:0"]["candidate_ids"]
    assert ids["unknown"] in initial["slots"]["Verkauf:0"]["candidate_ids"]
    assert _claim(admin, ids, db.ROLE_SALE, "young").get_json()["ok"]
    after_first = _candidates(admin, ids)
    assert ids["unknown"] not in after_first["slots"]["Verkauf:1"]["candidate_ids"]
    assert ids["adult"] in after_first["slots"]["Verkauf:1"]["candidate_ids"]
    refusal = _claim(admin, ids, db.ROLE_SALE, "unknown", 1)
    assert refusal.status_code == 400, "a stale final-slot candidate bypassed adult coverage"
    assert _candidates(admin, ids)["slots"]["Verkauf:1"]["occupant_id"] is None
    assert _claim(admin, ids, db.ROLE_SALE, "adult", 1).get_json()["ok"]
    release = admin.post("/api/assignment/release", json={
        "game_id": ids["game"], "role": db.ROLE_SALE, "slot": 1,
        "expected_person_id": ids["adult"],
    }, headers=h.csrf_headers())
    assert release.get_json()["ok"], "the sole adult's authorized withdrawal was refused"
    assert any(d["code"] == "missing_adult_seller" for d in release.get_json()["staffing"]["deficiencies"])
    assert ids["unknown"] not in _candidates(admin, ids)["slots"]["Verkauf:1"]["candidate_ids"]
    text = admin.get(f'/api/games/{ids["game"]}/candidates').get_data(as_text=True)
    assert "birth_date" not in text and "1992-01-01" not in text and "2012-01-01" not in text
    with h.Session(engine) as session:
        assert len(session.scalars(select(db.AssignmentAudit)).all()) == 3, "refusal wrote a successful audit"


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_all_web_tiers_and_past_admin_corrections_enforce_timing_age(_today):
    app, engine, ids = _site()
    for actor in ("young", "mv", "admin"):
        client = _client(app, ids[actor])
        response = _claim(client, ids, db.ROLE_TIMEKEEPER, "young")
        assert response.status_code == 400, f"{actor} bypassed stored eligibility"
    with h.Session(engine) as session:
        session.get(db.Game, ids["game"]).date = "01.08.2026"
        session.commit()
    response = _claim(_client(app, ids["admin"]), ids, db.ROLE_SECRETARY, "young")
    assert response.status_code == 400, "past-game admin rights became an age override"


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_full_invalid_roster_reports_truthful_occupancy_without_private_data(_today):
    app, engine, ids = _site()
    with h.Session(engine) as session:
        game = session.get(db.Game, ids["game"])
        # Legacy occupied duties must survive reevaluation even when no claim
        # would be accepted for the same roster now.
        for index, (role, slot) in enumerate(db.required_positions(game)):
            person = db.Person(name=f"Legacy helper {index}", birth_date=None)
            session.add(person)
            session.flush()
            session.add(db.Assignment(game=game, person=person, role=role, slot=slot))
        session.commit()
    guest_text = app.test_client().get("/").get_data(as_text=True)
    tree = html.fromstring(guest_text)
    assert "8 von 8 Pflichtdiensten besetzt" in guest_text, "eligibility fabricated a vacancy"
    assert tree.xpath('//div[@data-eligibility-status and not(@hidden)]'), "full invalid staffing looked complete"
    assert "Legacy helper 0" in guest_text
    assert "birth_date" not in guest_text and "1992-01-01" not in guest_text
    assert not tree.xpath('//select[@data-role]') and "data-occupant-id" not in guest_text
    stats = _client(app, ids["admin"]).get("/statistik").get_data(as_text=True)
    assert "TuS – Test" in html.fromstring(stats).text_content()
    assert "Geburtsdatum" in stats and "1992-01-01" not in stats
    with h.Session(engine) as session:
        game = session.get(db.Game, ids["game"])
        assert db.missing_slots(game) == {}, "physical vacancy semantics changed"
        assert not db.staffing_status(game)["complete"]
        service = notifier.Notifier(h.load_club_config(), session, h.SEASON)
        with patch.object(service, "_notify_game_helpers", return_value=0), patch.object(service, "_dispatch", return_value=1) as dispatch:
            assert service.notify_game_day(game.date) == 1, "full invalid roster missed responsible MV follow-up"
            assert dispatch.call_args.args[0]["name"] == "Verantwortlicher"
            body = dispatch.call_args.kwargs["mail_body"]
            assert "Altersanforderungen" in body and "1991-01-01" not in body
            game.team.mv_person.account_status = db.ACCOUNT_INACTIVE
            assert service.notify_game_day(game.date) == 0, "inactive MV received follow-up"
            game.team.mv_person_id = None
            assert service.notify_game_day(game.date) == 0, "missing MV caused fallback routing"


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_game_and_birth_corrections_reevaluate_without_rewriting_assignments(_today):
    app, engine, ids = _site()
    with h.Session(engine) as session:
        game = session.get(db.Game, ids["game"])
        person = session.get(db.Person, ids["adult"])
        person.birth_date = date(2012, 11, 1)
        game.ak = "BL mD"
        assignment = db.claim_slot(session, game, db.ROLE_TIMEKEEPER, 0, None, person)
        session.commit()
        assignment_id = assignment.id
        audits = [(a.id, a.game_snapshot, a.affected_person_name) for a in session.scalars(select(db.AssignmentAudit))]
        game.date = "31.10.2026"
        assert any(d["code"] == "underage" for d in db.staffing_status(game)["deficiencies"])
        game.date = "01.11.2026"
        game.ak = "BL M"
        assert any(d["code"] == "underage" for d in db.staffing_status(game)["deficiencies"])
        person.birth_date = date(1992, 1, 1)
        assert not any(d["role"] == db.ROLE_TIMEKEEPER for d in db.staffing_status(game)["deficiencies"])
        person.birth_date = None
        session.commit()
        assert session.get(db.Assignment, assignment_id).person_id == person.id
        assert audits == [(a.id, a.game_snapshot, a.affected_person_name) for a in session.scalars(select(db.AssignmentAudit))]
    admin = _client(app, ids["admin"])
    candidates = _candidates(admin, ids)
    assert ids["adult"] in candidates["slots"]["Zeitnehmer:0"]["candidate_ids"], "invalid current occupant disappeared"
    assert ids["adult"] not in candidates["slots"]["Sekretär:0"]["candidate_ids"]
    release = admin.post("/api/assignment/release", json={
        "game_id": ids["game"], "role": db.ROLE_TIMEKEEPER, "slot": 0,
        "expected_person_id": ids["adult"],
    }, headers=h.csrf_headers())
    assert release.get_json()["ok"], "unknown-date current occupant could not be released"


if __name__ == "__main__":
    h.run_all(dict(globals()))
