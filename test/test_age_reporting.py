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


def test_browser_rejected_takeover_preserves_saved_responsibility():
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
    assert release.status_code == 400, "ordinary withdrawal must not vacate the sole adult's duty"
    request = admin.post("/api/assignment/request-replacement", json={
        "game_id": ids["game"], "role": db.ROLE_SALE, "slot": 1,
        "expected_person_id": ids["adult"],
    }, headers=h.csrf_headers())
    assert request.get_json()["ok"], "the sole adult must be able to seek replacement"
    assert request.get_json()["assignment"]["person_id"] == ids["adult"]
    assert not any(d["code"] == "missing_adult_seller" for d in request.get_json()["staffing"]["deficiencies"])
    takeover_candidates = _candidates(admin, ids)["slots"]["Verkauf:1"]["candidate_ids"]
    assert ids["unknown"] not in takeover_candidates and ids["young"] not in takeover_candidates
    assert ids["mv"] in takeover_candidates, "eligible adult takeover disappeared"
    refused = admin.post("/api/assignment/takeover", json={
        "game_id": ids["game"], "role": db.ROLE_SALE, "slot": 1,
        "expected_person_id": ids["adult"], "person_id": ids["unknown"],
        "expected_request_token": request.get_json()["assignment"]["replacement_token"],
    }, headers=h.csrf_headers())
    assert refused.status_code == 400, "takeover counted the outgoing adult toward final coverage"
    assert _candidates(admin, ids)["slots"]["Verkauf:1"]["occupant_id"] == ids["adult"]
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
    assert "7 von 7 Pflichtdiensten besetzt" in guest_text, "eligibility fabricated a vacancy"
    assert tree.xpath('//div[@data-eligibility-status and not(@hidden)]'), "full invalid staffing looked complete"
    assert tree.xpath('//details[@class="game-card"]//*[@data-progress-count]/text()') == ["7 von 7 Pflichtdiensten besetzt"]
    assert tree.xpath('//details[@class="game-card"]//*[@data-progress-percent]/text()') == ["100 %"]
    assert tree.xpath('//details[@class="game-card"]//*[@class="coverage-fill"]/@style') == ["width:100%"]
    assert not tree.xpath('//details[@class="game-card"]//*[@data-vacancies]/span'), "age deficiency invented a vacancy badge"
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
    request = admin.post("/api/assignment/request-replacement", json={
        "game_id": ids["game"], "role": db.ROLE_TIMEKEEPER, "slot": 0,
        "expected_person_id": ids["adult"],
    }, headers=h.csrf_headers())
    assert request.get_json()["ok"], "unknown-date current occupant could not seek replacement"
    assert request.get_json()["assignment"]["person_id"] == ids["adult"]
    token = request.get_json()["assignment"]["replacement_token"]
    with h.Session(engine) as session:
        assert session.get(db.Assignment, assignment_id).person_id == ids["adult"]
        assert any(d["code"] == "unknown_birth_date" for d in db.staffing_status(session.get(db.Game, ids["game"]))["deficiencies"])
    takeover = admin.post("/api/assignment/takeover", json={
        "game_id": ids["game"], "role": db.ROLE_TIMEKEEPER, "slot": 0,
        "expected_person_id": ids["adult"], "expected_request_token": token,
        "person_id": ids["mv"],
    }, headers=h.csrf_headers())
    assert takeover.get_json()["ok"], "an eligible incoming helper could not resolve the deficient duty"
    assert not any(d["role"] == db.ROLE_TIMEKEEPER for d in takeover.get_json()["staffing"]["deficiencies"])
    with h.Session(engine) as session:
        assert session.get(db.Assignment, assignment_id).person_id == ids["mv"], "takeover lost assignment identity"


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_security_candidates_claims_and_takeovers_enforce_age_for_every_web_tier(_today):
    app, engine, ids = _site()
    admin = _client(app, ids["admin"])
    candidates = _candidates(admin, ids)["slots"]["Ordnungsdienst:0"]["candidate_ids"]
    assert ids["adult"] in candidates
    assert ids["young"] not in candidates and ids["unknown"] not in candidates
    for actor in ("young", "mv", "admin"):
        response = _claim(_client(app, ids[actor]), ids, db.ROLE_SECURITY, "young")
        assert response.status_code == 400, f"{actor} bypassed Ordnungsdienst age"
        assert "mindestens 18 Jahre" in response.get_json()["error"]
    assert _claim(admin, ids, db.ROLE_SECURITY, "adult").get_json()["ok"]
    target = dict(game_id=ids["game"], role=db.ROLE_SECURITY, slot=0, expected_person_id=ids["adult"])
    request = admin.post("/api/assignment/request-replacement", json=target, headers=h.csrf_headers()).get_json()
    token = request["assignment"]["replacement_token"]
    with h.Session(engine) as session:
        before = session.query(db.AssignmentAudit).count()
    for actor in ("young", "mv", "admin"):
        response = _client(app, ids[actor]).post("/api/assignment/takeover", json={
            **target, "person_id": ids["young"], "expected_request_token": token,
        }, headers=h.csrf_headers())
        assert response.status_code == 400
    with h.Session(engine) as session:
        saved = session.get(db.Game, ids["game"]).assignment_by_role(db.ROLE_SECURITY)
        assert saved.person_id == ids["adult"] and saved.replacement_token == token
        assert session.query(db.AssignmentAudit).count() == before


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_revised_policy_html_confirmation_rechecks_saved_age_and_youth_category(_today):
    for role in (db.ROLE_TIMEKEEPER, db.ROLE_SECURITY, db.ROLE_SALE):
        app, engine, ids = _site()
        admin = _client(app, ids["admin"])
        person = "unknown" if role == db.ROLE_SALE else "adult"
        slot = 1 if role == db.ROLE_SALE else 0
        with h.Session(engine) as session:
            game = session.get(db.Game, ids["game"])
            if role == db.ROLE_SALE:
                game.ak = "SPF Mini"
                db.claim_slot(session, game, db.ROLE_SALE, 0, None, session.get(db.Person, ids["young"]))
            else:
                session.get(db.Person, ids[person]).birth_date = date(2010 if role == db.ROLE_TIMEKEEPER else 2008, 11, 1)
            session.commit()
            before = session.query(db.AssignmentAudit).count()
        target = dict(game_id=ids["game"], role=role, slot=slot, person_id=ids[person],
                      expected_person_id="", action="claim")
        assert ids[person] in _candidates(admin, ids)["slots"][f"{role}:{slot}"]["candidate_ids"]
        review = admin.post("/dienst", data=h.csrf_data({**target, "decision": "review"}))
        assert review.status_code == 200 and "Bestätigen" in review.get_data(as_text=True)
        assert admin.post("/dienst", data=h.csrf_data({**target, "decision": "cancel"})).status_code == 302
        with h.Session(engine) as session:
            assert session.query(db.AssignmentAudit).count() == before
            assert session.get(db.Game, ids["game"]).assignment_by_role(role, slot) is None
        assert admin.post("/dienst", data=h.csrf_data({**target, "decision": "review"})).status_code == 200
        with h.Session(engine) as session:
            if role == db.ROLE_SALE:
                session.get(db.Game, ids["game"]).ak = "BL M"
            else:
                session.get(db.Person, ids[person]).birth_date = date(2012, 11, 1)
            session.commit()
        refused = admin.post("/dienst", data=h.csrf_data({**target, "decision": "confirm"}))
        assert refused.status_code == 400, "HTML confirmation trusted a stale age/category"
        with h.Session(engine) as session:
            assert session.query(db.AssignmentAudit).count() == before
            assert session.get(db.Game, ids["game"]).assignment_by_role(role, slot) is None
        # A new, eligible confirmation must work with the revised policy.
        with h.Session(engine) as session:
            if role == db.ROLE_SALE:
                session.get(db.Game, ids["game"]).ak = "SPF Mini"
            else:
                session.get(db.Person, ids[person]).birth_date = date(2010 if role == db.ROLE_TIMEKEEPER else 2008, 11, 1)
            session.commit()
        saved = admin.post("/dienst", data=h.csrf_data({**target, "decision": "confirm"}))
        assert saved.status_code == 302
        candidates = _candidates(admin, ids)
        assert candidates["slots"][f"{role}:{slot}"]["occupant_id"] == ids[person]
        assert not any(d["role"] == role for d in candidates["staffing"]["deficiencies"])


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_youth_sale_web_takeover_and_refresh_remove_only_adult_without_warning(_today):
    app, engine, ids = _site()
    with h.Session(engine) as session:
        session.get(db.Game, ids["game"]).ak = "BL mD"
        session.commit()
    admin = _client(app, ids["admin"])
    assert _claim(admin, ids, db.ROLE_SALE, "adult").get_json()["ok"]
    assert _claim(admin, ids, db.ROLE_SALE, "young", 1).get_json()["ok"]
    target = dict(game_id=ids["game"], role=db.ROLE_SALE, slot=0, expected_person_id=ids["adult"])
    wanted = admin.post("/api/assignment/request-replacement", json=target, headers=h.csrf_headers()).get_json()
    assert ids["unknown"] in _candidates(admin, ids)["slots"]["Verkauf:0"]["candidate_ids"]
    saved = admin.post("/api/assignment/takeover", json={
        **target, "person_id": ids["unknown"],
        "expected_request_token": wanted["assignment"]["replacement_token"],
    }, headers=h.csrf_headers()).get_json()
    assert saved["ok"] and saved["assignment"]["person_id"] == ids["unknown"]
    assert not any(d["role"] == db.ROLE_SALE for d in saved["staffing"]["deficiencies"])
    refreshed = _candidates(admin, ids)
    assert refreshed["slots"]["Verkauf:0"]["occupant_id"] == ids["unknown"]
    assert refreshed["staffing"] == saved["staffing"]


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_retained_security_deficiencies_and_relaxed_rules_agree_across_reporting(_today):
    from datetime import datetime
    from zoneinfo import ZoneInfo
    for born, code in ((date(2012, 1, 1), "underage"), (None, "unknown_birth_date")):
        app, engine, ids = _site()
        with h.Session(engine) as session:
            game = session.get(db.Game, ids["game"])
            # Synthetic pre-policy roster, including retained locks/requests.
            for index, (role, slot) in enumerate(db.required_positions(game)):
                birthday = (born if role == db.ROLE_SECURITY else date(2010, 11, 1)
                            if role == db.ROLE_TIMEKEEPER else None if role == db.ROLE_SALE
                            else h.ADULT_BIRTH_DATE)
                person = db.Person(name=f"Retained helper {index}", birth_date=birthday)
                session.add(person)
                session.flush()
                session.add(db.Assignment(game=game, person=person, role=role, slot=slot,
                                          locked_at=datetime(2026, 8, 30, 9, tzinfo=ZoneInfo("Europe/Berlin")),
                                          replacement_token=f"retained-token-{index}"))
            game.ak = "SPF Mini"
            session.commit()
            def snapshot(session):
                return [(a.id, a.person_id, a.role, a.slot, a.locked_at, a.replacement_token)
                        for a in session.query(db.Assignment).order_by(db.Assignment.id)]
            before = snapshot(session)
            audits = session.query(db.AssignmentAudit).count()
            status = db.staffing_status(game)
            assert status["required_filled"] == status["required_total"] == 5
            assert [(d["role"], d["code"]) for d in status["deficiencies"]] == [(db.ROLE_SECURITY, code)]
            service = notifier.Notifier(h.load_club_config(), session, h.SEASON)
            with patch.object(service, "_notify_game_helpers", return_value=0), patch.object(service, "_dispatch", return_value=1) as dispatch:
                assert service.notify_game_day(game.date) == 1, "retained security deficiency missed MV follow-up"
                assert "Ordner" in dispatch.call_args.kwargs["mail_body"]
            for category in ("BL M", "GE", "SPF Mini"):
                game.ak = category
                result = db.staffing_status(game)
                assert any(d["role"] == db.ROLE_SALE for d in result["deficiencies"]) == (category != "SPF Mini")
                if category != "GE":
                    assert not any(d["role"] == db.ROLE_TIMEKEEPER for d in result["deficiencies"])
            session.commit()
        guest_text = app.test_client().get("/").get_data(as_text=True)
        tree = html.fromstring(guest_text)
        assert tree.xpath('//details[@class="game-card"]//*[@data-progress-count]/text()') == ["5 von 5 Pflichtdiensten besetzt"]
        assert not tree.xpath('//details[@class="game-card"]//*[@data-vacancies]/span')
        warnings = tree.xpath('//div[@data-eligibility-status]')[0].text_content()
        assert "Ordner" in warnings and "Verkauf" not in warnings
        admin = _client(app, ids["admin"])
        stats = admin.get("/statistik").get_data(as_text=True)
        assert "Ordner" in html.fromstring(stats).text_content() and "SPF Mini" in html.fromstring(stats).text_content()
        candidates_text = admin.get(f'/api/games/{ids["game"]}/candidates').get_data(as_text=True)
        for response in (guest_text, stats, candidates_text):
            assert "2010-11-01" not in response and "2012-01-01" not in response and '"age"' not in response
            assert '"birth_date":' not in response
        with h.Session(engine) as session:
            assert snapshot(session) == before, "reporting changed a binding duty, lock or request"
            assert session.query(db.AssignmentAudit).count() == audits, "policy reevaluation fabricated audits"
            game = session.get(db.Game, ids["game"])
            game.assignment_by_role(db.ROLE_SECURITY).person.birth_date = h.ADULT_BIRTH_DATE
            session.commit()
            assert db.staffing_status(game)["complete"], "youth sale or 16-year-old timing still deficient"
            service = notifier.Notifier(h.load_club_config(), session, h.SEASON)
            with patch.object(service, "_notify_game_helpers", return_value=0), patch.object(service, "_dispatch", return_value=1) as dispatch:
                assert service.notify_game_day(game.date) == 0
                dispatch.assert_not_called()
        warnings = html.fromstring(app.test_client().get("/").get_data(as_text=True)).xpath('//div[@data-eligibility-status]')[0]
        assert warnings.get("hidden") is not None


if __name__ == "__main__":
    h.run_all(dict(globals()))
