"""Synthetic web, HTML confirmation and actual reminder-batch commitments."""
import os
from datetime import date, datetime
from zoneinfo import ZoneInfo
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from lxml import html

import helpers as h
import db
import main
import webapp


def fixture():
    directory = TemporaryDirectory(prefix="nuliga-sticky-web-")
    path = str(Path(directory.name) / "synthetic.db")
    engine = db.make_engine(path)
    db.initialize_db(engine)
    with h.Session(engine) as session:
        db.sync_games(session, [{"game_nr": "1", "date": "01.11.2026", "time": "12:00",
            "day": "So", "hall": 1, "ak": "BL M", "home": "TuS", "guest": "Test", "score": ""}], h.SEASON)
        game = session.query(db.Game).one()
        team = session.query(db.Team).filter_by(name="BL M").one()
        people = [db.Person(name=name, birth_date=h.ADULT_BIRTH_DATE, teams=[team],
                            is_admin=name == "Admin") for name in ("Original", "Replacement", "Admin", "MV")]
        session.add_all(people)
        session.flush()
        team.mv_person_id = people[3].id
        blocks = db.get_day_blocks(session, h.SEASON, game.date)
        cake = next(b for b in blocks if b.phase == db.BLOCK_CAKE_DELIVERY)
        db.configure_cake_block(session, cake, "11:00", 2, None, None)
        session.commit()
        ids = {p.name: p.id for p in people}
        ids.update(game=game.id, **{b.phase: b.id for b in blocks})
    with patch.dict(os.environ, {"NULIGAHELPER_DB": path}):
        app = webapp.create_app()
    return directory, engine, app, ids


def client_for(app, ids, name):
    client = app.test_client()
    h.sign_in(client, ids[name])
    return client


def write(client, path, **payload):
    return client.post(path, json=payload, headers=h.csrf_headers())


def audit_count(engine):
    with h.Session(engine) as session:
        return session.query(db.AssignmentAudit).count()


@patch("common.effective_today", return_value=date(2026, 10, 30))
def test_game_request_and_atomic_takeover_keep_responsibility_and_privacy(_today):
    directory, engine, app, ids = fixture()
    try:
        original = client_for(app, ids, "Original")
        replacement = client_for(app, ids, "Replacement")
        target = dict(game_id=ids["game"], role=db.ROLE_SALE, slot=0)
        assert write(original, "/api/assignment/claim", **target, expected_person_id=None).json["ok"]
        requested = write(original, "/api/assignment/request-replacement", **target,
                          expected_person_id=ids["Original"])
        assert requested.json["assignment"]["person_id"] == ids["Original"], "search must never empty a duty"
        token = requested.json["assignment"]["replacement_token"]
        candidates = replacement.get(f'/api/games/{ids["game"]}/candidates').json
        assert ids["Replacement"] in candidates["slots"][f"{db.ROLE_SALE}:0"]["candidate_ids"]
        assert "Ersatz gesucht" in app.test_client().get("/").get_data(as_text=True)
        public = app.test_client().get("/").get_data(as_text=True)
        assert "replacement_token" not in public and "data-occupant-id" not in public
        response = write(replacement, "/api/assignment/takeover", **target,
                         expected_person_id=ids["Original"], expected_request_token=token)
        assert response.status_code == 200 and response.json["assignment"]["person_id"] == ids["Replacement"]
        assert response.json["staffing"]["required_filled"] == 1
        stale = write(original, "/api/assignment/takeover", **target,
                      expected_person_id=ids["Original"], expected_request_token=token)
        assert stale.status_code == 409, "stale outgoing ownership must not overwrite the takeover winner"
    finally:
        engine.dispose()
        directory.cleanup()


@patch("common.effective_today", return_value=date(2026, 10, 30))
def test_every_block_uses_request_takeover_and_no_ordinary_release(_today):
    directory, engine, app, ids = fixture()
    try:
        original = client_for(app, ids, "Original")
        replacement = client_for(app, ids, "Replacement")
        for phase in db.BLOCK_PHASES:
            target = dict(block_id=ids[phase], slot=0)
            assert write(original, "/api/block-assignment/claim", **target, expected_person_id=None).json["ok"]
            refused = write(original, "/api/block-assignment/release", **target, expected_person_id=ids["Original"])
            assert refused.status_code == 400
            wanted = write(original, "/api/block-assignment/request-replacement", **target,
                           expected_person_id=ids["Original"])
            assert wanted.json["assignment"]["person_id"] == ids["Original"]
            token = wanted.json["assignment"]["replacement_token"]
            done = write(replacement, "/api/block-assignment/takeover", **target,
                         expected_person_id=ids["Original"], expected_request_token=token)
            assert done.json["assignment"]["person_id"] == ids["Replacement"], phase
    finally:
        engine.dispose()
        directory.cleanup()


@patch("common.effective_today", return_value=date(2026, 10, 30))
def test_html_confirmation_cancel_and_concurrent_winner_never_need_compensating_release(_today):
    directory, engine, app, ids = fixture()
    try:
        first = client_for(app, ids, "Original")
        second = client_for(app, ids, "Replacement")
        targets = [dict(game_id=ids["game"], role=db.ROLE_SALE, slot=0)] + [
            dict(block_id=ids[phase], slot=0) for phase in db.BLOCK_PHASES]
        for target in targets:
            before = audit_count(engine)
            review = first.post("/dienst", data=h.csrf_data({**target, "person_id": ids["Original"],
                                                           "expected_person_id": "", "decision": "review"}))
            assert review.status_code == 200
            assert "Bestätigen" in review.get_data(as_text=True) and "Abbrechen" in review.get_data(as_text=True)
            assert audit_count(engine) == before, "selection and review are read-only"
            cancel = first.post("/dienst", data=h.csrf_data({**target, "decision": "cancel"}))
            assert cancel.status_code == 302 and audit_count(engine) == before
            api = "/api/block-assignment/claim" if "block_id" in target else "/api/assignment/claim"
            assert write(second, api, **target, expected_person_id=None).json["ok"]
            winner_audits = audit_count(engine)
            assert first.post("/dienst", data=h.csrf_data({**target, "decision": "cancel"})).status_code == 302
            assert audit_count(engine) == winner_audits
            conflict = first.post("/dienst", data=h.csrf_data({**target, "person_id": ids["Original"],
                 "expected_person_id": "", "action": "claim", "decision": "confirm"}))
            assert conflict.status_code == 409 and audit_count(engine) == winner_audits
        guest = app.test_client().get("/dienst?game_id=1&role=Verkauf&slot=0")
        assert guest.status_code == 302, "HTML confirmation is never public"
    finally:
        engine.dispose()
        directory.cleanup()


@patch("common.effective_today", return_value=date(2026, 10, 30))
def test_actual_batch_locks_contactless_helpers_and_late_claims_even_when_delivery_fails(_today):
    directory, engine, app, ids = fixture()
    try:
        first = client_for(app, ids, "Original")
        admin = client_for(app, ids, "Admin")
        second = client_for(app, ids, "Replacement")
        game_target = dict(game_id=ids["game"], role=db.ROLE_SALE, slot=0)
        assert write(first, "/api/assignment/claim", **game_target, expected_person_id=None).json["ok"]
        wanted = write(first, "/api/assignment/request-replacement", **game_target, expected_person_id=ids["Original"])
        token = wanted.json["assignment"]["replacement_token"]
        class FakeNotifier:
            def __init__(self, *args): pass
            def notify_game_day(self, *_): raise RuntimeError("synthetic transport failure")
            def __getattr__(self, _): return lambda *args: 0
        with patch("db.duty_now", return_value=datetime(2026, 10, 31, 9, tzinfo=ZoneInfo("Europe/Berlin"))):
            with patch.object(main, "Notifier", FakeNotifier):
                try:
                    main._send_notifications({}, engine, h.SEASON, date(2026, 10, 31), db.SyncEvents())
                except RuntimeError as exc:
                    assert "synthetic transport" in str(exc)
                else:
                    raise AssertionError("synthetic transport failure was not exercised")
            with h.Session(engine) as session:
                game = session.get(db.Game, ids["game"])
                assignment = next(a for a in game.assignments if a.slot == 0)
                assert db.duty_is_locked(session, game, assignment) and not assignment.replacement_token
            refused = write(second, "/api/assignment/takeover", **game_target,
                            expected_person_id=ids["Original"], expected_request_token=token)
            assert refused.status_code in (400, 409)
            for person_client in (first, admin):
                assert write(person_client, "/api/assignment/request-replacement", **game_target,
                             expected_person_id=ids["Original"]).status_code == 400
            for target, api in [(dict(game_id=ids["game"], role=db.ROLE_SALE, slot=1), "/api/assignment/claim")] + [
                (dict(block_id=ids[phase], slot=1), "/api/block-assignment/claim") for phase in db.BLOCK_PHASES]:
                late = write(second, api, **target, expected_person_id=None)
                assert late.json["ok"] and late.json["assignment"]["locked"], "late empty claim must lock in the claim transaction"
    finally:
        engine.dispose()
        directory.cleanup()


@patch("common.effective_today", return_value=date(2026, 10, 31))
def test_fixed_nine_oclock_cutoff_needs_no_job_and_cancel_never_materializes_mutations(_today):
    directory, engine, app, ids = fixture()
    berlin = ZoneInfo("Europe/Berlin")
    try:
        original = client_for(app, ids, "Original")
        replacement = client_for(app, ids, "Replacement")
        target = dict(game_id=ids["game"], role=db.ROLE_SALE, slot=0)
        with patch("db.duty_now", return_value=datetime(2026, 10, 31, 8, 59, tzinfo=berlin)):
            assert write(original, "/api/assignment/claim", **target, expected_person_id=None).json["ok"]
            wanted = write(original, "/api/assignment/request-replacement", **target, expected_person_id=ids["Original"])
            token = wanted.json["assignment"]["replacement_token"]
            before = audit_count(engine)
        with patch("db.duty_now", return_value=datetime(2026, 10, 31, 9, 0, tzinfo=berlin)):
            cancelled = original.post("/dienst", data=h.csrf_data({**target, "decision": "cancel"}))
            assert cancelled.status_code == 302 and audit_count(engine) == before, "Cancel must not even materialize due locks"
            page = html.fromstring(original.get("/").get_data(as_text=True))
            field = page.xpath(f'//*[@id="game-{ids["game"]}"]//*[@data-occupant-id="{ids["Original"]}"]')[0]
            assert not field.xpath('.//select | .//button[@data-request-replacement]'), "scheduled lock applies without cron"
            assert audit_count(engine) == before, "schedule reads derive the clock without assignment writes"
            candidates = replacement.get(f'/api/games/{ids["game"]}/candidates').json
            assert candidates["slots"][f"{db.ROLE_SALE}:0"]["locked"]
            assert not candidates["slots"][f"{db.ROLE_SALE}:0"]["candidate_ids"]
            refused = write(replacement, "/api/assignment/takeover", **target,
                            expected_person_id=ids["Original"], expected_request_token=token)
            assert refused.status_code in (400, 409)
            with h.Session(engine) as session:
                assignment = session.query(db.Assignment).filter_by(game_id=ids["game"]).one()
                assert assignment.locked_at is not None and assignment.replacement_token is None
                assert assignment.person_id == ids["Original"]
            late_page = html.fromstring(replacement.get("/").get_data(as_text=True))
            assert late_page.xpath(f'//*[@id="game-{ids["game"]}"]//select[@data-role="Verkauf" and @data-slot="1"]'), "genuinely empty late duties must stay claimable"
    finally:
        engine.dispose()
        directory.cleanup()


@patch("common.effective_today", return_value=date(2026, 10, 30))
def test_admin_cannot_cancel_current_commitments_by_deactivation_or_deletion(_today):
    directory, engine, app, ids = fixture()
    try:
        original = client_for(app, ids, "Original")
        admin = client_for(app, ids, "Admin")
        assert write(original, "/api/assignment/claim", game_id=ids["game"], role=db.ROLE_SALE,
                     slot=0, expected_person_id=None).json["ok"]
        before = audit_count(engine)
        for action in ("deactivate", "delete"):
            response = admin.post(f'/personen/{ids["Original"]}/{action}', data=h.csrf_data())
            assert response.status_code == 409 and not response.is_json
            assert "verbindliche" in response.get_data(as_text=True)
            with h.Session(engine) as session:
                person = session.get(db.Person, ids["Original"])
                assert person is not None and person.account_status == db.ACCOUNT_ACTIVE
                assert len(person.assignments) == 1
            assert audit_count(engine) == before, "refused maintenance must not claim a cancellation"
    finally:
        engine.dispose()
        directory.cleanup()


@patch("common.effective_today", return_value=date(2026, 10, 25))
def test_weekly_reminder_keeps_search_open_and_preflight_contention_is_reported_without_false_success(_today):
    directory, engine, app, ids = fixture()
    try:
        original = client_for(app, ids, "Original")
        target = dict(game_id=ids["game"], role=db.ROLE_SALE, slot=0)
        assert write(original, "/api/assignment/claim", **target, expected_person_id=None).json["ok"]
        wanted = write(original, "/api/assignment/request-replacement", **target, expected_person_id=ids["Original"])
        token = wanted.json["assignment"]["replacement_token"]
        class FakeNotifier:
            def __init__(self, *args): pass
            def __getattr__(self, _): return lambda *args: 0
        with patch.object(main, "Notifier", FakeNotifier), patch("db.duty_now", return_value=datetime(2026,10,25,9,tzinfo=ZoneInfo("Europe/Berlin"))):
            main._send_notifications({}, engine, h.SEASON, date(2026,10,25), db.SyncEvents())
        with h.Session(engine) as session:
            assignment = session.query(db.Assignment).filter_by(game_id=ids["game"]).one()
            assert assignment.replacement_token == token and assignment.locked_at is None
            assert session.query(db.DutyCutoff).count() == 0
        before = audit_count(engine)
        with patch("db.lock_due_duties", side_effect=db.AssignmentTemporarilyUnavailableError()):
            response = write(original, "/api/assignment/request-replacement", **target, expected_person_id=ids["Original"])
            assert response.status_code == 503 and response.json["code"] == "temporarily_unavailable"
            html_response = original.post("/dienst", data=h.csrf_data({**target,"action":"request-replacement", "expected_person_id":ids["Original"]}))
            assert html_response.status_code == 503 and not html_response.is_json
        assert audit_count(engine) == before
    finally:
        engine.dispose()
        directory.cleanup()


if __name__ == "__main__":
    h.run_all(dict(globals()))
