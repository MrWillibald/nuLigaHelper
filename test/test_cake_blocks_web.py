"""Cake settings, saved schedule state, scoped candidates and reporting."""

import os
import shutil
import subprocess
import tempfile
from datetime import date
from unittest.mock import patch

from lxml import html

import helpers as h
import db
import webapp


def _site():
    path = os.path.join(h._TEST_DIR, f"cake-web-{next(tempfile._get_candidate_names())}.db")
    engine = db.make_engine(path)
    db.initialize_db(engine)
    previous = os.environ["NULIGAHELPER_DB"]
    os.environ["NULIGAHELPER_DB"] = path
    try:
        app = webapp.create_app()
    finally:
        os.environ["NULIGAHELPER_DB"] = previous
    with h.Session(engine) as session:
        db.sync_games(session, [
            {"day": "Sa", "date": "05.09.2026", "time": "12:00", "hall": 1,
             "game_nr": "1", "ak": "BL mD", "home": "TuS", "guest": "A", "score": ""},
            {"day": "Sa", "date": "01.01.2020", "time": "12:00", "hall": 1,
             "game_nr": "2", "ak": "BL F", "home": "TuS", "guest": "B", "score": ""},
        ], h.SEASON)
        team = session.query(db.Team).filter_by(name="BL mD").one()
        admin = db.Person(name="Admin", birth_date=h.ADULT_BIRTH_DATE, is_admin=True)
        mv = db.Person(name="MV", birth_date=h.ADULT_BIRTH_DATE, teams=[team])
        volunteer = db.Person(name="Cake Only", birth_date=h.ADULT_BIRTH_DATE,
                              email="private-cake@fixture.test", teams=[team])
        other = db.Person(name="Other", birth_date=h.ADULT_BIRTH_DATE)
        inactive = db.Person(name="Inactive", birth_date=h.ADULT_BIRTH_DATE,
                             account_status=db.ACCOUNT_INACTIVE)
        pending = db.Person(name="Pending", birth_date=h.ADULT_BIRTH_DATE,
                            account_status=db.ACCOUNT_VERIFIED)
        session.add_all([admin, mv, volunteer, other, inactive, pending])
        session.flush()
        team.mv_person_id = mv.id
        blocks = {block.phase: block for block in db.get_day_blocks(session, h.SEASON, "05.09.2026")}
        past = next(block for block in db.get_day_blocks(session, h.SEASON, "01.01.2020")
                    if block.phase == db.BLOCK_CAKE_DELIVERY)
        session.commit()
        ids = {name: person.id for name, person in {
            "admin": admin, "mv": mv, "volunteer": volunteer, "other": other,
            "inactive": inactive, "pending": pending,
        }.items()}
        ids.update(cake=blocks[db.BLOCK_CAKE_DELIVERY].id,
                   preparation=blocks[db.BLOCK_PREPARATION].id,
                   cleanup=blocks[db.BLOCK_CLEANUP].id, past=past.id,
                   team=team.id)
    return app, engine, ids


def _client(app, ids, person="admin"):
    client = app.test_client()
    h.sign_in(client, ids[person])
    return client


def _settings(client, block_id, quantity=4, time="10:00", expected_quantity=None,
              expected_time=None, csrf=True):
    return client.post(f"/api/blocks/{block_id}/cake-settings", json={
        "cake_quantity": quantity, "delivery_time": time,
        "expected_cake_quantity": expected_quantity, "expected_delivery_time": expected_time,
    }, headers=h.csrf_headers() if csrf else {})


def _assignment(client, block_id, person_id, slot=0, action="claim"):
    return client.post(f"/api/block-assignment/{action}", json={
        "block_id": block_id, "slot": slot, "person_id": person_id,
        "expected_person_id": None if action == "claim" else person_id,
    }, headers=h.csrf_headers())


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_setup_zero_and_four_cakes_are_distinct_saved_collapsed_states(_today):
    app, _engine, ids = _site()
    admin = _client(app, ids)
    tree = html.fromstring(admin.get("/").get_data(as_text=True))
    card = tree.get_element_by_id(f'block-{ids["cake"]}')
    assert card.get("open") is None and card.xpath("./summary")
    assert "Admin-Einrichtung erforderlich" in card.text_content()
    assert not card.xpath('.//select[@data-block-assignment]')
    assert card.xpath('.//div[@class="coverage"]')[0].get("hidden") is not None
    time_input = card.xpath('.//form[@data-cake-config]//input[@name="delivery_time"]')[0]
    assert time_input.get("type") == "text", "the delivery control must not inherit a locale-dependent AM/PM picker"
    assert time_input.get("placeholder") == "HH:MM"
    assert time_input.get("pattern") == "([01][0-9]|2[0-3]):[0-5][0-9]"
    assert "Lieferzeit (24 h)" in card.text_content()
    assert admin.get(f'/api/blocks/{ids["cake"]}/candidates').get_json()["people"] == []
    assert _assignment(admin, ids["cake"], ids["volunteer"]).status_code == 400
    configured = _settings(admin, ids["cake"]).get_json()["block"]
    assert configured["cake_quantity"] == 4 and configured["time"] == "10:00"
    assert len(configured["slots"]) == 4 and configured["progress"]["total"] == 4
    tree = html.fromstring(admin.get("/?date=05.09.2026").get_data(as_text=True))
    card = tree.get_element_by_id(f'block-{ids["cake"]}')
    assert len(card.xpath('.//select[@data-block-assignment]')) == 4
    assert "4 Kuchen angefragt" in card.xpath("./summary")[0].text_content()
    assert not card.xpath('.//*[contains(@class,"game-responsible")]')
    cards = tree.xpath('//details[contains(@class,"game-card")]')
    assert "task-block-preparation" in cards[0].get("class")
    assert "task-block-cake_delivery" in cards[1].get("class")
    assert "task-block-cleanup" in cards[-1].get("class")
    zero = _settings(admin, ids["cake"], 0, expected_quantity=4, expected_time="10:00")
    assert zero.status_code == 200 and zero.get_json()["block"]["slots"] == []
    card = html.fromstring(admin.get("/").get_data(as_text=True)).get_element_by_id(f'block-{ids["cake"]}')
    assert "Keine Kuchen angefragt" in card.text_content()
    assert card.xpath('.//div[@class="coverage"]')[0].get("hidden") is not None


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_settings_require_admin_csrf_valid_clock_and_strict_whole_quantity(_today):
    app, engine, ids = _site()
    assert _settings(app.test_client(), ids["cake"]).status_code == 401
    for person in ("mv", "volunteer", "other", "inactive", "pending"):
        assert _settings(_client(app, ids, person), ids["cake"]).status_code in (401, 403)
        assert "data-cake-config" not in _client(app, ids, person).get("/").get_data(as_text=True)
    admin = _client(app, ids)
    assert _settings(admin, ids["cake"], csrf=False).status_code == 403
    for quantity in (True, False, -1, 1.5, 2.0, "4", None):
        assert _settings(admin, ids["cake"], quantity).status_code == 400, quantity
    for time in ("24:00", "10:60", "9:00", "10:00 Uhr", "", None):
        assert _settings(admin, ids["cake"], time=time).status_code == 400, time
    assert _settings(admin, ids["preparation"]).status_code == 404
    missing_expected = admin.post(f'/api/blocks/{ids["cake"]}/cake-settings',
                                 json={"delivery_time": "10:00", "cake_quantity": 4},
                                 headers=h.csrf_headers())
    assert missing_expected.status_code == 400
    with h.Session(engine) as session:
        block = session.get(db.DayBlock, ids["cake"])
        assert block.cake_quantity is None and block.delivery_time is None
        assert not block.assignments


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_stale_settings_and_occupied_reduction_keep_saved_positions_and_history(_today):
    app, engine, ids = _site()
    admin = _client(app, ids)
    assert _settings(admin, ids["cake"]).status_code == 200
    claim = _assignment(admin, ids["cake"], ids["volunteer"], 3).get_json()
    assert claim["block"]["progress"]["filled"] == 1
    blocked = _settings(admin, ids["cake"], 2, "11:00", 4, "10:00")
    assert blocked.status_code == 400 and "freigegeben" in blocked.get_json()["error"]
    stale = _settings(admin, ids["cake"], 5)
    assert stale.status_code == 409
    current = stale.get_json()
    assert current["current_delivery_time"] == "10:00" and current["current_cake_quantity"] == 4
    assert current["block"]["progress"] == {"filled": 1, "total": 4, "percent": 25}
    assert current["block"]["slots"][3]["person_id"] == ids["volunteer"]
    assert _assignment(admin, ids["cake"], ids["volunteer"], 3, "release").status_code == 200
    assert _settings(admin, ids["cake"], 2, "11:00", 4, "10:00").status_code == 200
    assert _assignment(admin, ids["cake"], ids["other"], 3).status_code == 400
    with h.Session(engine) as session:
        assert session.query(db.BlockAssignment).filter_by(block_id=ids["cake"]).count() == 0
        audits = session.query(db.AssignmentAudit).filter_by(block_id=ids["cake"]).all()
        assert len(audits) == 2 and all("Kuchenlieferung 4" in entry.block_snapshot for entry in audits)
    history = admin.get("/audit").get_data(as_text=True)
    assert "Kuchenlieferung 4" in history and "Cake Only" in history


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_cake_candidates_and_claims_use_current_capacity_active_people_and_self_scope(_today):
    app, _engine, ids = _site()
    admin = _client(app, ids)
    _settings(admin, ids["cake"])
    _settings(admin, ids["past"])
    candidate_url = f'/api/blocks/{ids["cake"]}/candidates'
    assert app.test_client().get(candidate_url).status_code == 401
    member = _client(app, ids, "volunteer")
    mv = _client(app, ids, "mv")
    assert {person["id"] for person in member.get(candidate_url).get_json()["people"]} == {ids["volunteer"]}
    assert {person["id"] for person in mv.get(candidate_url).get_json()["people"]} == {ids["mv"]}
    assert _assignment(mv, ids["cake"], ids["other"]).status_code == 403
    assert _assignment(admin, ids["cake"], ids["inactive"]).status_code == 400
    assert _assignment(member, ids["cake"], ids["volunteer"]).status_code == 200
    duplicate = _assignment(member, ids["cake"], ids["volunteer"], 1)
    assert duplicate.status_code == 400
    assert _assignment(member, ids["preparation"], ids["volunteer"]).status_code == 200
    assert _assignment(member, ids["cleanup"], ids["volunteer"]).status_code == 200
    stale = _assignment(admin, ids["cake"], ids["other"])
    assert stale.status_code == 409 and stale.get_json()["current_person_id"] == ids["volunteer"]
    candidates = admin.get(candidate_url)
    assert candidates.headers["Cache-Control"] == "private, no-store"
    assert ids["volunteer"] not in candidates.get_json()["slots"]["1"]["candidate_ids"]
    assert "private-cake@fixture.test" not in candidates.get_data(as_text=True)
    assert "birth_date" not in candidates.get_data(as_text=True)
    assert _assignment(member, ids["past"], ids["volunteer"]).status_code == 403
    assert _assignment(admin, ids["past"], ids["volunteer"]).status_code == 200
    assert member.get(f'/api/blocks/{ids["past"]}/candidates').get_json()["slots"] == {}


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_guest_names_only_and_cake_person_team_date_filters(_today):
    app, _engine, ids = _site()
    admin = _client(app, ids)
    _settings(admin, ids["cake"])
    _assignment(admin, ids["cake"], ids["volunteer"])
    guest = app.test_client()
    page = guest.get("/?person=Cake%20Only&date=05.09.2026").get_data(as_text=True)
    assert "Cake Only" in page and "task-block-cake_delivery" in page
    assert "Nr. 1" not in page and "task-block-preparation" not in page and "task-block-cleanup" not in page
    assert "private-cake@fixture.test" not in page and "person_id" not in page
    assert "data-cake-config" not in page and "data-block-assignment" not in page
    assert "data-occupant-id" not in page and 'id="block-' not in page
    assert "data-candidate-url" not in page and f'value="{ids["volunteer"]}"' not in page
    playing = guest.get(f'/?playing_team=team-{ids["team"]}').get_data(as_text=True)
    assert "task-block-cake_delivery" in playing
    responsible = guest.get(f'/?responsible_team=team-{ids["team"]}').get_data(as_text=True)
    assert "task-block-cake_delivery" not in responsible
    assert "task-block-preparation" in guest.get("/").get_data(as_text=True)


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_statistics_report_setup_separately_from_configured_missing_cakes_and_zero(_today):
    app, _engine, ids = _site()
    admin = _client(app, ids)
    page = html.fromstring(admin.get("/statistik").get_data(as_text=True))
    cake_row = page.xpath('//table[contains(@class,"gap-table")]//tr[contains(.,"Kuchenlieferung")]')[0]
    assert "Admin-Einrichtung erforderlich" in cake_row.text_content()
    assert cake_row.xpath('.//div[@class="gap-list"]/span/text()') == ["Admin-Einrichtung erforderlich"], (
        "unconfigured cakes must not invent a missing-duty count"
    )
    _settings(admin, ids["cake"])
    _assignment(admin, ids["cake"], ids["volunteer"])
    _assignment(admin, ids["cake"], ids["other"], 1)
    page = html.fromstring(admin.get("/statistik").get_data(as_text=True))
    cake_row = page.xpath('//table[contains(@class,"gap-table")]//tr[contains(.,"Kuchenlieferung")]')[0]
    assert "Kuchenlieferung (2×)" in cake_row.text_content()
    assert "Admin-Einrichtung erforderlich" not in cake_row.text_content()
    assert "1× Kuchenlieferung" in page.text_content() and "Kuchenlieferung 1" not in page.text_content()
    _assignment(admin, ids["cake"], ids["volunteer"], action="release")
    _assignment(admin, ids["cake"], ids["other"], 1, "release")
    _settings(admin, ids["cake"], 0, expected_quantity=4, expected_time="10:00")
    page = html.fromstring(admin.get("/statistik").get_data(as_text=True))
    assert not page.xpath('//table[contains(@class,"gap-table")]//tr[contains(.,"Kuchenlieferung")]')


def test_cake_configuration_browser_state_regression():
    node = shutil.which("node")
    if node is None:
        return
    result = subprocess.run([node, "test/js_cake_blocks.mjs"], cwd=h.PROJECT_DIR,
                            text=True, capture_output=True, timeout=30)
    assert result.returncode == 0, result.stdout + result.stderr


if __name__ == "__main__":
    h.run_all(dict(globals()))
