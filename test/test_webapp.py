"""Authenticated webapp click-through scenario against a throwaway database."""

import re
import os
import tempfile
from datetime import datetime, timedelta

import helpers as h
import db
import webapp

_previous_db = os.environ["NULIGAHELPER_DB"]
_db_path = os.path.join(h._TEST_DIR, f"web-{next(tempfile._get_candidate_names())}.db")
os.environ["NULIGAHELPER_DB"] = _db_path
try:
    app = webapp.create_app()
finally:
    os.environ["NULIGAHELPER_DB"] = _previous_db
ENGINE = db.make_engine(_db_path)
client = app.test_client()

with h.Session(ENGINE) as session:
    games = h.sample_games()
    original_1001 = next(game for game in games if game["game_nr"] == 1001)
    games.append({
        **original_1001,
        "source_key": "test:1001-duplicate",
        "time": "16:30",
        "guest": "Duplicate Tournament Team",
    })
    db.sync_games(session, games, h.SEASON)
    game_data = next(game for game in games if game["game_nr"] == 1001)
    playing = session.query(db.Team).filter_by(name=game_data["ak"]).one()
    responsible = session.query(db.Team).filter(db.Team.id != playing.id).first()
    unrelated = session.query(db.Team).filter(
        db.Team.id.notin_([playing.id, responsible.id])
    ).first()
    support = db.get_support_team(session)
    admin = db.Person(
        name="Admin Test", email="admin@example.test", team=support, is_admin=True
    )
    alice = db.Person(name="Alex Test", email="alex@example.test", team=playing)
    duplicate = db.Person(name="Alex Test", phone="+491700000002", team=responsible)
    outsider = db.Person(name="Outside Test", team=unrelated)
    session.add_all([admin, alice, duplicate, outsider])
    session.commit()
    game = session.query(db.Game).filter_by(source_key="test:1001").one()
    duplicate_game = session.query(db.Game).filter_by(
        source_key="test:1001-duplicate"
    ).one()
    GAME_ID = game.id
    DUPLICATE_GAME_ID = duplicate_game.id
    ADMIN_ID, ALICE_ID, DUPLICATE_ID, OUTSIDER_ID = (
        admin.id, alice.id, duplicate.id, outsider.id
    )
    RESPONSIBLE_ID, SUPPORT_ID = responsible.id, support.id

CSRF = h.sign_in(client, ADMIN_ID)


def _json(url, body):
    return client.post(url, json=body, headers=h.csrf_headers(CSRF))


def _form(url, body=None, **kwargs):
    return client.post(url, data=h.csrf_data(body, CSRF), **kwargs)


def test_01_admin_sign_in_exposes_controls_without_contacts_on_schedule():
    page = client.get("/").get_data(as_text=True)
    assert page.count('class="team-select"') == len(games)
    assert "admin@example.test" not in page and "+4917" not in page
    assert "Alex Test ·" in page, "duplicate names must be qualified by team"


def test_02_responsible_team_and_claim_release_flow():
    response = _json(f"/api/games/{GAME_ID}/team", {"team_id": RESPONSIBLE_ID})
    assert response.get_json() == {"ok": True}
    claim = _json("/api/assignment/claim", {
        "game_id": GAME_ID,
        "role": db.ROLE_SALE,
        "slot": 0,
        "expected_person_id": None,
        "person_id": ALICE_ID,
    })
    assert claim.status_code == 200 and claim.get_json()["ok"]
    conflict = _json("/api/assignment/claim", {
        "game_id": GAME_ID,
        "role": db.ROLE_SALE,
        "slot": 0,
        "expected_person_id": None,
        "person_id": DUPLICATE_ID,
    })
    assert conflict.status_code == 409
    assert conflict.get_json()["current_person_id"] == ALICE_ID
    release = _json("/api/assignment/release", {
        "game_id": GAME_ID,
        "role": db.ROLE_SALE,
        "slot": 0,
        "expected_person_id": ALICE_ID,
    })
    assert release.get_json() == {"ok": True}


def test_03_existing_rules_and_advisory_warning_remain():
    first = _json("/api/assignment/claim", {
        "game_id": GAME_ID, "role": db.ROLE_TIMEKEEPER, "slot": 0,
        "expected_person_id": None, "person_id": ALICE_ID,
    })
    assert first.get_json()["warning"] == "Person spielt selbst in diesem Spiel."
    second = _json("/api/assignment/claim", {
        "game_id": GAME_ID, "role": db.ROLE_SECRETARY, "slot": 0,
        "expected_person_id": None, "person_id": ALICE_ID,
    })
    assert second.status_code == 400 and "bereits" in second.get_json()["error"]
    outside = _json("/api/assignment/claim", {
        "game_id": GAME_ID, "role": db.ROLE_SECURITY, "slot": 0,
        "expected_person_id": None, "person_id": OUTSIDER_ID,
    })
    assert outside.get_json()["warning"] == "Person gehört nicht zum verantwortlichen Team."


def test_03_sparse_sale_slot_keeps_its_stored_position():
    response = _json("/api/assignment/claim", {
        "game_id": GAME_ID, "role": db.ROLE_SALE, "slot": 1,
        "expected_person_id": None, "person_id": DUPLICATE_ID,
    })
    assert response.status_code == 200
    page = client.get("/").get_data(as_text=True)
    card_start = page.index(f'id="game-{GAME_ID}"')
    card = page[card_start:page.index("</article>", card_start)]
    sales = re.findall(
        r'<select[^>]*data-role="Verkauf"[^>]*>.*?</select>', card, re.S
    )
    assert len(sales) == 2
    selected = rf'<option value="{DUPLICATE_ID}"[^>]*selected'
    assert not re.search(selected, sales[0])
    assert re.search(selected, sales[1])


def test_04_person_crud_uses_internal_identity():
    response = _form("/personen/add", {
        "name": "Alex Test", "team_id": SUPPORT_ID, "email": "third@example.test"
    }, follow_redirects=True)
    assert response.status_code == 200
    with h.Session(ENGINE) as session:
        matches = session.query(db.Person).filter_by(name="Alex Test").all()
        assert len(matches) == 3
        created = next(person for person in matches if person.email == "third@example.test")
        created_id = created.id
    _form(f"/personen/{created_id}/edit", {
        "name": "Renamed Test", "team_id": SUPPORT_ID, "phone": "+491701234568"
    })
    with h.Session(ENGINE) as session:
        person = session.get(db.Person, created_id)
        assert person.name == "Renamed Test" and person.phone == "+491701234568"


def test_05_team_mv_and_statistics_are_available_to_admin():
    response = _json(f"/api/teams/{RESPONSIBLE_ID}/mv", {
        "person_id": DUPLICATE_ID
    })
    assert response.get_json() == {"ok": True}
    statistics = client.get("/statistik")
    assert statistics.status_code == 200
    statistics_page = statistics.get_data(as_text=True)
    assert "Offene Dienste" in statistics_page
    assert 'class="data-table responsive-table stat-table' in statistics_page
    assert 'data-label="Person"' in statistics_page


def test_06_audit_is_newest_first_and_read_only():
    with h.Session(ENGINE) as session:
        game = session.get(db.Game, GAME_ID)
        now = datetime.now()
        common = dict(
            actor_person_id=ADMIN_ID, actor_tier="admin", action="claim",
            affected_person_id=ALICE_ID, game_id=GAME_ID,
            role=db.ROLE_SALE, slot=0, affected_person_name="Alex Test",
            game_snapshot=f"{game.game_nr} marker",
        )
        session.add_all([
            db.AssignmentAudit(changed_at=now - timedelta(minutes=1), actor_name="Older Marker", **common),
            db.AssignmentAudit(changed_at=now, actor_name="Newer Marker", **common),
        ])
        session.commit()
    page = client.get("/audit").get_data(as_text=True)
    assert "Änderungsprotokoll" in page and "Alex Test" in page
    assert 'class="audit-filter-card"' in page
    assert 'class="data-table responsive-table audit-table"' in page
    for label in ("Zeit", "Akteur", "Aktion", "Person", "Spiel", "Dienst"):
        assert f'data-label="{label}"' in page
    audit_table = page[
        page.index('<table class="data-table responsive-table audit-table"'):
        page.index("</table>")
    ]
    assert "<form" not in audit_table and "data-delete" not in audit_table
    assert page.index("Newer Marker") < page.index("Older Marker")
    assert "15:00 · BL mD · TuS Raubling – SBC Traunstein" in page
    assert "16:30 · BL mD · TuS Raubling – Duplicate Tournament Team" in page
    assert "1001 marker" in client.get(
        f"/audit?game_id={GAME_ID}"
    ).get_data(as_text=True)
    assert "Alex Test" in client.get(
        f"/audit?person_id={ALICE_ID}"
    ).get_data(as_text=True)
    assert client.post("/audit/1/delete", headers=h.csrf_headers(CSRF)).status_code == 404


def test_07_guest_schedule_has_no_roster_payload():
    guest = app.test_client()
    page = guest.get("/").get_data(as_text=True)
    assert 'data-role="' not in page and 'class="team-select"' not in page
    assert "third@example.test" not in page and "+4917" not in page
    assert "Renamed Test" not in page, "unassigned roster members must not leak"
    assert "Alex Test" in page, "assigned helper names remain public"
    assert f'id="game-{GAME_ID}"' not in page
    assert f'id="game-{DUPLICATE_GAME_ID}"' not in page


def test_08_csrf_is_required_for_json_and_forms():
    assert client.post(
        f"/api/games/{GAME_ID}/team", json={"team_id": None}
    ).status_code == 403
    assert client.post(
        "/personen/add", data={"name": "Forged", "team_id": SUPPORT_ID}
    ).status_code == 403
    with open(h.PROJECT_DIR + "/static/app.js", encoding="utf-8") as source:
        javascript = source.read()
    assert "response.status === 401" in javascript
    assert "response.status === 409" in javascript
    assert "Deine Sitzung ist abgelaufen" in javascript


def test_09_delete_warns_about_deactivation_and_cascades_with_audit():
    page = client.get("/personen").get_data(as_text=True)
    assert "Deaktivieren" in page
    assert "stattdessen deaktivieren" in open(
        h.PROJECT_DIR + "/static/app.js", encoding="utf-8"
    ).read()
    _form(f"/personen/{DUPLICATE_ID}/delete", follow_redirects=True)
    with h.Session(ENGINE) as session:
        assert session.get(db.Person, DUPLICATE_ID) is None
        assert session.query(db.AssignmentAudit).count() >= 3


if __name__ == "__main__":
    h.run_all(dict(globals()))
