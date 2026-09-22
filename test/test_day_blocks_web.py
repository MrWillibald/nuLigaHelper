"""Web authorization, privacy, filtering and CAS responses for day blocks."""

import os
import tempfile
from datetime import date
from pathlib import Path
from unittest.mock import patch

import helpers as h
import db
import webapp


previous_db = os.environ["NULIGAHELPER_DB"]
database_path = os.path.join(
    h._TEST_DIR, f"day-block-web-{next(tempfile._get_candidate_names())}.db"
)
os.environ["NULIGAHELPER_DB"] = database_path
try:
    db.initialize_db(db.make_engine(database_path))
    app = webapp.create_app()
finally:
    os.environ["NULIGAHELPER_DB"] = previous_db
ENGINE = db.make_engine(database_path)

with h.Session(ENGINE) as session:
    rows = [
        {
            "day": "Sa", "date": "05.09.2026", "time": "10:00",
            "hall": 1, "game_nr": "1", "ak": "BL mD",
            "home": "TuS", "guest": "A", "score": "",
        },
        {
            "day": "Sa", "date": "05.09.2026", "time": "18:00",
            "hall": 1, "game_nr": "2", "ak": "BL M",
            "home": "TuS", "guest": "B", "score": "",
        },
        {
            "day": "Mi", "date": "01.01.2020", "time": "10:00",
            "hall": 1, "game_nr": "3", "ak": "BL mD",
            "home": "TuS", "guest": "Alt", "score": "",
        },
    ]
    db.sync_games(session, rows, h.SEASON)
    team = session.query(db.Team).filter_by(name="BL mD").one()
    admin = db.Person(name="Admin", email="admin@example.test", is_admin=True)
    mv = db.Person(name="MV", email="mv@example.test", teams=[team])
    member = db.Person(name="Block Match", email="member@example.test", teams=[team])
    other = db.Person(name="Other", email="other@example.test")
    inactive = db.Person(
        name="Inactive", email="inactive@example.test",
        account_status=db.ACCOUNT_INACTIVE,
    )
    session.add_all([admin, mv, member, other, inactive])
    session.flush()
    team.mv_person_id = mv.id
    future = db.get_day_blocks(session, h.SEASON, "05.09.2026")[0]
    past = db.get_day_blocks(session, h.SEASON, "01.01.2020")[0]
    db.claim_block_slot(session, future, 0, None, member)
    session.commit()
    IDS = {
        "admin": admin.id, "mv": mv.id, "member": member.id,
        "other": other.id, "inactive": inactive.id,
        "future": future.id, "past": past.id,
    }


def _client(person):
    client = app.test_client()
    token = h.sign_in(client, IDS[person])
    return client, token


def _post(client, token, action, block_id, slot, expected, person_id=None):
    payload = {
        "block_id": block_id, "slot": slot, "expected_person_id": expected,
    }
    if person_id is not None:
        payload["person_id"] = person_id
    return client.post(
        f"/api/block-assignment/{action}", json=payload,
        headers=h.csrf_headers(token),
    )


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_guest_sees_names_and_cards_without_ids_contacts_roster_or_controls(_today):
    page = app.test_client().get("/").get_data(as_text=True)
    preparation = page.index("task-block-preparation")
    first_game = page.index("Nr. 1")
    second_game = page.index("Nr. 2")
    cleanup = page.index("task-block-cleanup")
    assert preparation < first_game < second_game < cleanup
    assert "↗" in page[preparation:first_game]
    assert "↘" in page[cleanup:]
    assert "Block Match" in page
    assert "member@example.test" not in page
    assert f'value="{IDS["member"]}"' not in page
    assert "data-block-assignment" not in page and "person_id" not in page
    assert "Verantwortlich" not in page[preparation:first_game]


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_block_name_filter_keeps_whole_date_and_full_boundary_times(_today):
    page = app.test_client().get("/?person=Block%20Match").get_data(as_text=True)
    assert "Nr. 1" in page and "Nr. 2" in page
    assert "08:30" in page and "19:00" in page


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_member_and_mv_are_self_only_and_conflicts_report_current_occupant(_today):
    member, member_token = _client("member")
    assert _post(
        member, member_token, "claim", IDS["future"], 1, None, IDS["other"]
    ).status_code == 403
    stale = _post(member, member_token, "claim", IDS["future"], 0, None, IDS["member"])
    assert stale.status_code == 409
    assert stale.get_json()["current_person_id"] == IDS["member"]
    mv, mv_token = _client("mv")
    assert _post(
        mv, mv_token, "claim", IDS["future"], 1, None, IDS["other"]
    ).status_code == 403


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_admin_can_manage_active_people_and_past_blocks_but_not_inactive(_today):
    admin, token = _client("admin")
    inactive = _post(
        admin, token, "claim", IDS["future"], 1, None, IDS["inactive"]
    )
    assert inactive.status_code == 400
    assert _post(
        admin, token, "claim", IDS["future"], 1, None, IDS["other"]
    ).get_json() == {"ok": True, "block_id": IDS["future"]}
    assert _post(
        admin, token, "claim", IDS["past"], 0, None, IDS["other"]
    ).status_code == 200


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_non_admin_cannot_change_past_block_and_csrf_is_required(_today):
    member, token = _client("member")
    assert _post(
        member, token, "claim", IDS["past"], 1, None, IDS["member"]
    ).status_code == 403
    response = member.post("/api/block-assignment/claim", json={
        "block_id": IDS["future"], "slot": 2,
        "expected_person_id": None, "person_id": IDS["member"],
    })
    assert response.status_code == 403


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_statistics_count_occupied_block_slots_and_report_open_block_slots(_today):
    admin, _token = _client("admin")
    page = admin.get("/statistik").get_data(as_text=True)
    assert "Block Match" in page and "Vorbereitung" in page
    assert "Aufräumen" in page
    assert "Vorbereitung 1" not in page and "Vorbereitung 3" not in page
    assert "Aufräumen 1" not in page
    assert "Unterstützung" not in page, "empty optional game duty is not a gap"


def test_block_colors_follow_preparation_and_cleanup_direction():
    css = Path(h.PROJECT_DIR, "static", "style.css").read_text(encoding="utf-8")
    block_css = css[css.index(".task-block-card"):css.index("@keyframes savedflash")]
    assert ".task-block-card{--block-accent:var(--navy)}" in block_css
    assert ".task-block-preparation{--block-accent:#a0e656}" in block_css
    assert ".task-block-cleanup{--block-accent:#ffb752}" in block_css
    assert "background:var(--block-accent)" in block_css
    assert "linear-gradient" not in block_css
    assert "border-left" not in block_css
    assert ".task-block-no-time{background:" not in block_css


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_audit_view_lists_block_target_filter_and_snapshot(_today):
    admin, _token = _client("admin")
    page = admin.get("/audit").get_data(as_text=True)
    assert f'value="block-{IDS["future"]}"' in page
    assert "05.09.2026 | Vorbereitung" in page
    assert "Block Match" in page


if __name__ == "__main__":
    h.run_all(dict(globals()))
