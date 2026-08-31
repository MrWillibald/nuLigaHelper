"""Authorization refusal scenario across guest, member, MV and admin tiers."""

import os
import tempfile

import helpers as h
import db
import webapp

_previous_db = os.environ["NULIGAHELPER_DB"]
_db_path = os.path.join(h._TEST_DIR, f"refusal-{next(tempfile._get_candidate_names())}.db")
os.environ["NULIGAHELPER_DB"] = _db_path
try:
    app = webapp.create_app()
finally:
    os.environ["NULIGAHELPER_DB"] = _previous_db
ENGINE = db.make_engine(_db_path)
app.add_url_rule("/unlisted-test-route", "unlisted_test_route", lambda: "unsafe")

with h.Session(ENGINE) as session:
    h.sync_sample_games(session)
    teams = [team for team in db.get_all_teams(session) if not team.is_support]
    own_team, other_team = teams[:2]
    support = db.get_support_team(session)
    admin = db.Person(name="Admin", email="admin@x.test", team=support, is_admin=True)
    mv = db.Person(name="MV", email="mv@x.test", team=own_team)
    member = db.Person(name="Member", email="member@x.test", team=own_team)
    other = db.Person(name="Other", email="private@x.test", team=other_team)
    pending_other = db.Person(
        name="Pending Other", email="pending@x.test", desired_team=other_team,
        account_status=db.ACCOUNT_VERIFIED,
    )
    pending_support = db.Person(
        name="Pending Support", email="support@x.test", desired_team=support,
        account_status=db.ACCOUNT_VERIFIED,
    )
    session.add_all([admin, mv, member, other, pending_other, pending_support])
    session.flush()
    own_team.mv_person_id = mv.id
    games = session.query(db.Game).all()
    own_game, other_game = games[:2]
    own_game.team_id = own_team.id
    other_game.team_id = other_team.id
    past_game = db.Game(
        season_year=h.SEASON, game_nr=9900, date="01.01.2020", team=own_team
    )
    session.add(past_game)
    db.assign_person(session, own_game, other, db.ROLE_CLEANING)
    session.commit()
    IDS = {
        "admin": admin.id, "mv": mv.id, "member": member.id, "other": other.id,
        "pending_other": pending_other.id, "pending_support": pending_support.id,
        "own_game": own_game.id, "other_game": other_game.id, "past_game": past_game.id,
        "own_team": own_team.id,
    }


def _client(person):
    client = app.test_client()
    token = h.sign_in(client, IDS[person])
    return client, token


def _claim(client, token, game_id, person_id, role=db.ROLE_TIMEKEEPER):
    return client.post("/api/assignment/claim", json={
        "game_id": game_id, "role": role, "slot": 0,
        "expected_person_id": None, "person_id": person_id,
    }, headers=h.csrf_headers(token))


def _release(client, token, game_id, person_id, role, slot=0):
    return client.post("/api/assignment/release", json={
        "game_id": game_id, "role": role, "slot": slot,
        "expected_person_id": person_id,
    }, headers=h.csrf_headers(token))


def test_01_guest_and_unlisted_routes_fail_closed():
    guest = app.test_client()
    assert guest.get("/").status_code == 200
    assert guest.get("/personen").status_code == 302
    assert guest.get("/statistik").status_code == 302
    assert guest.get("/unlisted-test-route").status_code == 302
    response = guest.post("/api/assignment/claim", json={})
    assert response.status_code == 401 and response.get_json()["code"] == "session_expired"


def test_02_member_sees_no_foreign_contacts_and_admin_writes_are_refused():
    client, token = _client("member")
    roster = client.get("/personen").get_data(as_text=True)
    assert "member@x.test" in roster and "private@x.test" not in roster
    schedule = client.get("/").get_data(as_text=True)
    option_ids = set(int(value) for value in __import__("re").findall(
        r'<option value="(\d+)"', schedule
    ))
    assert option_ids <= {IDS["member"]}, "members may only receive their own person option"
    assert client.post("/personen/add", data=h.csrf_data({
        "name": "Forged", "team_id": IDS["own_team"]
    }, token)).status_code == 403
    assert client.post(f"/personen/{IDS['other']}/edit", data=h.csrf_data({
        "name": "Forged"
    }, token)).status_code == 403
    collision = client.post(f"/personen/{IDS['member']}/edit", data=h.csrf_data({
        "name": "Member", "email": "private@x.test"
    }, token), follow_redirects=True)
    assert "bereits verwendet" in collision.get_data(as_text=True)
    with h.Session(ENGINE) as session:
        assert session.get(db.Person, IDS["member"]).email == "member@x.test"
    assert client.post(f"/api/games/{IDS['own_game']}/team", json={
        "team_id": None
    }, headers=h.csrf_headers(token)).status_code == 403
    assert client.post(
        f"/personen/{IDS['other']}/deactivate",
        data=h.csrf_data(token=token),
    ).status_code == 403
    assert _claim(client, token, IDS["own_game"], IDS["other"]).status_code == 403
    assert _release(
        client, token, IDS["own_game"], IDS["other"], db.ROLE_CLEANING
    ).status_code == 403
    assert _claim(client, token, IDS["own_game"], IDS["member"]).status_code == 200
    assert _release(
        client, token, IDS["own_game"], IDS["member"], db.ROLE_TIMEKEEPER
    ).status_code == 200
    assert _claim(client, token, IDS["own_game"], IDS["member"]).status_code == 200


def test_03_mv_scope_requires_own_team_and_responsible_game():
    client, token = _client("mv")
    schedule = client.get("/").get_data(as_text=True)
    own_start = schedule.index(f'id="game-{IDS["own_game"]}"')
    own_card = schedule[own_start:schedule.index("</article>", own_start)]
    other_start = schedule.index(f'id="game-{IDS["other_game"]}"')
    other_card = schedule[other_start:schedule.index("</article>", other_start)]
    assert f'<option value="{IDS["member"]}"' in own_card
    assert f'<option value="{IDS["member"]}"' not in other_card
    assert f'<option value="{IDS["mv"]}"' in other_card, "MV keeps ordinary self-service rights"
    assert _release(
        client, token, IDS["own_game"], IDS["member"], db.ROLE_TIMEKEEPER
    ).status_code == 200
    assert _claim(
        client, token, IDS["own_game"], IDS["mv"], db.ROLE_SECRETARY
    ).status_code == 200
    assert _claim(
        client, token, IDS["own_game"], IDS["other"], db.ROLE_SECURITY
    ).status_code == 403
    assert _release(
        client, token, IDS["own_game"], IDS["other"], db.ROLE_CLEANING
    ).status_code == 403
    assert _claim(
        client, token, IDS["other_game"], IDS["member"], db.ROLE_SECURITY
    ).status_code == 403
    assert client.post(
        f"/registrierungen/{IDS['pending_other']}/approve",
        data=h.csrf_data(token=token),
    ).status_code == 403
    assert client.post(
        f"/personen/{IDS['other']}/deactivate",
        data=h.csrf_data(token=token),
    ).status_code == 403


def test_04_admin_can_approve_fallback_and_correct_past_game():
    client, token = _client("admin")
    approved = client.post(
        f"/registrierungen/{IDS['pending_support']}/approve",
        data=h.csrf_data(token=token),
    )
    assert approved.status_code == 302
    assert _claim(client, token, IDS["past_game"], IDS["other"]).status_code == 200
    deactivated = client.post(
        f"/personen/{IDS['other']}/deactivate",
        data=h.csrf_data(token=token),
    )
    assert deactivated.status_code == 302
    assert "Other" in client.get("/statistik").get_data(as_text=True), \
        "past work by inactive people must remain in statistics"


def test_05_member_and_mv_cannot_rewrite_past_games():
    member, member_token = _client("member")
    mv, mv_token = _client("mv")
    assert _claim(member, member_token, IDS["past_game"], IDS["member"]).status_code == 403
    assert _claim(mv, mv_token, IDS["past_game"], IDS["mv"]).status_code == 403


def test_06_admin_and_mv_rights_are_derived_together():
    with h.Session(ENGINE) as session:
        admin = session.get(db.Person, IDS["admin"])
        team = session.get(db.Team, IDS["own_team"])
        team.mv_person_id = admin.id
        session.commit()
    client, _ = _client("admin")
    assert client.get("/audit").status_code == 200
    page = client.get("/").get_data(as_text=True)
    assert 'class="team-select"' in page, "admin rights remain when the person is also MV"


if __name__ == "__main__":
    h.run_all(dict(globals()))
