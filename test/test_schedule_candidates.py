"""On-demand schedule candidates use current permissions and bounded queries."""

import os
import tempfile
from datetime import date
from unittest.mock import patch

from lxml import html
from sqlalchemy import event

import helpers as h
import db
import webapp


def _site():
    path = os.path.join(h._TEST_DIR, f"candidates-{next(tempfile._get_candidate_names())}.db")
    db.initialize_db(db.make_engine(path))
    previous = os.environ["NULIGAHELPER_DB"]
    os.environ["NULIGAHELPER_DB"] = path
    try:
        app = webapp.create_app()
    finally:
        os.environ["NULIGAHELPER_DB"] = previous
    engine = app.extensions["nuligahelper_auth_abuse"].engine
    with h.Session(engine) as session:
        db.sync_games(session, [
            {"day": "Sa", "date": "05.09.2026", "time": "12:00", "hall": 1,
             "game_nr": "1", "ak": "BL mD", "home": "TuS", "guest": "A", "score": ""},
            {"day": "Sa", "date": "01.01.2020", "time": "12:00", "hall": 1,
             "game_nr": "2", "ak": "BL F", "home": "TuS", "guest": "B", "score": ""},
        ], h.SEASON)
        team = session.query(db.Team).filter_by(name="BL mD").one()
        other_team = session.query(db.Team).filter_by(name="BL F").one()
        support = db.get_support_team(session)
        admin = db.Person(name="Admin", email="admin@fixture.test", is_admin=True)
        mv = db.Person(name="MV", email="mv@fixture.test", teams=[team])
        member = db.Person(name="Member", email="member@fixture.test", teams=[team, support])
        outsider = db.Person(name="Outsider", email="outsider@fixture.test", teams=[other_team])
        pending = db.Person(name="Pending", email="pending@fixture.test", account_status=db.ACCOUNT_VERIFIED)
        inactive = db.Person(name="Inactive", email="inactive@fixture.test", account_status=db.ACCOUNT_INACTIVE)
        session.add_all([admin, mv, member, outsider, pending, inactive])
        session.flush()
        team.mv_person_id = mv.id
        future = session.query(db.Game).filter_by(game_nr="1").one()
        past = session.query(db.Game).filter_by(game_nr="2").one()
        future.team_id = team.id
        db.claim_slot(session, future, db.ROLE_TIMEKEEPER, 0, None, member)
        db.claim_slot(session, past, db.ROLE_TIMEKEEPER, 0, None, outsider)
        outsider.account_status = db.ACCOUNT_INACTIVE
        block = db.get_day_blocks(session, h.SEASON, future.date)[0]
        past_block = db.get_day_blocks(session, h.SEASON, past.date)[0]
        session.commit()
        ids = {name: person.id for name, person in {
            "admin": admin, "mv": mv, "member": member, "outsider": outsider,
            "pending": pending,
        }.items()}
        ids.update(future=future.id, past=past.id, block=block.id,
                   past_block=past_block.id,
                   team=team.id, other_team=other_team.id)
    return app, engine, ids


def _signed_in(app, person_id):
    client = app.test_client()
    h.sign_in(client, person_id)
    return client


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_initial_schedule_size_does_not_grow_with_unassigned_roster(_today):
    app, engine, ids = _site()
    admin = _signed_in(app, ids["admin"])
    before = admin.get("/").get_data(as_text=True)
    with h.Session(engine) as session:
        support = db.get_support_team(session)
        session.add_all([
            db.Person(name=f"Extra {index:03}", email=f"extra{index}@fixture.test", teams=[support])
            for index in range(120)
        ])
        session.commit()
    after = admin.get("/").get_data(as_text=True)
    assert len(after) == len(before), "unassigned roster growth changed initial schedule size"
    assert "Extra 119" not in after
    assert len(html.fromstring(after).xpath('//select[@data-role]//option')) == len(
        html.fromstring(before).xpath('//select[@data-role]//option')
    )
    assert "Member · BL mD, Supporter" in after, "the assigned occupant must stay visible"


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_candidate_reads_enforce_current_tier_and_do_not_expose_contacts(_today):
    app, engine, ids = _site()
    url = f'/api/games/{ids["future"]}/candidates'
    assert app.test_client().get(url).status_code == 401
    assert _signed_in(app, ids["pending"]).get(url).status_code == 403

    member = _signed_in(app, ids["member"])
    member_payload = member.get(url).get_json()
    assert {person["id"] for person in member_payload["people"]} == {ids["member"]}
    assert f'{db.ROLE_TIMEKEEPER}:0' in member_payload["slots"]
    assert member_payload["slots"][f'{db.ROLE_TIMEKEEPER}:0']["occupant_id"] == ids["member"]

    mv = _signed_in(app, ids["mv"])
    mv_payload = mv.get(url).get_json()
    assert {person["id"] for person in mv_payload["people"]} == {ids["mv"], ids["member"]}
    assert {person["id"] for person in mv.get(f'/api/blocks/{ids["block"]}/candidates').get_json()["people"]} == {ids["mv"]}

    with h.Session(engine) as session:
        session.get(db.Team, ids["team"]).mv_person_id = None
        session.commit()
    assert {person["id"] for person in mv.get(url).get_json()["people"]} == {ids["mv"]}

    admin = _signed_in(app, ids["admin"])
    response = admin.get(url)
    assert response.headers["Cache-Control"] == "private, no-store"
    assert {person["id"] for person in response.get_json()["people"]} == {
        ids["admin"], ids["mv"], ids["member"],
    }
    assert "@fixture.test" not in response.get_data(as_text=True)
    admin_block = admin.get(f'/api/blocks/{ids["past_block"]}/candidates')
    assert admin_block.status_code == 200 and admin_block.get_json()["slots"]
    assert not member.get(f'/api/blocks/{ids["past_block"]}/candidates').get_json()["slots"]
    assert admin.get(f'/api/games/{ids["past"]}/candidates').get_json()["slots"][f'{db.ROLE_TIMEKEEPER}:0']["occupant_id"] == ids["outsider"]
    assert ids["outsider"] not in {person["id"] for person in admin.get(f'/api/games/{ids["past"]}/candidates').get_json()["people"]}
    assert not member.get(f'/api/games/{ids["past"]}/candidates').get_json()["slots"]


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_candidate_query_count_is_bounded_when_roster_grows(_today):
    app, engine, ids = _site()
    admin = _signed_in(app, ids["admin"])
    url = f'/api/games/{ids["future"]}/candidates'
    statements = []

    def count(_connection, _cursor, statement, _parameters, _context, _executemany):
        statements.append(statement)

    event.listen(engine, "before_cursor_execute", count)
    try:
        admin.get(url)
        small_count = len(statements)
        with h.Session(engine) as session:
            team = session.get(db.Team, ids["team"])
            session.add_all([
                db.Person(name=f"More {index:03}", teams=[team])
                for index in range(100)
            ])
            session.commit()
        statements.clear()
        response = admin.get(url)
        large_count = len(statements)
    finally:
        event.remove(engine, "before_cursor_execute", count)
    assert response.status_code == 200
    assert len(response.get_json()["people"]) >= 100
    assert large_count <= small_count + 2, (small_count, large_count)
    assert large_count < 20, "candidate loading issued one query per person or slot"


if __name__ == "__main__":
    h.run_all(dict(globals()))
