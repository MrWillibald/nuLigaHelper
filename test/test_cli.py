"""Management CLI uses stable person IDs and bootstraps administrators."""

import contextlib
import io
import os
import tempfile
import subprocess

import helpers as h
import db
import manage_db
import main
import webapp


def _run(path, *arguments):
    args = manage_db.build_parser().parse_args(["--db", path, *map(str, arguments)])
    output = io.StringIO()
    with contextlib.redirect_stdout(output):
        args.func(args)
    return output.getvalue()


def test_duplicate_names_are_discovered_and_selected_by_id():
    path = os.path.join(h._TEST_DIR, f"cli-{next(tempfile._get_candidate_names())}.db")
    engine = db.make_engine(path)
    db.init_db(engine)
    with h.Session(engine) as session:
        team = db.get_support_team(session)
        first = db.Person(name="Same Name", team=team)
        second = db.Person(name="Same Name", team=team)
        first_game = db.Game(
            season_year=h.SEASON, source_key="meeting:101", game_nr=1234,
            date="30.12.2099", time="10:00", ak="GE", home="Home", guest="Team A",
        )
        second_game = db.Game(
            season_year=h.SEASON, source_key="meeting:102", game_nr=1234,
            date="31.12.2099", time="11:00", ak="GE", home="Home", guest="Team B",
        )
        session.add_all([first, second, first_game, second_game])
        session.commit()
        first_id, second_id = first.id, second.id
        first_game_id, second_game_id = first_game.id, second_game.id

    listing = _run(path, "search-person", "Same")
    assert f"ID {first_id}" in listing and f"ID {second_id}" in listing
    games = _run(path, "--season", h.SEASON, "list-games", "--number", 1234)
    assert f"ID {first_game_id}" in games and f"ID {second_game_id}" in games
    assert "Team A" in games and "Team B" in games
    _run(
        path, "--season", h.SEASON, "assign", second_game_id,
        db.ROLE_TIMEKEEPER, second_id,
    )
    with h.Session(engine) as session:
        assignment = session.query(db.Assignment).one()
        assert assignment.person_id == second_id and assignment.game_id == second_game_id
    _run(path, "--season", h.SEASON, "set-jteam", first_game_id, "Supporter")
    with h.Session(engine) as session:
        assert session.get(db.Game, first_game_id).team_id is not None
        assert session.get(db.Game, second_game_id).team_id is None
    _run(
        path, "--season", h.SEASON, "unassign", second_game_id,
        db.ROLE_TIMEKEEPER, second_id,
    )
    with h.Session(engine) as session:
        assert session.query(db.Assignment).count() == 0


def test_grant_admin_changes_the_derived_account_fact():
    path = os.path.join(h._TEST_DIR, f"cli-{next(tempfile._get_candidate_names())}.db")
    engine = db.make_engine(path)
    db.init_db(engine)
    with h.Session(engine) as session:
        person = db.Person(name="Bootstrap", email="admin@example.test")
        session.add(person)
        session.commit()
        person_id = person.id
    _run(path, "grant-admin", person_id)
    with h.Session(engine) as session:
        assert session.get(db.Person, person_id).is_admin is True
    previous = os.environ["NULIGAHELPER_DB"]
    os.environ["NULIGAHELPER_DB"] = path
    try:
        app = webapp.create_app()
    finally:
        os.environ["NULIGAHELPER_DB"] = previous
    client = app.test_client()
    h.sign_in(client, person_id)
    assert client.get("/audit").status_code == 200


def test_contact_preflight_reports_issues_without_writing():
    path = os.path.join(h._TEST_DIR, f"cli-{next(tempfile._get_candidate_names())}.db")
    engine = db.make_engine(path)
    db.init_db(engine)
    with h.Session(engine) as session:
        session.add_all([
            db.Person(name="Changed", email=" HELPER@Example.Test "),
            db.Person(name="Collision", email="helper@example.test"),
            db.Person(name="Invalid", phone="not-a-number"),
        ])
        session.commit()
        before = [
            (person.id, person.email, person.phone)
            for person in session.query(db.Person).order_by(db.Person.id)
        ]

    output = _run(path, "contact-preflight")
    assert "CHANGED" in output
    assert "COLLISION" in output
    assert "INVALID" in output

    with h.Session(engine) as session:
        after = [
            (person.id, person.email, person.phone)
            for person in session.query(db.Person).order_by(db.Person.id)
        ]
    assert after == before, "preflight must never canonicalize or mutate records"


def test_launchers_validate_the_mandatory_secret():
    for script in ("run.sh", "run_webapp.sh"):
        path = os.path.join(h.PROJECT_DIR, script)
        assert subprocess.run(["bash", "-n", path], check=False).returncode == 0
        with open(path, encoding="utf-8") as source:
            assert "NULIGAHELPER_SECRET:?" in source.read()
    secret = os.environ.pop("NULIGAHELPER_SECRET")
    try:
        try:
            main.main()
        except RuntimeError as exc:
            assert "NULIGAHELPER_SECRET" in str(exc)
        else:
            raise AssertionError("the daily job must refuse an absent secret")
    finally:
        os.environ["NULIGAHELPER_SECRET"] = secret


if __name__ == "__main__":
    h.run_all(dict(globals()))
