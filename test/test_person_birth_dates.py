"""Required date collection, identity-safe correction, and private roster visibility."""

import contextlib
import io
from datetime import date

import helpers as h
import db
import manage_db
from test_auth import (
    _capture_messages, _challenge, _code, _csrf, _new_app, _restore_messages,
)


def _roster():
    app, engine = _new_app()
    with h.Session(engine) as session:
        h.sync_sample_games(session)
        teams = [team for team in db.get_all_teams(session) if not team.is_support]
        own_team, other_team = teams[:2]
        people = {
            "admin": db.Person(name="Birth Admin", is_admin=True, birth_date=date(1981, 2, 3)),
            "mv": db.Person(name="Birth MV", teams=[own_team], birth_date=date(1982, 3, 4)),
            "member": db.Person(name="Birth Member", email="birth-member@example.test",
                                teams=[own_team], birth_date=date(2000, 4, 5)),
            "other": db.Person(name="Birth Other", teams=[other_team], birth_date=date(1999, 6, 7)),
            "legacy": db.Person(name="Birth Legacy", email="birth-legacy@example.test"),
            "pending": db.Person(name="Birth Pending", account_status=db.ACCOUNT_VERIFIED),
        }
        session.add_all(people.values())
        session.flush()
        own_team.mv_person_id = people["mv"].id
        session.commit()
        ids = {key: person.id for key, person in people.items()}
        ids.update(own_team=own_team.id, other_team=other_team.id)
    return app, engine, ids


def _signed_client(app, ids, role):
    client = app.test_client()
    h.sign_in(client, ids[role])
    return client


def _edit(client, person_id, **values):
    return client.post(f"/personen/{person_id}/edit", data=h.csrf_data(values))


def _cli(path, *arguments):
    args = manage_db.build_parser().parse_args(["--db", str(path), *map(str, arguments)])
    output = io.StringIO()
    with contextlib.redirect_stdout(output):
        args.func(args)
    return output.getvalue()


def test_registration_requires_date_before_sending_and_preserves_signed_step():
    app, engine, ids = _roster()
    client = app.test_client()
    messages, originals = _capture_messages()
    try:
        payload = {
            "csrf_token": _csrf(client, "/registrieren"),
            "action": "request_code", "name": "Birth Registrant",
            "email": "birth-register@example.test", "channel": "email",
            "team_ids": [ids["own_team"]], "consent": "yes",
        }
        for invalid in (None, "", "2001-02-29", "9999-12-31"):
            submitted = dict(payload)
            if invalid is not None:
                submitted["birth_date"] = invalid
            result = client.post("/registrieren", data=submitted)
            assert 'id="birth_date-error"' in result.get_data(as_text=True)
            assert messages == [], "invalid dates must not send a registration code"
            with h.Session(engine) as session:
                assert session.query(db.Person).filter_by(email=payload["email"]).first() is None

        requested = client.post("/registrieren", data={**payload, "birth_date": "2004-02-29"})
        challenge = _challenge(requested)
        with h.Session(engine) as session:
            person = session.query(db.Person).filter_by(email=payload["email"]).one()
            registered_id = person.id
            assert person.birth_date == date(2004, 2, 29)
        retry = client.post("/registrieren", data={**payload, "birth_date": "2000-01-01"})
        assert _challenge(retry) != challenge, "a resubmission must create a new signed challenge"
        challenge = _challenge(retry)
        confirmed = client.post("/registrieren", data={
            "csrf_token": payload["csrf_token"], "action": "confirm_code",
            "challenge": challenge, "code": _code(messages[-1]),
            "birth_date": "1990-01-01", "name": "Tampered confirmation",
        })
        assert confirmed.status_code == 302
        with h.Session(engine) as session:
            person = session.get(db.Person, registered_id)
            assert person.birth_date == date(2004, 2, 29), "only the signed stored registration is confirmed"
            assert person.name == "Birth Registrant" and person.account_status == db.ACCOUNT_VERIFIED
        assert all("2004-02-29" not in message["body"] for message in messages)
    finally:
        _restore_messages(originals)


def test_registration_existing_contacts_do_not_change_birth_date_or_public_state():
    app, engine, ids = _roster()
    client = app.test_client()
    messages, originals = _capture_messages()
    try:
        csrf = _csrf(client, "/registrieren")
        payload = {"csrf_token": csrf, "action": "request_code", "name": "Attempted replacement",
                   "birth_date": "1980-01-01", "channel": "email", "consent": "yes",
                   "team_ids": [ids["own_team"]]}
        known = client.post("/registrieren", data={**payload, "email": "birth-member@example.test"})
        unused = client.post("/registrieren", data={**payload, "email": "unused-birth@example.test"})
        assert _challenge(known) and _challenge(unused)
        assert "Falls die Angaben verwendet werden können" in known.get_data(as_text=True)
        assert "Falls die Angaben verwendet werden können" in unused.get_data(as_text=True)
        with h.Session(engine) as session:
            person = session.get(db.Person, ids["member"])
            assert person.birth_date == date(2000, 4, 5) and person.name == "Birth Member"
    finally:
        _restore_messages(originals)


def test_admin_and_mv_creation_require_date_and_roll_back_all_fields():
    app, engine, ids = _roster()
    admin = _signed_client(app, ids, "admin")
    mv = _signed_client(app, ids, "mv")
    for client, role in ((admin, "Admin"), (mv, "MV")):
        payload = {"name": f"Created {role}", "team_id": ids["own_team"],
                   "team_ids": [ids["own_team"]], "email": "", "phone": ""}
        for invalid in (None, "2001-02-29", "9999-01-01"):
            submitted = dict(payload)
            if invalid is not None:
                submitted["birth_date"] = invalid
            assert client.post("/personen/add", data=h.csrf_data(submitted)).status_code == 302
            with h.Session(engine) as session:
                assert session.query(db.Person).filter_by(name=payload["name"]).first() is None
        assert client.post("/personen/add", data=h.csrf_data({**payload, "birth_date": "2007-08-09"})).status_code == 302
        with h.Session(engine) as session:
            person = session.query(db.Person).filter_by(name=payload["name"]).one()
            assert person.birth_date == date(2007, 8, 9) and person.email is None and person.phone is None
            assert db.membership_team_ids(person) == (ids["own_team"],)
            created_id = person.id
        if role == "MV":
            assert "2007-08-09" not in mv.get("/personen").get_data(as_text=True)
            assert _edit(mv, created_id, birth_date="1990-01-01").status_code == 403
    assert mv.post("/personen/add", data=h.csrf_data({
        "name": "Wrong team", "team_id": ids["other_team"], "birth_date": "2000-01-01",
    })).status_code == 403
    invalid_team = admin.post("/personen/add", data=h.csrf_data({
        "name": "Incomplete insertion", "team_ids": [999999], "birth_date": "2000-01-01",
    }))
    assert invalid_team.status_code == 302
    with h.Session(engine) as session:
        assert session.query(db.Person).filter_by(name="Incomplete insertion").first() is None


def test_birth_date_completion_correction_and_clearing_preserve_identity_and_other_data():
    app, engine, ids = _roster()
    member = _signed_client(app, ids, "member")
    legacy = _signed_client(app, ids, "legacy")
    admin = _signed_client(app, ids, "admin")
    assert _edit(legacy, ids["legacy"], name="Legacy contact repair", email="repaired@example.test", birth_date="").status_code == 302
    with h.Session(engine) as session:
        person = session.get(db.Person, ids["legacy"])
        assert person.birth_date is None and person.email == "repaired@example.test"
    assert _edit(legacy, ids["legacy"], name="Legacy completed", email="repaired@example.test", birth_date="1991-03-05").status_code == 302
    assert _edit(member, ids["member"], name="Corrected self", email="birth-member@example.test", birth_date="2001-04-05").status_code == 302
    for invalid in ("", "2001-02-29", "9999-12-31"):
        assert _edit(member, ids["member"], name="Must not save", email="must-not-save@example.test", birth_date=invalid).status_code == 302
        with h.Session(engine) as session:
            person = session.get(db.Person, ids["member"])
            assert person.birth_date == date(2001, 4, 5) and person.name == "Corrected self"
            assert person.email == "birth-member@example.test", "invalid correction must roll back unrelated fields"
    assert _edit(member, ids["member"], name="Contact edit only", email="birth-member@example.test").status_code == 302
    assert _edit(admin, ids["other"], name="Admin corrected", birth_date="1998-06-07").status_code == 302
    assert _edit(member, ids["other"], birth_date="1990-01-01").status_code == 403
    with h.Session(engine) as session:
        assert session.get(db.Person, ids["legacy"]).birth_date == date(1991, 3, 5)
        assert session.get(db.Person, ids["member"]).birth_date == date(2001, 4, 5)
        assert session.get(db.Person, ids["other"]).birth_date == date(1998, 6, 7)


def test_person_page_date_visibility_is_self_admin_only_and_missing_queue_is_admin_only():
    app, _, ids = _roster()
    for role, own_date in (("member", "2000-04-05"), ("mv", "1982-03-04")):
        client = _signed_client(app, ids, role)
        page = client.get("/personen").get_data(as_text=True)
        assert own_date in page
        assert "1999-06-07" not in page and "1981-02-03" not in page
        assert 'name="birth_date" value="missing"' not in page
        assert "Birth Pending" not in client.get("/personen?birth_date=missing").get_data(as_text=True)
    admin = _signed_client(app, ids, "admin")
    page = admin.get("/personen").get_data(as_text=True)
    assert all(value in page for value in ("2000-04-05", "1982-03-04", "1999-06-07"))
    queue = admin.get("/personen?birth_date=missing").get_data(as_text=True)
    roster = queue.split('<div class="people-grid persons-grid">', 1)[1]
    assert "Birth Legacy" in roster and "Birth Pending" in roster
    assert "Birth Member" not in roster and "1999-06-07" not in roster
    guest = app.test_client().get("/personen")
    assert guest.status_code == 302 and "/login" in guest.location


def test_cli_requires_valid_date_and_corrects_duplicate_name_by_id_without_printing_dates():
    engine = h.make_engine()
    path = engine.url.database
    with contextlib.redirect_stderr(io.StringIO()):
        try:
            _cli(path, "add-person", "Missing date")
        except SystemExit as exc:
            assert exc.code == 2
        else:
            raise AssertionError("CLI creation must require --birth-date")
    for invalid in ("2001-02-29", "9999-12-31"):
        try:
            _cli(path, "add-person", "Invalid date", "--birth-date", invalid)
        except SystemExit as exc:
            assert invalid not in str(exc), "validation errors should not log the raw date"
        else:
            raise AssertionError("CLI creation must reject invalid dates")
    with h.Session(engine) as session:
        assert session.query(db.Person).count() == 0, "invalid CLI input cannot leave partial records"
    for value in ("1985-02-03", "1990-05-06"):
        output = _cli(path, "add-person", "Same Birth Name", "--birth-date", value)
        assert value not in output
    with h.Session(engine) as session:
        first, second = session.query(db.Person).order_by(db.Person.id).all()
        first_id, second_id = first.id, second.id
    output = _cli(path, "set-birth-date", second_id, "1992-07-08")
    assert f"ID {second_id}" in output and "1992-07-08" not in output
    with h.Session(engine) as session:
        assert session.get(db.Person, first_id).birth_date == date(1985, 2, 3)
        assert session.get(db.Person, second_id).birth_date == date(1992, 7, 8)
    for person_id, value in ((999999, "1990-01-01"), (second_id, ""), (second_id, "2001-02-29")):
        try:
            _cli(path, "set-birth-date", person_id, value)
        except SystemExit:
            pass
        else:
            raise AssertionError("CLI corrections must select an existing ID and valid date")
    with h.Session(engine) as session:
        assert session.get(db.Person, second_id).birth_date == date(1992, 7, 8)


if __name__ == "__main__":
    h.run_all(dict(globals()))
