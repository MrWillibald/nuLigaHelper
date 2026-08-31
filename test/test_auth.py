"""Passwordless authentication and registration tests."""

import os
import re
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta

import helpers as h
import db
import notifier
import webapp


def _new_app():
    path = os.path.join(h._TEST_DIR, f"auth-{next(tempfile._get_candidate_names())}.db")
    previous = os.environ["NULIGAHELPER_DB"]
    os.environ["NULIGAHELPER_DB"] = path
    try:
        app = webapp.create_app()
    finally:
        os.environ["NULIGAHELPER_DB"] = previous
    app.config["TESTING"] = True
    return app, db.make_engine(path)


def _csrf(client, path="/login"):
    page = client.get(path).get_data(as_text=True)
    return re.search(r'name="csrf-token" content="([^"]+)"', page).group(1)


def _capture_messages():
    messages = []
    original = notifier.Notifier.send_account_message

    def fake(self, person, subject, mail_body, sms_body):
        messages.append((person.id, subject, mail_body, sms_body))
        return 1

    notifier.Notifier.send_account_message = fake
    return messages, original


def test_app_requires_secret_and_configures_sliding_hour():
    secret = os.environ.pop("NULIGAHELPER_SECRET")
    try:
        try:
            webapp.create_app()
        except RuntimeError as exc:
            assert "NULIGAHELPER_SECRET" in str(exc)
        else:
            raise AssertionError("startup without a secret must fail")
    finally:
        os.environ["NULIGAHELPER_SECRET"] = secret
    app, _ = _new_app()
    assert app.permanent_session_lifetime == timedelta(hours=1)
    assert app.config["SESSION_REFRESH_EACH_REQUEST"] is True


def test_email_token_is_single_use_and_expiry_is_checked():
    app, engine = _new_app()
    with h.Session(engine) as session:
        person = db.Person(name="Mail", email="mail@example.test")
        session.add(person)
        session.commit()
        person_id = person.id
    messages, original = _capture_messages()
    try:
        client = app.test_client()
        csrf = _csrf(client)
        sent = client.post("/login", data={
            "contact": "mail@example.test", "csrf_token": csrf
        })
        assert sent.status_code == 200 and len(messages) == 1
        link = re.search(r"http://localhost(/login/token/\S+)", messages[-1][2]).group(1)
        assert client.get(link).status_code == 302
        assert client.get(link).status_code == 400, "a consumed link must not replay"

        client = app.test_client()
        csrf = _csrf(client)
        client.post("/login", data={
            "contact": "mail@example.test", "csrf_token": csrf
        })
        expired_link = re.search(
            r"http://localhost(/login/token/\S+)", messages[-1][2]
        ).group(1)
        with h.Session(engine) as session:
            token = session.query(db.AuthToken).order_by(db.AuthToken.id.desc()).first()
            token.expires_at = datetime.now() - timedelta(seconds=1)
            session.commit()
        assert client.get(expired_link).status_code == 400
        with client.session_transaction() as browser_session:
            assert browser_session.get("person_id") is None
    finally:
        notifier.Notifier.send_account_message = original


def test_concurrent_token_consumption_has_one_winner():
    app, engine = _new_app()
    with h.Session(engine) as session:
        session.add(db.Person(name="Mail", email="race@example.test"))
        session.commit()
    messages, original = _capture_messages()
    try:
        requester = app.test_client()
        csrf = _csrf(requester)
        requester.post("/login", data={
            "contact": "race@example.test", "csrf_token": csrf
        })
        link = re.search(r"http://localhost(/login/token/\S+)", messages[-1][2]).group(1)

        def consume():
            return app.test_client().get(link).status_code

        with ThreadPoolExecutor(max_workers=2) as executor:
            statuses = sorted(executor.map(lambda _: consume(), range(2)))
        assert statuses == [302, 400]
    finally:
        notifier.Notifier.send_account_message = original


def test_sms_code_and_channel_preference():
    app, engine = _new_app()
    with h.Session(engine) as session:
        mail = db.Person(name="Mail", email="mail@example.test", phone="+49170001")
        phone = db.Person(name="Phone", phone="+49170002")
        contactless = db.Person(name="No Contact")
        session.add_all([mail, phone, contactless])
        session.commit()
        contactless_id = contactless.id
    messages, original = _capture_messages()
    try:
        mail_client = app.test_client()
        token = _csrf(mail_client)
        mail_client.post("/login", data={
            "contact": "mail@example.test", "csrf_token": token
        })
        assert messages[-1][2] and not messages[-1][3], "mail must be preferred"

        phone_client = app.test_client()
        token = _csrf(phone_client)
        phone_client.post("/login", data={
            "contact": "+49170002", "csrf_token": token
        })
        code = re.search(r"(\d{6})", messages[-1][3]).group(1)
        response = phone_client.post("/login/code", data={
            "contact": "+49170002", "code": code, "csrf_token": token
        })
        assert response.status_code == 302
        with phone_client.session_transaction() as browser_session:
            assert browser_session.get("person_id") is not None
        assert not any(message[0] == contactless_id for message in messages)
    finally:
        notifier.Notifier.send_account_message = original


def test_login_is_enumeration_safe_and_rate_limited():
    app, engine = _new_app()
    with h.Session(engine) as session:
        session.add(db.Person(name="Known", email="known@example.test"))
        session.commit()
    messages, original = _capture_messages()
    try:
        client = app.test_client()
        csrf = _csrf(client)
        known = client.post("/login", data={
            "contact": "known@example.test", "csrf_token": csrf
        }).data
        unknown = client.post("/login", data={
            "contact": "unknown@example.test", "csrf_token": csrf
        }).data
        assert known == unknown
        for _ in range(4):
            client.post("/login", data={
                "contact": "known@example.test", "csrf_token": csrf
            })
        assert len(messages) == 3, "mail requests must stop at the per-person cap"
    finally:
        notifier.Notifier.send_account_message = original


def test_registration_verification_pending_gate_and_mv_approval():
    app, engine = _new_app()
    with h.Session(engine) as session:
        team = db.get_or_create_team(session, "BL mD")
        mv = db.Person(name="MV", email="mv@example.test", team=team)
        session.add(mv)
        session.flush()
        team.mv_person_id = mv.id
        session.commit()
        team_id, mv_id = team.id, mv.id
    messages, original = _capture_messages()
    try:
        registrant = app.test_client()
        csrf = _csrf(registrant, "/registrieren")
        missing = registrant.post("/registrieren", data={
            "name": "New", "team_id": team_id, "email": "new@example.test",
            "csrf_token": csrf,
        }, follow_redirects=True)
        assert "Zustimmung" in missing.get_data(as_text=True)
        new_response = registrant.post("/registrieren", data={
            "name": "New", "team_id": team_id, "email": "new@example.test",
            "consent": "yes", "csrf_token": csrf,
        })
        known_response = registrant.post("/registrieren", data={
            "name": "Other", "team_id": team_id, "email": "new@example.test",
            "consent": "yes", "csrf_token": csrf,
        })
        assert new_response.data == known_response.data
        verify_link = re.search(
            r"http://localhost(/registrieren/verifizieren/\S+)", messages[0][2]
        ).group(1)
        assert registrant.get(verify_link).status_code == 200
        with h.Session(engine) as session:
            person = session.query(db.Person).filter_by(name="New").one()
            assert person.account_status == db.ACCOUNT_VERIFIED
            assert person not in db.get_all_persons(session)
            person_id = person.id
        assert messages[-1][0] == mv_id, "the requested team's MV must be notified"

        h.sign_in(registrant, person_id)
        assert registrant.get("/personen").location.endswith("/registrierung/status")
        assert 'data-role="' not in registrant.get("/").get_data(as_text=True)
        mv_client = app.test_client()
        mv_csrf = h.sign_in(mv_client, mv_id)
        approved = mv_client.post(
            f"/registrierungen/{person_id}/approve",
            data=h.csrf_data(token=mv_csrf),
        )
        assert approved.status_code == 302
        with h.Session(engine) as session:
            person = session.get(db.Person, person_id)
            assert person.account_status == db.ACCOUNT_ACTIVE
            assert person.team_id == team_id
    finally:
        notifier.Notifier.send_account_message = original


def test_registration_without_mv_notifies_admin_fallback():
    app, engine = _new_app()
    with h.Session(engine) as session:
        support = db.get_support_team(session)
        admin = db.Person(
            name="Admin", email="admin@example.test", team=support, is_admin=True
        )
        session.add(admin)
        session.commit()
        support_id, admin_id = support.id, admin.id
    messages, original = _capture_messages()
    try:
        client = app.test_client()
        csrf = _csrf(client, "/registrieren")
        client.post("/registrieren", data={
            "name": "Support New", "team_id": support_id,
            "email": "support-new@example.test", "consent": "yes",
            "csrf_token": csrf,
        })
        verify_link = re.search(
            r"http://localhost(/registrieren/verifizieren/\S+)", messages[0][2]
        ).group(1)
        client.get(verify_link)
        assert messages[-1][0] == admin_id
    finally:
        notifier.Notifier.send_account_message = original


def test_active_request_refreshes_session_cookie_and_missing_session_is_401():
    app, engine = _new_app()
    with h.Session(engine) as session:
        person = db.Person(name="Member", email="member@example.test")
        session.add(person)
        session.commit()
        person_id = person.id
    client = app.test_client()
    h.sign_in(client, person_id)
    response = client.get("/statistik")
    assert "Set-Cookie" in response.headers, "active use must renew the session cookie"
    guest = app.test_client()
    response = guest.post("/api/assignment/claim", json={})
    assert response.status_code == 401
    assert response.get_json()["code"] == "session_expired"

    app.permanent_session_lifetime = timedelta(seconds=1)
    expiring = app.test_client()
    token = h.sign_in(expiring, person_id)
    time.sleep(2.1)
    expired = expiring.post(
        "/api/assignment/claim", json={}, headers=h.csrf_headers(token)
    )
    assert expired.status_code == 401


def test_deactivated_login_matches_unknown_contact():
    app, engine = _new_app()
    with h.Session(engine) as session:
        session.add(db.Person(
            name="Inactive", email="inactive@example.test",
            account_status=db.ACCOUNT_INACTIVE,
        ))
        session.commit()
    messages, original = _capture_messages()
    try:
        client = app.test_client()
        csrf = _csrf(client)
        inactive = client.post("/login", data={
            "contact": "inactive@example.test", "csrf_token": csrf
        })
        unknown = client.post("/login", data={
            "contact": "unknown@example.test", "csrf_token": csrf
        })
        assert inactive.data == unknown.data and not messages
    finally:
        notifier.Notifier.send_account_message = original


if __name__ == "__main__":
    h.run_all(dict(globals()))
