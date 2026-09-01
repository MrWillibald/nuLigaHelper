"""Passwordless authentication, registration and contact-route tests."""

import os
import re
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta

from itsdangerous import URLSafeTimedSerializer

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


def _challenge(response):
    match = re.search(
        r'name="challenge" value="([^"]+)"',
        response.get_data(as_text=True),
    )
    assert match, "the confirmation state must contain an opaque challenge"
    return match.group(1)


def _code(message):
    match = re.search(r"\b(\d{6})\b", message["body"])
    assert match, "authentication messages must contain a six-digit code"
    return match.group(1)


def _capture_messages():
    messages = []
    originals = (
        notifier.Notifier.send_account_message,
        notifier.Notifier.send_account_message_via,
    )

    def fake_fallback(self, person, subject, mail_body, sms_body):
        messages.append({
            "person_id": person.id,
            "channel": "fallback",
            "subject": subject,
            "body": mail_body or sms_body,
        })
        return 1

    def fake_via(self, person, channel, subject, body):
        messages.append({
            "person_id": person.id,
            "channel": channel,
            "subject": subject,
            "body": body,
        })
        return 1

    notifier.Notifier.send_account_message = fake_fallback
    notifier.Notifier.send_account_message_via = fake_via
    return messages, originals


def _restore_messages(originals):
    (
        notifier.Notifier.send_account_message,
        notifier.Notifier.send_account_message_via,
    ) = originals


def _request_login(client, csrf, *, channel="email", contact="mail@example.test"):
    data = {
        "action": "request_code",
        "channel": channel,
        "csrf_token": csrf,
    }
    if channel == "email":
        data["email"] = contact
    else:
        data.update({
            "country_code": "+49",
            "phone": contact,
        })
    return client.post("/login", data=data)


def _confirm_login(client, csrf, challenge, code):
    return client.post("/login", data={
        "action": "confirm_code",
        "challenge": challenge,
        "code": code,
        "csrf_token": csrf,
    })


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


def test_email_code_is_purpose_bound_single_use_and_expiring():
    app, engine = _new_app()
    with h.Session(engine) as session:
        person = db.Person(name="Mail", email="mail@example.test")
        session.add(person)
        session.commit()
        person_id = person.id
    messages, originals = _capture_messages()
    try:
        client = app.test_client()
        csrf = _csrf(client)
        requested = _request_login(
            client, csrf, contact="  MAIL@Example.Test "
        )
        challenge = _challenge(requested)
        code = _code(messages[-1])
        assert messages[-1]["channel"] == "email"
        assert "mail@example.test" not in requested.get_data(as_text=True)
        with h.Session(engine) as session:
            assert session.query(db.AuthToken).one().code == code

        wrong_purpose = client.post("/registrieren", data={
            "action": "confirm_code",
            "challenge": challenge,
            "code": code,
            "csrf_token": csrf,
        })
        assert wrong_purpose.status_code == 200
        assert _confirm_login(client, csrf, challenge, "000000").status_code == 200
        success = _confirm_login(client, csrf, challenge, code)
        assert success.status_code == 302
        assert success.location.endswith("/")
        with client.session_transaction() as browser_session:
            assert browser_session["person_id"] == person_id

        replay = app.test_client()
        replay_csrf = _csrf(replay)
        refused = _confirm_login(replay, replay_csrf, challenge, code)
        assert refused.status_code == 200
        with replay.session_transaction() as browser_session:
            assert browser_session.get("person_id") is None

        expiring = app.test_client()
        expiring_csrf = _csrf(expiring)
        expiring_response = _request_login(
            expiring, expiring_csrf, contact="mail@example.test"
        )
        expiring_challenge = _challenge(expiring_response)
        expiring_code = _code(messages[-1])
        with h.Session(engine) as session:
            token = session.query(db.AuthToken).order_by(db.AuthToken.id.desc()).first()
            token.expires_at = datetime.now() - timedelta(seconds=1)
            session.commit()
        refused = _confirm_login(
            expiring, expiring_csrf, expiring_challenge, expiring_code
        )
        assert refused.status_code == 200
        with expiring.session_transaction() as browser_session:
            assert browser_session.get("person_id") is None
    finally:
        _restore_messages(originals)


def test_tampered_challenge_and_concurrent_consumption_have_no_extra_winner():
    app, engine = _new_app()
    with h.Session(engine) as session:
        session.add(db.Person(name="Race", email="race@example.test"))
        session.commit()
    messages, originals = _capture_messages()
    try:
        requester = app.test_client()
        csrf = _csrf(requester)
        response = _request_login(
            requester, csrf, contact="race@example.test"
        )
        challenge, code = _challenge(response), _code(messages[-1])
        tampered = _confirm_login(requester, csrf, challenge + "x", code)
        assert tampered.status_code == 200
        with requester.session_transaction() as browser_session:
            assert browser_session.get("person_id") is None

        def consume():
            client = app.test_client()
            token = _csrf(client)
            response = _confirm_login(client, token, challenge, code)
            with client.session_transaction() as browser_session:
                signed_in = browser_session.get("person_id") is not None
            return response.status_code, signed_in

        with ThreadPoolExecutor(max_workers=2) as executor:
            outcomes = list(executor.map(lambda _: consume(), range(2)))
        assert sum(signed_in for _, signed_in in outcomes) == 1
        assert sorted(status for status, _ in outcomes) == [200, 302]
    finally:
        _restore_messages(originals)


def test_selected_channel_controls_delivery_for_people_with_both_or_one():
    app, engine = _new_app()
    phone = "+491701234567"
    with h.Session(engine) as session:
        both = db.Person(
            name="Both", email="both@example.test", phone=phone
        )
        mail_only = db.Person(name="Mail", email="mail@example.test")
        contactless = db.Person(name="None")
        session.add_all([both, mail_only, contactless])
        session.commit()
        both_id, none_id = both.id, contactless.id
    messages, originals = _capture_messages()
    try:
        sms_client = app.test_client()
        csrf = _csrf(sms_client)
        sms_response = _request_login(
            sms_client, csrf, channel="sms", contact="0170 1234567"
        )
        sms_challenge = _challenge(sms_response)
        assert messages[-1]["person_id"] == both_id
        assert messages[-1]["channel"] == "sms"
        sms_login = _confirm_login(
            sms_client, csrf, sms_challenge, _code(messages[-1])
        )
        assert sms_login.status_code == 302
        with sms_client.session_transaction() as browser_session:
            assert browser_session["person_id"] == both_id

        mail_client = app.test_client()
        csrf = _csrf(mail_client)
        mail_response = _request_login(
            mail_client, csrf, contact="BOTH@example.test"
        )
        mail_challenge = _challenge(mail_response)
        assert messages[-1]["person_id"] == both_id
        assert messages[-1]["channel"] == "email"
        mail_login = _confirm_login(
            mail_client, csrf, mail_challenge, _code(messages[-1])
        )
        assert mail_login.status_code == 302
        with mail_client.session_transaction() as browser_session:
            assert browser_session["person_id"] == both_id

        unknown_client = app.test_client()
        csrf = _csrf(unknown_client)
        _request_login(
            unknown_client, csrf, contact="nobody@example.test"
        )
        assert not any(message["person_id"] == none_id for message in messages)
    finally:
        _restore_messages(originals)


def test_unknown_and_ineligible_login_are_same_shaped_dummy_challenges():
    app, engine = _new_app()
    with h.Session(engine) as session:
        session.add_all([
            db.Person(name="Known", email="known@example.test"),
            db.Person(
                name="Inactive",
                email="inactive@example.test",
                account_status=db.ACCOUNT_INACTIVE,
            ),
        ])
        session.commit()
    messages, originals = _capture_messages()
    try:
        client = app.test_client()
        csrf = _csrf(client)
        known = _request_login(
            client, csrf, contact="known@example.test"
        )
        unknown = _request_login(
            client, csrf, contact="unknown@example.test"
        )
        inactive = _request_login(
            client, csrf, contact="inactive@example.test"
        )
        assert len(messages) == 1
        for response in (known, unknown, inactive):
            body = response.get_data(as_text=True)
            assert "Falls die Angaben bekannt sind" in body
            assert 'name="challenge"' in body
            assert "Known" not in body and "Inactive" not in body
        assert "unknown@example.test" not in unknown.get_data(as_text=True)
        dummy = _challenge(unknown)
        refused = _confirm_login(client, csrf, dummy, "123456")
        assert refused.status_code == 200
        with client.session_transaction() as browser_session:
            assert browser_session.get("person_id") is None
    finally:
        _restore_messages(originals)


def test_login_validation_canonical_lookup_and_rate_limits():
    app, engine = _new_app()
    with h.Session(engine) as session:
        session.add_all([
            db.Person(name="Mail", email="mail@example.test"),
            db.Person(name="SMS", phone="+491701234567"),
        ])
        session.commit()
    messages, originals = _capture_messages()
    try:
        invalid = app.test_client()
        csrf = _csrf(invalid)
        bad_mail = _request_login(invalid, csrf, contact="kein-kontakt")
        bad_phone = _request_login(
            invalid, csrf, channel="sms", contact="123"
        )
        assert "gültige E-Mail-Adresse" in bad_mail.get_data(as_text=True)
        assert "gültige Telefonnummer" in bad_phone.get_data(as_text=True)
        assert not messages

        mail_client = app.test_client()
        csrf = _csrf(mail_client)
        for _ in range(5):
            _request_login(
                mail_client, csrf, contact=" MAIL@Example.Test "
            )
        assert [m["channel"] for m in messages].count("email") == 3

        sms_client = app.test_client()
        csrf = _csrf(sms_client)
        for _ in range(4):
            _request_login(
                sms_client, csrf, channel="sms", contact="0170-1234567"
            )
        assert [m["channel"] for m in messages].count("sms") == 2
    finally:
        _restore_messages(originals)


def test_registration_code_replacement_verification_and_mv_approval():
    app, engine = _new_app()
    with h.Session(engine) as session:
        team = db.get_or_create_team(session, "BL mD")
        mv = db.Person(name="MV", email="mv@example.test", team=team)
        session.add(mv)
        session.flush()
        team.mv_person_id = mv.id
        session.commit()
        team_id, mv_id = team.id, mv.id
    messages, originals = _capture_messages()
    try:
        client = app.test_client()
        csrf = _csrf(client, "/registrieren")
        base = {
            "action": "request_code",
            "name": "New Helper",
            "team_id": team_id,
            "channel": "email",
            "email": " New.Helper@Example.Test ",
            "consent": "yes",
            "csrf_token": csrf,
        }
        first = client.post("/registrieren", data=base)
        first_challenge, first_code = _challenge(first), _code(messages[-1])
        second = client.post("/registrieren", data={
            **base,
            "name": "Ignored Replacement",
            "email": "new.helper@example.test",
        })
        second_challenge, second_code = _challenge(second), _code(messages[-1])
        with h.Session(engine) as session:
            people = session.query(db.Person).filter_by(
                email="new.helper@example.test"
            ).all()
            assert len(people) == 1
            person_id = people[0].id
            assert people[0].name == "New Helper"

        assert client.post("/registrieren", data={
            "action": "confirm_code",
            "challenge": first_challenge,
            "code": first_code,
            "csrf_token": csrf,
        }).status_code == 200
        completed = client.post("/registrieren", data={
            "action": "confirm_code",
            "challenge": second_challenge,
            "code": second_code,
            "csrf_token": csrf,
        })
        assert completed.status_code == 302
        assert completed.location.endswith("/registrierung/status")
        with h.Session(engine) as session:
            person = session.get(db.Person, person_id)
            assert person.account_status == db.ACCOUNT_VERIFIED
            assert person not in db.get_all_persons(session)
        assert any(
            message["channel"] == "fallback"
            and message["person_id"] == mv_id
            for message in messages
        )
        assert client.get("/personen").location.endswith("/registrierung/status")
    finally:
        _restore_messages(originals)


def test_registration_rejects_invalid_or_unconsented_writes_and_duplicate_accounts():
    app, engine = _new_app()
    with h.Session(engine) as session:
        team = db.get_or_create_team(session, "BL mD")
        existing = db.Person(name="Existing", email="existing@example.test", team=team)
        session.add(existing)
        session.commit()
        team_id, existing_id = team.id, existing.id
    messages, originals = _capture_messages()
    try:
        client = app.test_client()
        csrf = _csrf(client, "/registrieren")
        common_data = {
            "action": "request_code",
            "name": "New",
            "team_id": team_id,
            "channel": "email",
            "csrf_token": csrf,
        }
        missing = client.post("/registrieren", data={
            **common_data, "email": "new@example.test"
        })
        invalid = client.post("/registrieren", data={
            **common_data, "email": "invalid", "consent": "yes"
        })
        assert "Zustimmung" in missing.get_data(as_text=True)
        assert re.search(
            r'id="consent"[^>]*aria-invalid="true"',
            missing.get_data(as_text=True),
        )
        assert "gültige E-Mail-Adresse" in invalid.get_data(as_text=True)
        with h.Session(engine) as session:
            assert session.query(db.Person).count() == 1

        duplicate = client.post("/registrieren", data={
            **common_data,
            "email": " EXISTING@Example.Test ",
            "consent": "yes",
        })
        assert _challenge(duplicate)
        assert messages[-1]["person_id"] == existing_id
        assert "bereits ein Konto" in messages[-1]["body"]
        with h.Session(engine) as session:
            assert session.query(db.Person).count() == 1
    finally:
        _restore_messages(originals)


def test_registration_validates_and_stores_every_supplied_contact():
    app, engine = _new_app()
    with h.Session(engine) as session:
        team = db.get_or_create_team(session, "BL mD")
        session.commit()
        team_id = team.id
    messages, originals = _capture_messages()
    try:
        client = app.test_client()
        csrf = _csrf(client, "/registrieren")
        base = {
            "action": "request_code",
            "team_id": team_id,
            "consent": "yes",
            "country_code": "+49",
            "csrf_token": csrf,
        }

        empty = client.post("/registrieren", data={
            **base, "name": "Empty", "channel": "email",
        })
        invalid_phone = client.post("/registrieren", data={
            **base,
            "name": "Invalid Phone",
            "channel": "email",
            "email": "valid@example.test",
            "phone": "123",
        })
        invalid_email = client.post("/registrieren", data={
            **base,
            "name": "Invalid Mail",
            "channel": "sms",
            "email": "invalid",
            "phone": "0170 1234567",
        })
        mismatched_prefix = client.post("/registrieren", data={
            **base,
            "name": "Wrong Prefix",
            "channel": "sms",
            "phone": "+43 664 1234567",
        })
        assert "E-Mail-Adresse oder Mobilnummer" in empty.get_data(as_text=True)
        assert "gültige Telefonnummer" in invalid_phone.get_data(as_text=True)
        assert "gültige E-Mail-Adresse" in invalid_email.get_data(as_text=True)
        assert "passt nicht zur gewählten Ländervorwahl" in (
            mismatched_prefix.get_data(as_text=True)
        )
        invalid_phone_body = invalid_phone.get_data(as_text=True)
        assert 'value="valid@example.test"' in invalid_phone_body
        assert 'value="123"' in invalid_phone_body
        assert re.search(r'id="phone"[^>]*aria-invalid="true"', invalid_phone_body)
        assert not messages
        with h.Session(engine) as session:
            assert session.query(db.Person).count() == 0

        email_only = client.post("/registrieren", data={
            **base,
            "name": "Mail Only",
            "channel": "email",
            "email": " Mail.Only@Example.Test ",
        })
        phone_only = client.post("/registrieren", data={
            **base,
            "name": "Phone Only",
            "channel": "sms",
            "phone": "0170 2345678",
        })
        both = client.post("/registrieren", data={
            **base,
            "name": "Both",
            "channel": "email",
            "email": " Both@Example.Test ",
            "phone": "0170 3456789",
        })
        assert all(_challenge(response) for response in (email_only, phone_only, both))
        assert [message["channel"] for message in messages] == [
            "email", "sms", "email",
        ]
        with h.Session(engine) as session:
            people = {
                person.name: person for person in session.query(db.Person).all()
            }
            assert people["Mail Only"].email == "mail.only@example.test"
            assert people["Mail Only"].phone is None
            assert people["Phone Only"].email is None
            assert people["Phone Only"].phone == "+491702345678"
            assert people["Both"].email == "both@example.test"
            assert people["Both"].phone == "+491703456789"
    finally:
        _restore_messages(originals)


def test_registration_contact_conflicts_are_atomic_and_generic():
    app, engine = _new_app()
    with h.Session(engine) as session:
        team = db.get_or_create_team(session, "BL mD")
        email_owner = db.Person(
            name="Mail Owner", email="owner@example.test", team=team
        )
        phone_owner = db.Person(
            name="Phone Owner", phone="+491701234567", team=team
        )
        session.add_all([email_owner, phone_owner])
        session.commit()
        team_id = team.id
        email_owner_id = email_owner.id
    messages, originals = _capture_messages()
    try:
        client = app.test_client()
        csrf = _csrf(client, "/registrieren")
        base = {
            "action": "request_code",
            "name": "Must Not Exist",
            "team_id": team_id,
            "consent": "yes",
            "country_code": "+49",
            "csrf_token": csrf,
        }
        responses = [
            client.post("/registrieren", data={
                **base,
                "channel": "sms",
                "email": "owner@example.test",
                "phone": "0170 9999999",
            }),
            client.post("/registrieren", data={
                **base,
                "channel": "email",
                "email": "owner@example.test",
                "phone": "0170 1234567",
            }),
            client.post("/registrieren", data={
                **base,
                "channel": "email",
                "email": "owner@example.test",
                "phone": "0170 8888888",
            }),
        ]
        for response in responses:
            body = response.get_data(as_text=True)
            assert _challenge(response)
            assert "Falls die Angaben verwendet werden können" in body
            assert "Mail Owner" not in body and "Phone Owner" not in body
        assert len(messages) == 1
        assert messages[0]["person_id"] == email_owner_id
        assert messages[0]["channel"] == "email"
        assert "bereits ein Konto" in messages[0]["body"]
        with h.Session(engine) as session:
            assert session.query(db.Person).count() == 2
            email_owner = session.get(db.Person, email_owner_id)
            assert email_owner.email == "owner@example.test"
            assert email_owner.phone is None
    finally:
        _restore_messages(originals)


def test_active_and_verified_login_redirects_are_explicit():
    app, engine = _new_app()
    with h.Session(engine) as session:
        active = db.Person(name="Active", email="active@example.test")
        verified = db.Person(
            name="Pending",
            email="pending@example.test",
            account_status=db.ACCOUNT_VERIFIED,
        )
        session.add_all([active, verified])
        session.commit()
    messages, originals = _capture_messages()
    try:
        for address, target in (
            ("active@example.test", "/"),
            ("pending@example.test", "/registrierung/status"),
        ):
            client = app.test_client()
            csrf = _csrf(client)
            response = _request_login(client, csrf, contact=address)
            confirmed = _confirm_login(
                client, csrf, _challenge(response), _code(messages[-1])
            )
            assert confirmed.status_code == 302
            assert confirmed.location.endswith(target)
    finally:
        _restore_messages(originals)


def test_legacy_code_urls_redirect_and_old_email_links_still_consume():
    app, engine = _new_app()
    with h.Session(engine) as session:
        login_person = db.Person(name="Legacy Login", email="legacy@example.test")
        team = db.get_or_create_team(session, "BL mD")
        verify_person = db.register_person(
            session, "Legacy Register", team, email="register@example.test"
        )
        session.add_all([login_person, verify_person])
        session.flush()
        now = datetime.now()
        login_nonce = "legacy-login-nonce"
        verify_nonce = "legacy-verify-nonce"
        session.add_all([
            db.AuthToken(
                nonce=login_nonce,
                code=None,
                purpose="login",
                person=login_person,
                issued_at=now,
                expires_at=now + timedelta(minutes=15),
            ),
            db.AuthToken(
                nonce=verify_nonce,
                code=None,
                purpose="verify",
                person=verify_person,
                issued_at=now,
                expires_at=now + timedelta(minutes=15),
            ),
        ])
        session.commit()
        verify_person_id = verify_person.id
    serializer = URLSafeTimedSerializer(
        os.environ["NULIGAHELPER_SECRET"], salt="nuligahelper-auth"
    )
    login_token = serializer.dumps({
        "nonce": login_nonce,
        "purpose": "login",
    })
    verify_token = serializer.dumps({
        "nonce": verify_nonce,
        "purpose": "verify",
    })

    client = app.test_client()
    assert client.get("/login/code").location.endswith("/login")
    assert client.get("/registrieren/code").location.endswith("/registrieren")
    assert client.get(f"/login/token/{login_token}").status_code == 302
    assert client.get(f"/login/token/{login_token}").status_code == 400

    registrant = app.test_client()
    response = registrant.get(
        f"/registrieren/verifizieren/{verify_token}"
    )
    assert response.status_code == 200
    with h.Session(engine) as session:
        assert (
            session.get(db.Person, verify_person_id).account_status
            == db.ACCOUNT_VERIFIED
        )


def test_auth_markup_is_labelled_ordered_and_works_without_javascript():
    app, _ = _new_app()
    client = app.test_client()
    login = client.get("/login").get_data(as_text=True)
    register = client.get("/registrieren").get_data(as_text=True)

    for page in (login, register):
        assert '<fieldset class="auth-contact"' in page
        assert "<legend>Kontaktweg</legend>" in page
        assert 'for="email"' in page and 'id="email-help"' in page
        assert 'for="phone"' in page and 'id="phone-help"' in page
        assert 'for="country-code">Ländervorwahl</label>' in page
        assert page.index('id="email"') < page.index('id="phone"')
        assert page.index('id="phone"') < page.index("<legend>Kontaktweg")
        assert 'name="action" value="request_code"' in page
        assert 'name="action" value="confirm_code"' in page
        assert "auth-button-secondary" in page
        assert "auth-button-primary" in page
        assert 'name="code"' in page and "one-time-code" in page
    assert login.index("CODE ANFORDERN") < login.index("ANMELDEN")
    assert register.index('id="name"') < register.index('id="team"')
    assert register.index('id="team"') < register.index('id="consent"')
    assert register.index('id="consent"') < register.index('id="email"')
    assert register.index("CODE ANFORDERN") < register.index(
        "REGISTRIERUNG ABSCHLIESSEN"
    )
    assert "SMS-Code eingeben" not in login
    assert "<script" in login, "JavaScript may enhance but must not own submission"


def test_auth_progressive_enhancement_and_responsive_css_contract():
    with open(h.PROJECT_DIR + "/static/app.js", encoding="utf-8") as source:
        javascript = source.read()
    assert "function updateRouteAvailability()" in javascript
    assert "emailInput.checkValidity()" in javascript
    assert "input.disabled = locked || !validRoutes[input.value]" in javascript
    assert "invalidSupplied || !selected" in javascript
    assert "addEventListener(\"input\", updateRouteAvailability)" in javascript
    assert "countrySelect.addEventListener(\"change\", updateRouteAvailability)" in javascript
    assert "panel.hidden" not in javascript

    with open(h.PROJECT_DIR + "/static/style.css", encoding="utf-8") as source:
        css = source.read()
    assert ".auth-card{width:min(100%,620px)" in css
    assert ".auth-radio:has(input:disabled)" in css
    assert "@media(max-width:700px)" in css
    assert ".auth-radio-group,.auth-phone-group{grid-template-columns:1fr}" in css


def test_person_contact_writes_validate_atomically_and_use_canonical_uniqueness():
    app, engine = _new_app()
    with h.Session(engine) as session:
        support = db.get_support_team(session)
        admin = db.Person(
            name="Admin",
            email="admin@example.test",
            team=support,
            is_admin=True,
        )
        member = db.Person(
            name="Member",
            email="member@example.test",
            team=support,
        )
        session.add_all([admin, member])
        session.commit()
        admin_id, member_id, support_id = admin.id, member.id, support.id

    admin_client = app.test_client()
    csrf = h.sign_in(admin_client, admin_id)
    invalid = admin_client.post(f"/personen/{member_id}/edit", data={
        "name": "Changed",
        "team_id": support_id,
        "email": "invalid",
        "phone": "",
        "csrf_token": csrf,
    })
    assert invalid.status_code == 302
    with h.Session(engine) as session:
        member = session.get(db.Person, member_id)
        assert member.name == "Member"
        assert member.email == "member@example.test"

    admin_client.post("/personen/add", data={
        "name": "Canonical",
        "team_id": support_id,
        "email": " Canonical@Example.Test ",
        "phone": "",
        "csrf_token": csrf,
    })
    admin_client.post("/personen/add", data={
        "name": "Duplicate",
        "team_id": support_id,
        "email": "canonical@example.test",
        "phone": "",
        "csrf_token": csrf,
    })
    with h.Session(engine) as session:
        matches = session.query(db.Person).filter_by(
            email="canonical@example.test"
        ).all()
        assert len(matches) == 1

    member_client = app.test_client()
    member_csrf = h.sign_in(member_client, member_id)
    member_client.post(f"/personen/{member_id}/edit", data={
        "name": "Member",
        "email": " MEMBER@Example.Test ",
        "phone": "",
        "csrf_token": member_csrf,
    })
    with h.Session(engine) as session:
        assert session.get(db.Person, member_id).email == "member@example.test"


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
    assert "Set-Cookie" in response.headers
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


if __name__ == "__main__":
    h.run_all(dict(globals()))
