"""Shared message-source regressions for account and authentication workflows."""

from dataclasses import replace
import io
import logging
from unittest.mock import patch

import helpers as h
import auth_abuse
import db
import messages
from test_auth import (
    _capture_messages,
    _challenge,
    _code,
    _csrf,
    _new_app,
    _request_login,
    _restore_messages,
)


def _marked_account_templates():
    """Give each source/channel distinct wording without replacing its contract."""
    templates = {}
    for key in (
        "auth.registration_code", "auth.login_code", "auth.existing_account",
        "account.approval_request", "account.welcome",
    ):
        template = messages.CATALOG[key]
        templates[key] = replace(
            template,
            subject=f"Synthetic {key}: {template.subject}",
            email=f"EMAIL {key}\n{template.email}",
            sms=f"SMS {key}\n{template.sms or template.email}",
            shared_sms=False,
        )
    return templates


def test_account_workflows_use_shared_source_and_selected_channel_variants():
    app, engine = _new_app()
    with h.Session(engine) as session:
        first_team = db.get_or_create_team(session, "BL mD")
        second_team = db.get_or_create_team(session, "BL mC")
        email_admin = db.Person(
            name="Mail Admin {literal}", email="mail-admin@example.test",
            phone="+491701111111", is_admin=True,
        )
        sms_admin = db.Person(name="SMS Admin", phone="+491702222222", is_admin=True)
        dual_user = db.Person(
            name="Dual User {literal}", email="dual@example.test",
            phone="+491705551111",
        )
        sms_registrant = db.Person(
            name="SMS Registrant", phone="+491704444444",
            account_status=db.ACCOUNT_VERIFIED, teams=[first_team],
        )
        session.add_all([email_admin, sms_admin, dual_user, sms_registrant])
        session.commit()
        team_ids = [first_team.id, second_team.id]
        admin_id, sms_admin_id = email_admin.id, sms_admin.id
        dual_id, sms_registrant_id = dual_user.id, sms_registrant.id

    sent, originals = _capture_messages()
    try:
        with patch.dict(messages.CATALOG, _marked_account_templates()):
            for channel, contact in (
                ("sms", "0170-5551111"), ("email", "dual@example.test"),
            ):
                client = app.test_client()
                response = _request_login(client, _csrf(client), channel=channel, contact=contact)
                assert _challenge(response)
                assert sent[-1]["person_id"] == dual_id
                assert sent[-1]["channel"] == channel
                assert sent[-1]["subject"].startswith("Synthetic auth.login_code:")
                assert sent[-1]["body"].startswith(f"{channel.upper()} auth.login_code\n")
                assert "Dual User {literal}" in sent[-1]["body"]
                assert "15 Minuten" in sent[-1]["body"] and _code(sent[-1])

            unknown = app.test_client()
            count = len(sent)
            response = _request_login(unknown, _csrf(unknown), contact="unknown@example.test")
            assert _challenge(response) and len(sent) == count
            assert "Falls die Angaben bekannt sind" in response.get_data(as_text=True)

            registrant = app.test_client()
            csrf = _csrf(registrant, "/registrieren")
            registration_data = {
                "action": "request_code", "name": "New Helper {literal}",
                "birth_date": "1990-01-01", "team_ids": team_ids,
                "channel": "sms", "email": "new-helper@example.test",
                "phone": "0170-8883333", "country_code": "+49",
                "consent": "yes", "csrf_token": csrf,
            }
            response = registrant.post("/registrieren", data=registration_data)
            assert sent[-1]["channel"] == "sms", "registration must use the selected route"
            assert sent[-1]["body"].startswith("SMS auth.registration_code\n")
            code = _code(sent[-1])
            challenge = _challenge(response)
            registrant_id = sent[-1]["person_id"]
            before_verification = len(sent)
            assert registrant.post("/registrieren", data={
                "action": "confirm_code", "challenge": challenge,
                "code": code, "csrf_token": csrf,
            }).status_code == 302
            approval_requests = sent[before_verification:]
            assert [(item["person_id"], item["channel"]) for item in approval_requests] == [
                (admin_id, "email"), (sms_admin_id, "sms"),
            ]
            for item in approval_requests:
                assert item["subject"].startswith("Synthetic account.approval_request:")
                assert item["body"].startswith(f"{item['channel'].upper()} account.approval_request\n")
                assert "New Helper {literal}" in item["body"]
                assert "BL mD" in item["body"] and "BL mC" in item["body"]
                assert "Helfer verwalten" in item["body"]

            admin = app.test_client()
            admin_csrf = h.sign_in(admin, admin_id)
            for person_id, expected_channel in ((registrant_id, "email"), (sms_registrant_id, "sms")):
                assert admin.post(f"/registrierungen/{person_id}/approve", data=h.csrf_data(token=admin_csrf)).status_code == 302
                welcome = sent[-1]
                assert welcome["channel"] == expected_channel
                assert welcome["subject"].startswith("Synthetic account.welcome:")
                assert welcome["body"].startswith(f"{expected_channel.upper()} account.welcome\n")
                assert "herzlich Willkommen beim nuLigaHelper des TuS Raubling Handball!" in welcome["body"]
                assert "Registrierung wurde freigegeben" in welcome["body"]
                assert "Heimspielplan" in welcome["body"]

            duplicate = app.test_client()
            duplicate_csrf = _csrf(duplicate, "/registrieren")
            response = duplicate.post("/registrieren", data={
                **registration_data, "email": "dual@example.test",
                "phone": "0170-5551111", "csrf_token": duplicate_csrf,
            })
            assert _challenge(response), "existing contacts retain an opaque confirmation state"
            assert "Falls die Angaben verwendet werden können" in response.get_data(as_text=True)
            assert sent[-1]["person_id"] == dual_id and sent[-1]["channel"] == "sms"
            assert sent[-1]["body"].startswith("SMS auth.existing_account\n")
            assert "bereits ein Konto" in sent[-1]["body"]
    finally:
        _restore_messages(originals)


def test_account_rendering_failure_is_redacted_best_effort_after_commit():
    app, engine = _new_app()
    with h.Session(engine) as session:
        team = db.get_support_team(session)
        admin = db.Person(name="Admin", email="admin@example.test", is_admin=True)
        person = db.Person(
            name="private-name-canary", email="private-contact@example.test",
            account_status=db.ACCOUNT_VERIFIED, teams=[team],
        )
        session.add_all([admin, person])
        session.commit()
        admin_id, person_id = admin.id, person.id
    invalid_templates = {
        key: replace(messages.CATALOG[key], subject="private-template-canary\nInvalid header")
        for key in ("account.welcome", "auth.login_code")
    }
    sent, originals = _capture_messages()
    log = io.StringIO()
    handler = logging.StreamHandler(log)
    auth_abuse.LOGGER.addHandler(handler)
    try:
        with patch.dict(messages.CATALOG, invalid_templates):
            admin_client = app.test_client()
            csrf = h.sign_in(admin_client, admin_id)
            assert admin_client.post(f"/registrierungen/{person_id}/approve", data=h.csrf_data(token=csrf)).status_code == 302
            with h.Session(engine) as session:
                assert session.get(db.Person, person_id).account_status == db.ACCOUNT_ACTIVE
            client = app.test_client()
            response = _request_login(client, _csrf(client), contact="private-contact@example.test")
            assert response.status_code == 200 and _challenge(response)
            assert "Falls die Angaben bekannt sind" in response.get_data(as_text=True)
            assert not sent, "invalid rendered subjects must never reach account dispatch"
        diagnostics = log.getvalue()
        assert "auth_delivery_failed" in diagnostics
        assert "message_key=account.welcome" in diagnostics
        assert "message_key=auth.login_code" in diagnostics
        for private_value in ("private-name-canary", "private-contact@example.test", "private-template-canary"):
            assert private_value not in diagnostics, "best-effort error logs must contain identifiers only"
    finally:
        auth_abuse.LOGGER.removeHandler(handler)
        _restore_messages(originals)


if __name__ == "__main__":
    h.run_all(globals())
