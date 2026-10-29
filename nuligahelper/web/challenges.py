# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Account helpers shared by login, registration and helper
# management: contact lookup and validation, abuse limits, signed
# one-time challenges and account messages.
# ---------------------------------------------------------------

import secrets
from datetime import datetime, timedelta

from flask import request
from itsdangerous import BadSignature, SignatureExpired
from sqlalchemy import select, update

from nuligahelper import auth_abuse
from nuligahelper import common
from nuligahelper.db import contacts
from nuligahelper import db
from nuligahelper.jobs import messages
from nuligahelper.jobs import notifier

from nuligahelper.web.context import get_session, state


def _contact_person(
    channel: str,
    contact: str,
    statuses: tuple[str, ...] | None = None,
) -> db.Person | None:
    column = db.Person.email if channel == "email" else db.Person.phone
    query = get_session().query(db.Person).filter(column == contact)
    if statuses:
        query = query.filter(db.Person.account_status.in_(statuses))
    return query.order_by(db.Person.id).first()


def _contact_in_use(
    channel: str,
    contact: str | None,
    exclude_person_id: int | None = None,
) -> bool:
    if not contact:
        return False
    person = _contact_person(channel, contact)
    return person is not None and person.id != exclude_person_id


def _auth_values() -> dict[str, str]:
    country_code = (request.form.get("country_code") or "+49").strip()
    return {
        "channel": (request.form.get("channel") or "").strip(),
        "birth_date": (request.form.get("birth_date") or "").strip(),
        "email": (request.form.get("email") or "").strip(),
        "phone": (request.form.get("phone") or "").strip(),
        "country_code": country_code,
        "custom_country_code": (
            request.form.get("custom_country_code") or ""
        ).strip(),
    }


def _validated_auth_contact(
    values: dict[str, str],
) -> tuple[str | None, str | None, dict[str, str]]:
    channel = values["channel"]
    try:
        if channel == "email":
            return (
                channel,
                contacts.normalize_email(values["email"], required=True),
                {},
            )
        if channel == "sms":
            calling_code = (
                values["custom_country_code"]
                if values["country_code"] == "custom"
                else values["country_code"]
            )
            return (
                channel,
                contacts.normalize_phone(
                    values["phone"], calling_code, required=True
                ),
                {},
            )
        return None, None, {
            "channel": "Bitte wähle E-Mail oder SMS als Kontaktweg."
        }
    except contacts.ContactValidationError as exc:
        return channel, None, {exc.field_name: exc.message}


def _validated_registration_contacts(
    values: dict[str, str],
) -> tuple[
    str | None,
    str | None,
    str | None,
    str | None,
    dict[str, str],
]:
    errors: dict[str, str] = {}
    email = phone = None
    try:
        email = contacts.normalize_email(values["email"])
    except contacts.ContactValidationError as exc:
        errors[exc.field_name] = exc.message

    calling_code = (
        values["custom_country_code"]
        if values["country_code"] == "custom"
        else values["country_code"]
    )
    try:
        phone = contacts.normalize_phone(values["phone"], calling_code)
    except contacts.ContactValidationError as exc:
        errors[exc.field_name] = exc.message

    channel = values["channel"]
    if not email and not phone and not errors:
        errors["channel"] = "Bitte gib eine E-Mail-Adresse oder Mobilnummer ein."
    elif channel not in {"email", "sms"}:
        errors["channel"] = "Bitte wähle E-Mail oder SMS als Kontaktweg."
    elif (
        (channel == "email" and not email)
        or (channel == "sms" and not phone)
    ):
        errors["channel"] = "Bitte wähle einen gültigen vorhandenen Kontaktweg."

    destination = {"email": email, "sms": phone}.get(channel)
    return email, phone, channel, destination, errors


def _normalized_person_contacts() -> tuple[str | None, str | None, dict[str, str]]:
    errors: dict[str, str] = {}
    email = phone = None
    try:
        email = contacts.normalize_email(request.form.get("email"))
    except contacts.ContactValidationError as exc:
        errors[exc.field_name] = exc.message
    try:
        phone = contacts.normalize_phone(request.form.get("phone"))
    except contacts.ContactValidationError as exc:
        errors[exc.field_name] = exc.message
    return email, phone, errors


def _client_subject() -> str | None:
    return auth_abuse.attributed_client(
        request.remote_addr,
        request.headers.get("X-Forwarded-For"),
        state().abuse_config,
    )


def _reserve(rules: list[auth_abuse.Rule]) -> auth_abuse.Decision:
    client = _client_subject()
    if client is None:
        auth_abuse.LOGGER.error(
            "auth_abuse_storage_error action=client_attribution channel=any reason=invalid_proxy_metadata"
        )
        return auth_abuse.Decision(False, (), True)
    return state().abuse_service.reserve([
        auth_abuse.Rule(rule.policy, client if rule.subject == "@client" else rule.subject,
                        rule.prehashed)
        for rule in rules
    ])


def _request_rules(
    action: str,
    channel: str,
    contacts_for_limits: list[tuple[str, str]],
    people: list[db.Person],
) -> list[auth_abuse.Rule]:
    rules = [auth_abuse.Rule(f"{action}_client", "@client")]
    for contact_channel, contact in contacts_for_limits:
        rules.append(auth_abuse.Rule(
            f"{action}_contact_{contact_channel}", contact
        ))
    for person in {person.id: person for person in people}.values():
        rules.append(auth_abuse.Rule(f"{action}_person_{channel}", str(person.id)))
    if channel == "sms":
        selected_contacts = [value for kind, value in contacts_for_limits if kind == "sms"]
        for contact in selected_contacts:
            rules.append(auth_abuse.Rule("sms_contact_cap", contact))
        for person in {person.id: person for person in people}.values():
            rules.append(auth_abuse.Rule("sms_person_cap", str(person.id)))
        rules.append(auth_abuse.Rule("sms_global_cap", "application"))
    return rules


def _account_notifier():
    club_config = common.load_config()["club"]
    return notifier.Notifier(
        club_config,
        get_session(),
        common.season_year_for(common.effective_today()),
    )


def _challenge_payload(
    nonce: str, purpose: str, channel: str, masked_destination: str,
    contact_subject: str, person_subject: str, team_ids: tuple[int, ...] = (),
) -> str:
    return state().serializer.dumps({
        "nonce": nonce,
        "purpose": purpose,
        "channel": channel,
        "masked_destination": masked_destination,
        "contact_subject": contact_subject,
        "person_subject": person_subject,
        "team_ids": list(team_ids),
    })


def _decode_challenge(
    signed_challenge: str | None, purpose: str
) -> dict | None:
    if not signed_challenge:
        return None
    try:
        payload = state().serializer.loads(signed_challenge, max_age=15 * 60)
    except (BadSignature, SignatureExpired):
        return None
    if (
        payload.get("purpose") != purpose
        or not isinstance(payload.get("nonce"), str)
        or payload.get("channel") not in {"email", "sms"}
        or not isinstance(payload.get("contact_subject"), str)
        or len(payload["contact_subject"]) != 64
        or not isinstance(payload.get("person_subject"), str)
        or len(payload["person_subject"]) != 64
        or not isinstance(payload.get("team_ids"), list)
        or any(not isinstance(team_id, int) for team_id in payload["team_ids"])
    ):
        return None
    return payload


def _issue_challenge(
    person: db.Person, purpose: str, channel: str, destination: str
) -> tuple[str, str]:
    now = datetime.now()
    get_session().query(db.AuthToken).filter(
        db.AuthToken.person_id == person.id,
        db.AuthToken.purpose == purpose,
        db.AuthToken.used_at.is_(None),
    ).update({db.AuthToken.used_at: now}, synchronize_session=False)
    nonce = secrets.token_urlsafe(24)
    code = f"{secrets.randbelow(1_000_000):06d}"
    get_session().add(db.AuthToken(
        nonce=nonce,
        code=code,
        purpose=purpose,
        person=person,
        issued_at=now,
        expires_at=now + timedelta(minutes=15),
    ))
    get_session().commit()
    return (
        _challenge_payload(
            nonce, purpose, channel, contacts.mask_contact(channel, destination),
            state().abuse_service.digest("contact", channel, destination),
            state().abuse_service.digest("person", channel, str(person.id)),
            db.membership_team_ids(person) if purpose == "verify" else (),
        ),
        code,
    )


def _dummy_challenge(
    purpose: str, channel: str, destination: str
) -> str:
    return _challenge_payload(
        secrets.token_urlsafe(24),
        purpose,
        channel,
        contacts.mask_contact(channel, destination),
        state().abuse_service.digest("contact", channel, destination),
        state().abuse_service.digest("person", channel, f"dummy:{destination}"),
    )


def _challenge_record(
    purpose: str, signed_challenge: str | None
) -> tuple[dict | None, db.AuthToken | None]:
    payload = _decode_challenge(signed_challenge, purpose)
    if payload is None:
        return None, None
    record = get_session().query(db.AuthToken).filter(
        db.AuthToken.nonce == payload["nonce"],
        db.AuthToken.purpose == purpose,
        db.AuthToken.used_at.is_(None),
    ).first()
    return payload, record


def _consume_challenge(
    purpose: str, signed_challenge: str | None, code: str | None
) -> db.Person | None:
    if not code or len(code) != 6 or not code.isdigit():
        return None
    payload, record = _challenge_record(purpose, signed_challenge)
    now = datetime.now()
    if (
        record is None
        or record.code != code
        or record.expires_at < now
    ):
        return None
    if purpose == "verify":
        challenge_person = get_session().get(db.Person, record.person_id)
        if (
            challenge_person is None
            or tuple(sorted(payload["team_ids"]))
            != tuple(sorted(db.membership_team_ids(challenge_person)))
        ):
            return None
    person_id = record.person_id
    consumed = get_session().execute(
        update(db.AuthToken)
        .where(
            db.AuthToken.id == record.id,
            db.AuthToken.used_at.is_(None),
            db.AuthToken.expires_at >= now,
            db.AuthToken.code == code,
        )
        .values(used_at=now)
    )
    if consumed.rowcount != 1:
        get_session().rollback()
        return None
    get_session().commit()
    return get_session().get(db.Person, person_id)


def _consume_legacy_link(
    purpose: str, signed_token: str
) -> db.Person | None:
    try:
        payload = state().serializer.loads(signed_token, max_age=15 * 60)
    except (BadSignature, SignatureExpired):
        return None
    if (
        payload.get("purpose") != purpose
        or not isinstance(payload.get("nonce"), str)
    ):
        return None
    record = get_session().query(db.AuthToken).filter(
        db.AuthToken.nonce == payload["nonce"],
        db.AuthToken.purpose == purpose,
        db.AuthToken.code.is_(None),
        db.AuthToken.used_at.is_(None),
    ).first()
    now = datetime.now()
    if record is None or record.expires_at < now:
        return None
    person_id = record.person_id
    consumed = get_session().execute(
        update(db.AuthToken)
        .where(
            db.AuthToken.id == record.id,
            db.AuthToken.code.is_(None),
            db.AuthToken.used_at.is_(None),
            db.AuthToken.expires_at >= now,
        )
        .values(used_at=now)
    )
    if consumed.rowcount != 1:
        get_session().rollback()
        return None
    get_session().commit()
    return get_session().get(db.Person, person_id)


def _safe_account_message(
    person: db.Person, message_key: str, **context
) -> bool:
    try:
        rendered = messages.render(message_key, **context)
        return bool(
            _account_notifier().send_account_message(
                person, rendered.subject, rendered.email, rendered.sms
            )
        )
    except Exception as exc:
        auth_abuse.LOGGER.error(
            "auth_delivery_failed action=account_message channel=preferred message_key=%s reason=%s",
            message_key, type(exc).__name__,
        )
        return False


def _safe_account_message_via(
    person: db.Person, channel: str, message_key: str, **context
) -> bool:
    try:
        rendered = messages.render(message_key, **context)
        body = rendered.email if channel == "email" else rendered.sms
        return bool(
            _account_notifier().send_account_message_via(
                person, channel, rendered.subject, body
            )
        )
    except Exception as exc:
        auth_abuse.LOGGER.error(
            "auth_delivery_failed action=account_message channel=%s message_key=%s reason=%s",
            channel, message_key, type(exc).__name__,
        )
        return False


def _send_challenge(
    person: db.Person,
    purpose: str,
    channel: str,
    destination: str,
) -> str:
    signed_challenge, code = _issue_challenge(
        person, purpose, channel, destination
    )
    message_key = (
        "auth.registration_code" if purpose == "verify" else "auth.login_code"
    )
    _safe_account_message_via(
        person, channel, message_key, recipient_name=person.name, auth_code=code
    )
    return signed_challenge


def _notify_registration_approvers(person: db.Person) -> None:
    team_label = db.membership_label(person)
    approvers = get_session().scalars(
        select(db.Person).where(
            db.Person.is_admin.is_(True),
            db.Person.account_status == db.ACCOUNT_ACTIVE,
        ).order_by(db.Person.id)
    )
    for approver in approvers:
        _safe_account_message(
            approver,
            "account.approval_request",
            recipient_name=approver.name,
            registrant_name=person.name,
            selected_team_names=team_label,
        )


def _notify_registration_approved(person: db.Person) -> None:
    _safe_account_message(
        person,
        "account.welcome",
        recipient_name=person.name,
    )
