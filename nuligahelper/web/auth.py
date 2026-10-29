# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Passwordless login and registration pages, registration status
# and admin approval decisions. Shared account helpers live in
# webapp.challenges.
# ---------------------------------------------------------------

from datetime import datetime

from flask import (
    flash,
    g,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from sqlalchemy import select
from sqlalchemy.exc import IntegrityError

from nuligahelper import auth_abuse
from nuligahelper import common
from nuligahelper import db

from nuligahelper.web.challenges import (
    _auth_values,
    _challenge_record,
    _consume_challenge,
    _consume_legacy_link,
    _contact_person,
    _decode_challenge,
    _dummy_challenge,
    _notify_registration_approved,
    _notify_registration_approvers,
    _request_rules,
    _reserve,
    _safe_account_message_via,
    _send_challenge,
    _validated_auth_contact,
    _validated_registration_contacts,
)
from nuligahelper.web.context import api_error, get_session, Registry
from nuligahelper.web.formatting import COUNTRY_CODES

routes = Registry()


def _render_login(
    *,
    values: dict[str, str] | None = None,
    errors: dict[str, str] | None = None,
    challenge: str | None = None,
    message: str | None = None,
):
    payload = _decode_challenge(challenge, "login")
    if message:
        flash(message, "info")
    return render_template(
        "login.html",
        values=values or {
            "channel": "email",
            "email": "",
            "phone": "",
            "country_code": "+49",
            "custom_country_code": "",
        },
        errors=errors or {},
        challenge=challenge if payload else None,
        masked_destination=(
            payload.get("masked_destination") if payload else None
        ),
        country_codes=COUNTRY_CODES,
    )


def _render_register(
    *,
    values: dict[str, str] | None = None,
    errors: dict[str, str] | None = None,
    challenge: str | None = None,
    message: str | None = None,
):
    payload = _decode_challenge(challenge, "verify")
    if message:
        flash(message, "info")
    initial = {
        "name": "",
        "birth_date": "",
        "team_ids": [],
        "consent": "",
        "channel": "email",
        "email": "",
        "phone": "",
        "country_code": "+49",
        "custom_country_code": "",
    }
    return render_template(
        "register.html",
        teams=db.get_all_teams(get_session()),
        values=values or initial,
        errors=errors or {},
        challenge=challenge if payload else None,
        masked_destination=(
            payload.get("masked_destination") if payload else None
        ),
        country_codes=COUNTRY_CODES,
        birth_date_max=common.effective_today().isoformat(),
    )


def _establish_session(person: db.Person):
    session.clear()
    session["person_id"] = person.id
    session.permanent = True
    session["feedback_auth_boundary"] = True


@routes.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "GET":
        return _render_login()
    action = request.form.get("action") or "request_code"
    if action == "reset":
        return redirect(url_for("login"))
    if action == "confirm_code":
        signed_challenge = request.form.get("challenge")
        payload, _ = _challenge_record("login", signed_challenge)
        rules = [auth_abuse.Rule("confirmation_client", "@client")]
        if payload is not None:
            rules.extend([
                auth_abuse.Rule("confirmation_contact", payload["contact_subject"], True),
                auth_abuse.Rule("confirmation_person", payload["person_subject"], True),
            ])
        allowed = _reserve(rules).allowed
        person = (
            _consume_challenge(
                "login",
                signed_challenge,
                (request.form.get("code") or "").strip(),
            )
            if allowed
            else None
        )
        if person is None or person.account_status not in (
            db.ACCOUNT_VERIFIED,
            db.ACCOUNT_ACTIVE,
        ):
            return _render_login(
                errors={"code": "Code ungültig oder abgelaufen."},
                challenge=signed_challenge,
            )
        _establish_session(person)
        flash("Du bist angemeldet.", "success")
        return redirect(
            url_for("schedule")
            if person.account_status == db.ACCOUNT_ACTIVE
            else url_for("registration_status")
        )

    values = _auth_values()
    channel, destination, errors = _validated_auth_contact(values)
    if errors:
        return _render_login(values=values, errors=errors)
    assert channel is not None and destination is not None
    person = _contact_person(channel, destination)
    decision = _reserve(_request_rules(
        "login", channel, [(channel, destination)],
        [person] if person is not None else [],
    ))
    challenge = None
    if (
        person is not None
        and person.account_status in (db.ACCOUNT_VERIFIED, db.ACCOUNT_ACTIVE)
        and decision.allowed
    ):
        challenge = _send_challenge(person, "login", channel, destination)
    if challenge is None:
        challenge = _dummy_challenge("login", channel, destination)
    return _render_login(
        values=values,
        challenge=challenge,
        message=(
            "Falls die Angaben bekannt sind, wurde ein sechsstelliger "
            "Code versendet."
        ),
    )


@routes.route("/login/token/<token>")
def login_token(token: str):
    person = _consume_legacy_link("login", token)
    if person is None or person.account_status not in (
        db.ACCOUNT_VERIFIED,
        db.ACCOUNT_ACTIVE,
    ):
        return render_template(
            "message.html",
            message="Anmeldung ungültig oder abgelaufen.",
        ), 400
    _establish_session(person)
    flash("Du bist angemeldet.", "success")
    return redirect(
        url_for("schedule")
        if person.account_status == db.ACCOUNT_ACTIVE
        else url_for("registration_status")
    )


@routes.route("/login/code", methods=["GET", "POST"])
def login_code():
    return redirect(url_for("login"), code=303)


@routes.post("/logout")
def logout():
    session.clear()
    session["feedback_auth_boundary"] = True
    flash("Du bist abgemeldet.", "success")
    return redirect(url_for("schedule"))


@routes.route("/registrieren", methods=["GET", "POST"])
def register():
    if request.method == "GET":
        return _render_register()
    action = request.form.get("action") or "request_code"
    if action == "reset":
        return redirect(url_for("register"))
    if action == "confirm_code":
        signed_challenge = request.form.get("challenge")
        payload, _ = _challenge_record("verify", signed_challenge)
        rules = [auth_abuse.Rule("confirmation_client", "@client")]
        if payload is not None:
            rules.extend([
                auth_abuse.Rule("confirmation_contact", payload["contact_subject"], True),
                auth_abuse.Rule("confirmation_person", payload["person_subject"], True),
            ])
        allowed = _reserve(rules).allowed
        person = (
            _consume_challenge(
                "verify",
                signed_challenge,
                (request.form.get("code") or "").strip(),
            )
            if allowed
            else None
        )
        if person is None or person.account_status != db.ACCOUNT_REGISTERED:
            return _render_register(
                errors={"code": "Code ungültig oder abgelaufen."},
                challenge=signed_challenge,
            )
        db.verify_person(get_session(), person)
        get_session().commit()
        _notify_registration_approvers(person)
        _establish_session(person)
        flash("Kontakt bestätigt. Die Freigabe steht noch aus.", "success")
        return redirect(url_for("registration_status"))

    values = _auth_values()
    values.update({
        "name": (request.form.get("name") or "").strip(),
        "team_ids": request.form.getlist("team_ids"),
        "consent": request.form.get("consent") or "",
    })
    errors: dict[str, str] = {}
    if not values["name"]:
        errors["name"] = "Bitte gib deinen Namen ein."
    birth_date = None
    try:
        birth_date = db.validate_birth_date(values["birth_date"])
    except ValueError as exc:
        errors["birth_date"] = str(exc)
    if values["consent"] != "yes":
        errors["consent"] = (
            "Die Zustimmung zur Veröffentlichung des Namens ist erforderlich."
        )
    try:
        selected_team_ids = {int(value) for value in values["team_ids"]}
    except (TypeError, ValueError):
        selected_team_ids = set()
    selected_teams = list(get_session().scalars(
        select(db.Team).where(db.Team.id.in_(selected_team_ids))
    )) if selected_team_ids else []
    if not selected_team_ids or {team.id for team in selected_teams} != selected_team_ids:
        errors["team_ids"] = "Bitte wähle mindestens eine gültige Mannschaft."
    email, phone, channel, destination, contact_errors = (
        _validated_registration_contacts(values)
    )
    errors.update(contact_errors)
    if errors:
        return _render_register(values=values, errors=errors)
    assert channel is not None and destination is not None and selected_teams

    email_person = _contact_person("email", email) if email else None
    phone_person = _contact_person("sms", phone) if phone else None
    selected_person = email_person if channel == "email" else phone_person
    conflicting_people = {
        person.id: person
        for person in (email_person, phone_person)
        if person is not None
    }
    limit_contacts = []
    if email:
        limit_contacts.append(("email", email))
    if phone:
        limit_contacts.append(("sms", phone))
    decision = _reserve(_request_rules(
        "registration", channel, limit_contacts,
        list(conflicting_people.values()),
    ))
    challenge = None
    if not conflicting_people and decision.allowed:
        try:
            person = db.register_person(
                get_session(), values["name"], selected_teams, email, phone,
                birth_date=birth_date,
            )
            get_session().commit()
        except IntegrityError:
            get_session().rollback()
        else:
            challenge = _send_challenge(
                person, "verify", channel, destination
            )
    elif (
        selected_person is not None
        and len(conflicting_people) == 1
        and decision.allowed
    ):
        if selected_person.account_status == db.ACCOUNT_REGISTERED:
            challenge = _send_challenge(
                selected_person, "verify", channel, destination
            )
        else:
            _safe_account_message_via(
                selected_person,
                channel,
                "auth.existing_account",
            )
    if challenge is None:
        challenge = _dummy_challenge("verify", channel, destination)
    return _render_register(
        values=values,
        challenge=challenge,
        message=(
            "Falls die Angaben verwendet werden können, wurde ein "
            "sechsstelliger Code versendet."
        ),
    )


@routes.route("/registrieren/code", methods=["GET", "POST"])
def registration_code():
    return redirect(url_for("register"), code=303)


@routes.route("/registrieren/verifizieren/<token>")
def verify_registration(token: str):
    person = _consume_legacy_link("verify", token)
    if person is None or person.account_status != db.ACCOUNT_REGISTERED:
        return render_template(
            "message.html",
            message="Bestätigung ungültig oder abgelaufen.",
        ), 400
    db.verify_person(get_session(), person)
    get_session().commit()
    _notify_registration_approvers(person)
    session["feedback_auth_boundary"] = True
    flash("Kontakt bestätigt.", "success")
    return render_template(
        "message.html",
        message="Deine Registrierung wartet auf Freigabe.",
    )


@routes.route("/registrierung/status")
def registration_status():
    if g.viewer is None or g.viewer.account_status != db.ACCOUNT_VERIFIED:
        return redirect(url_for("login"))
    return render_template("message.html", message="Deine Registrierung wartet auf Freigabe.")


@routes.post("/registrierungen/<int:person_id>/<decision>")
def decide_registration(person_id: int, decision: str):
    person = get_session().get(db.Person, person_id)
    if person is None or person.account_status != db.ACCOUNT_VERIFIED:
        return api_error("Registrierung nicht gefunden.", 404)
    if g.tier != "admin":
        return api_error("Keine Berechtigung.", 403)
    if decision == "approve":
        try:
            db.approve_person(get_session(), person)
        except ValueError as exc:
            return api_error(str(exc))
    elif decision == "reject":
        person.account_status = db.ACCOUNT_REJECTED
        person.rejected_at = datetime.now()
    else:
        return api_error("Ungültige Entscheidung.")
    get_session().commit()
    if decision == "approve":
        _notify_registration_approved(person)
    flash(f"Registrierung von '{person.name}' wurde "
          + ("freigegeben." if decision == "approve" else "abgelehnt."), "success")
    return redirect(url_for("persons"))


def register_routes(app) -> None:
    routes.register(app)
