# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Web interface: home game schedule with inline task assignment,
# helper/team management and statistics. Designed to visually
# integrate with www.handball-raubling.de
#
# Run locally:  ./venv/bin/python webapp.py   (http://<pi-ip>:8080)
# Optional env: NULIGAHELPER_DB=/path/to/nuliga_helper.db
# ---------------------------------------------------------------

import os
import secrets
import logging
from collections import defaultdict, deque
from datetime import datetime, timedelta

import common
import contact_validation as contacts
import db
import notifier
from flask import (
    Flask,
    flash,
    g,
    jsonify,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from itsdangerous import BadSignature, SignatureExpired, URLSafeTimedSerializer
from sqlalchemy import or_, update
from sqlalchemy.exc import IntegrityError

MONATE = [
    "Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember",
]

SLOT_LABELS = [
    ("Zeitnehmer", db.ROLE_TIMEKEEPER),
    ("Sekretär", db.ROLE_SECRETARY),
    ("Verkauf 1", db.ROLE_SALE),
    ("Verkauf 2", db.ROLE_SALE),
    ("Ordnungsdienst", db.ROLE_SECURITY),
    ("Reinigung", db.ROLE_CLEANING),
]

COUNTRY_CODES = [
    ("+49", "Deutschland (+49)"),
    ("+43", "Österreich (+43)"),
    ("+41", "Schweiz (+41)"),
    ("+39", "Italien (+39)"),
    ("+33", "Frankreich (+33)"),
    ("+420", "Tschechien (+420)"),
    ("custom", "Andere Ländervorwahl"),
]


def get_db_path() -> str:
    return os.environ.get(
        "NULIGAHELPER_DB",
        common.load_config()["club"].get("database", {}).get("path", db.DEFAULT_DB_PATH),
    )


def ak_color(ak: str | None) -> str:
    """Map an age class (e.g. 'BL mD') to the club's team colors."""
    parts = (ak or "").split()
    gender = parts[1][0].lower() if len(parts) > 1 else ""
    youth = len(parts) > 1 and len(parts[1]) > 1
    if gender == "w":
        return "#00C6D7" if youth else "#7CFFCB"
    if gender == "m":
        return "#6BB32C" if youth else "#DDFF00"
    return "#FFA01F"


def parse_date(date_str: str | None):
    try:
        return datetime.strptime(date_str or "", "%d.%m.%Y").date()
    except ValueError:
        return None


def display_time(time_str: str | None) -> str:
    parts = (time_str or "").split()
    return parts[0] if parts else ""


def _person_team(persons: list[dict], person_id: int) -> int | None:
    return next((p["team_id"] for p in persons if p["id"] == person_id), None)


def create_app() -> Flask:
    app = Flask(__name__)
    secret_key = os.environ.get("NULIGAHELPER_SECRET")
    if not secret_key:
        raise RuntimeError("NULIGAHELPER_SECRET muss gesetzt sein.")
    app.secret_key = secret_key
    app.config.update(
        PERMANENT_SESSION_LIFETIME=timedelta(hours=1),
        SESSION_REFRESH_EACH_REQUEST=True,
        SESSION_COOKIE_SAMESITE="Lax",
    )
    engine = db.make_engine(get_db_path())
    db.init_db(engine)
    serializer = URLSafeTimedSerializer(secret_key, salt="nuligahelper-auth")
    rate_events: dict[str, deque[datetime]] = defaultdict(deque)

    @app.teardown_appcontext
    def close_session(exception):
        session = g.pop("session", None)
        if session is not None:
            session.close()

    def get_session():
        if "session" not in g:
            g.session = db.Session(engine)
        return g.session

    def current_mv_team_ids(person: db.Person | None) -> set[int]:
        if person is None or person.account_status != db.ACCOUNT_ACTIVE:
            return set()
        return {
            team.id
            for team in get_session().query(db.Team).filter(
                db.Team.mv_person_id == person.id
            )
        }

    def tier_for(person: db.Person | None) -> str:
        if person is None or person.account_status != db.ACCOUNT_ACTIVE:
            return "guest"
        if person.is_admin:
            return "admin"
        if current_mv_team_ids(person):
            return "mv"
        return "member"

    @app.before_request
    def load_viewer():
        g.viewer = None
        person_id = session.get("person_id")
        if person_id is not None:
            person = get_session().get(db.Person, person_id)
            if person is not None and person.account_status in (
                db.ACCOUNT_VERIFIED,
                db.ACCOUNT_ACTIVE,
            ):
                g.viewer = person
            else:
                session.clear()
        g.tier = tier_for(g.viewer)
        g.mv_team_ids = current_mv_team_ids(g.viewer)

    public_endpoints = {
        "static",
        "schedule",
        "login",
        "login_token",
        "login_code",
        "register",
        "registration_code",
        "verify_registration",
    }

    @app.before_request
    def require_authentication():
        if request.endpoint in public_endpoints:
            return None
        is_json = request.path.startswith("/api/")
        if g.viewer is None:
            if is_json:
                return jsonify(
                    ok=False,
                    code="session_expired",
                    error="Deine Sitzung ist abgelaufen. Bitte melde dich erneut an.",
                ), 401
            return redirect(url_for("login", next=request.path))
        if (
            g.viewer.account_status == db.ACCOUNT_VERIFIED
            and request.endpoint not in {"registration_status", "logout"}
        ):
            if is_json:
                return api_error("Die Registrierung ist noch nicht freigegeben.", 403)
            return redirect(url_for("registration_status"))
        return None

    @app.before_request
    def csrf_protect():
        if request.method not in {"POST", "PUT", "PATCH", "DELETE"}:
            return None
        expected = session.get("csrf_token")
        supplied = request.headers.get("X-CSRF-Token") or request.form.get("csrf_token")
        if not expected or not supplied or not secrets.compare_digest(expected, supplied):
            if request.path.startswith("/api/"):
                return api_error("Ungültiges Sicherheitstoken.", 403)
            return render_template(
                "message.html", message="Das Formular ist abgelaufen. Bitte lade die Seite neu."
            ), 403
        return None

    @app.context_processor
    def inject_globals():
        if "csrf_token" not in session:
            session["csrf_token"] = secrets.token_urlsafe(24)
        return {
            "active_page": request.path,
            "viewer": g.get("viewer"),
            "viewer_tier": g.get("tier", "guest"),
            "csrf_token": session["csrf_token"],
        }

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
            "channel": (request.form.get("channel") or "email").strip(),
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

    def _rate_allowed(key: str, limit: int, window_minutes: int = 15) -> bool:
        now = datetime.now()
        cutoff = now - timedelta(minutes=window_minutes)
        events = rate_events[key]
        while events and events[0] < cutoff:
            events.popleft()
        if len(events) >= limit:
            return False
        events.append(now)
        return True

    def _account_notifier():
        club_config = common.load_config()["club"]
        return notifier.Notifier(
            club_config,
            get_session(),
            common.season_year_for(common.effective_today()),
        )

    def _challenge_payload(
        nonce: str, purpose: str, channel: str, masked_destination: str
    ) -> str:
        return serializer.dumps({
            "nonce": nonce,
            "purpose": purpose,
            "channel": channel,
            "masked_destination": masked_destination,
        })

    def _decode_challenge(
        signed_challenge: str | None, purpose: str
    ) -> dict | None:
        if not signed_challenge:
            return None
        try:
            payload = serializer.loads(signed_challenge, max_age=15 * 60)
        except (BadSignature, SignatureExpired):
            return None
        if (
            payload.get("purpose") != purpose
            or not isinstance(payload.get("nonce"), str)
            or payload.get("channel") not in {"email", "sms"}
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
                nonce, purpose, channel, contacts.mask_contact(channel, destination)
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
        _, record = _challenge_record(purpose, signed_challenge)
        now = datetime.now()
        if (
            record is None
            or record.code != code
            or record.expires_at < now
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
            payload = serializer.loads(signed_token, max_age=15 * 60)
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
        person: db.Person, subject: str, mail_body: str, sms_body: str
    ) -> bool:
        try:
            return bool(
                _account_notifier().send_account_message(
                    person, subject, mail_body, sms_body
                )
            )
        except Exception:
            logging.exception("Account message delivery failed for person ID %s", person.id)
            return False

    def _safe_account_message_via(
        person: db.Person, channel: str, subject: str, body: str
    ) -> bool:
        try:
            return bool(
                _account_notifier().send_account_message_via(
                    person, channel, subject, body
                )
            )
        except Exception:
            logging.exception(
                "Account message delivery failed for person ID %s via %s",
                person.id,
                channel,
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
        action = "Registrierung" if purpose == "verify" else "Anmeldung"
        body = (
            f"Hallo {person.name},\n\n"
            f"dein Code für die {action} bei nuLigaHelper lautet {code}. "
            "Er gilt 15 Minuten."
        )
        _safe_account_message_via(
            person, channel, f"{action} nuLigaHelper", body
        )
        return signed_challenge

    def _render_login(
        *,
        values: dict[str, str] | None = None,
        errors: dict[str, str] | None = None,
        challenge: str | None = None,
        message: str | None = None,
    ):
        payload = _decode_challenge(challenge, "login")
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
            request_message=message,
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
        initial = {
            "name": "",
            "team_id": "",
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
            request_message=message,
            country_codes=COUNTRY_CODES,
        )

    def _establish_session(person: db.Person):
        session.clear()
        session["person_id"] = person.id
        session.permanent = True

    @app.route("/login", methods=["GET", "POST"])
    def login():
        if request.method == "GET":
            return _render_login()
        action = request.form.get("action") or "request_code"
        if action == "reset":
            return redirect(url_for("login"))
        if action == "confirm_code":
            signed_challenge = request.form.get("challenge")
            _, record = _challenge_record("login", signed_challenge)
            client_allowed = _rate_allowed(
                f"confirm-client:{request.remote_addr or 'unknown'}", 10
            )
            person_allowed = record is not None and _rate_allowed(
                f"confirm-person:{record.person_id}", 5
            )
            person = (
                _consume_challenge(
                    "login",
                    signed_challenge,
                    (request.form.get("code") or "").strip(),
                )
                if client_allowed and person_allowed
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
        client_allowed = _rate_allowed(
            f"login-client:{request.remote_addr or 'unknown'}", 10
        )
        person = _contact_person(
            channel,
            destination,
            (db.ACCOUNT_VERIFIED, db.ACCOUNT_ACTIVE),
        )
        challenge = None
        if person is not None:
            channel_limit = 2 if channel == "sms" else 3
            person_allowed = _rate_allowed(
                f"login-person:{person.id}:{channel}", channel_limit
            )
            if client_allowed and person_allowed:
                challenge = _send_challenge(
                    person, "login", channel, destination
                )
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

    @app.route("/login/token/<token>")
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
        return redirect(
            url_for("schedule")
            if person.account_status == db.ACCOUNT_ACTIVE
            else url_for("registration_status")
        )

    @app.route("/login/code", methods=["GET", "POST"])
    def login_code():
        return redirect(url_for("login"), code=303)

    @app.post("/logout")
    def logout():
        session.clear()
        return redirect(url_for("schedule"))

    @app.route("/registrieren", methods=["GET", "POST"])
    def register():
        if request.method == "GET":
            return _render_register()
        action = request.form.get("action") or "request_code"
        if action == "reset":
            return redirect(url_for("register"))
        if action == "confirm_code":
            signed_challenge = request.form.get("challenge")
            _, record = _challenge_record("verify", signed_challenge)
            client_allowed = _rate_allowed(
                f"confirm-client:{request.remote_addr or 'unknown'}", 10
            )
            person_allowed = record is not None and _rate_allowed(
                f"confirm-person:{record.person_id}", 5
            )
            person = (
                _consume_challenge(
                    "verify",
                    signed_challenge,
                    (request.form.get("code") or "").strip(),
                )
                if client_allowed and person_allowed
                else None
            )
            if person is None or person.account_status != db.ACCOUNT_REGISTERED:
                return _render_register(
                    errors={"code": "Code ungültig oder abgelaufen."},
                    challenge=signed_challenge,
                )
            db.verify_person(get_session(), person)
            get_session().commit()
            _notify_registration_approver(person)
            _establish_session(person)
            return redirect(url_for("registration_status"))

        values = _auth_values()
        values.update({
            "name": (request.form.get("name") or "").strip(),
            "team_id": (request.form.get("team_id") or "").strip(),
            "consent": request.form.get("consent") or "",
        })
        errors: dict[str, str] = {}
        if not values["name"]:
            errors["name"] = "Bitte gib deinen Namen ein."
        if values["consent"] != "yes":
            errors["consent"] = (
                "Die Zustimmung zur Veröffentlichung des Namens ist erforderlich."
            )
        try:
            team_id = int(values["team_id"])
        except (TypeError, ValueError):
            team = None
        else:
            team = get_session().get(db.Team, team_id)
        if team is None:
            errors["team_id"] = "Bitte wähle eine gültige Mannschaft."
        channel, destination, contact_errors = _validated_auth_contact(values)
        errors.update(contact_errors)
        if errors:
            return _render_register(values=values, errors=errors)
        assert channel is not None and destination is not None and team is not None

        client_allowed = _rate_allowed(
            f"register-client:{request.remote_addr or 'unknown'}", 5
        )
        existing = _contact_person(channel, destination)
        challenge = None
        if existing is None and client_allowed:
            email = destination if channel == "email" else None
            phone = destination if channel == "sms" else None
            try:
                person = db.register_person(
                    get_session(), values["name"], team, email, phone
                )
                get_session().commit()
            except IntegrityError:
                get_session().rollback()
            else:
                challenge = _send_challenge(
                    person, "verify", channel, destination
                )
        elif existing is not None and client_allowed:
            person_limit = 2 if channel == "sms" else 3
            person_allowed = _rate_allowed(
                f"register-person:{existing.id}:{channel}", person_limit
            )
            if person_allowed and existing.account_status == db.ACCOUNT_REGISTERED:
                challenge = _send_challenge(
                    existing, "verify", channel, destination
                )
            elif person_allowed:
                _safe_account_message_via(
                    existing,
                    channel,
                    "Registrierung nuLigaHelper",
                    (
                        "Für diesen Kontakt besteht bereits ein Konto. "
                        "Bitte nutze die Anmeldung."
                    ),
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

    @app.route("/registrieren/code", methods=["GET", "POST"])
    def registration_code():
        return redirect(url_for("register"), code=303)

    @app.route("/registrieren/verifizieren/<token>")
    def verify_registration(token: str):
        person = _consume_legacy_link("verify", token)
        if person is None or person.account_status != db.ACCOUNT_REGISTERED:
            return render_template(
                "message.html",
                message="Bestätigung ungültig oder abgelaufen.",
            ), 400
        db.verify_person(get_session(), person)
        get_session().commit()
        _notify_registration_approver(person)
        return render_template(
            "message.html",
            message="Kontakt bestätigt. Die Freigabe steht noch aus.",
        )

    def _notify_registration_approver(person: db.Person) -> None:
        team = person.desired_team
        approver = None if team is None or team.is_support else team.mv_person
        if approver is None:
            approver = get_session().query(db.Person).filter(
                db.Person.is_admin.is_(True),
                db.Person.account_status == db.ACCOUNT_ACTIVE,
            ).order_by(db.Person.id).first()
        if approver is not None:
            _safe_account_message(
                approver,
                "Neue Registrierung",
                f"{person.name} wartet auf Freigabe für {team.name if team else '?' }.",
                f"Neue Registrierung: {person.name} ({team.name if team else '?' }).",
            )

    @app.route("/registrierung/status")
    def registration_status():
        if g.viewer is None or g.viewer.account_status != db.ACCOUNT_VERIFIED:
            return redirect(url_for("login"))
        return render_template("message.html", message="Deine Registrierung wartet auf Freigabe.")

    @app.post("/registrierungen/<int:person_id>/<decision>")
    def decide_registration(person_id: int, decision: str):
        person = get_session().get(db.Person, person_id)
        if person is None or person.account_status != db.ACCOUNT_VERIFIED:
            return api_error("Registrierung nicht gefunden.", 404)
        allowed = g.tier == "admin" or (
            g.tier == "mv" and person.desired_team_id in g.mv_team_ids
            and not person.desired_team.is_support
        )
        if not allowed:
            return api_error("Keine Berechtigung.", 403)
        if decision == "approve":
            db.approve_person(get_session(), person)
        elif decision == "reject":
            person.account_status = db.ACCOUNT_REJECTED
        else:
            return api_error("Ungültige Entscheidung.")
        get_session().commit()
        return redirect(url_for("persons"))

    def person_options(session) -> list[dict]:
        return [
            {
                "id": p.id,
                "name": p.name,
                "team_id": p.team_id,
                "team_name": p.team.name if p.team else "",
            }
            for p in db.get_all_persons(session)
        ]

    def team_options(session) -> list[dict]:
        return [
            {
                "id": t.id,
                "name": t.name,
                "is_support": t.is_support,
                "mv_person_id": t.mv_person_id,
                "members": [
                    {"id": p.id, "name": p.name}
                    for p in db.get_team_members(session, t)
                ],
            }
            for t in db.get_all_teams(session)
        ]

    # ------------------------------------------------------------------
    # Schedule overview
    # ------------------------------------------------------------------

    def build_schedule(session, season_year: int, viewer: db.Person | None = None):
        today = common.effective_today()
        games = session.query(db.Game).filter(
            db.Game.season_year == season_year
        ).all()
        games.sort(key=db.game_sort_key)

        persons = person_options(session)
        persons_by_id = {p["id"]: p for p in persons}
        teams = team_options(session)
        support = db.get_support_team(session)
        support_id = support.id if support else None
        playing_team_by_ak = {t["name"]: t["id"] for t in teams}

        def game_view(game):
            sales = {
                assignment.slot: assignment
                for assignment in game.assignments_by_role(db.ROLE_SALE)
            }
            responsible_team_id = game.team_id
            # the age class of the game itself identifies the team that PLAYS
            playing_team_id = playing_team_by_ak.get(game.ak or "")
            slots = []
            sale_idx = 0
            for label, role in SLOT_LABELS:
                if role == db.ROLE_SALE:
                    assignment = sales.get(sale_idx)
                    slot = sale_idx
                    sale_idx += 1
                else:
                    assignment = game.assignment_by_role(role)
                    slot = 0
                person_id = assignment.person_id if assignment is not None else None
                occupant = assignment.person if assignment is not None else None

                if person_id is None:
                    status = "none"
                elif playing_team_id and occupant.team_id == playing_team_id:
                    status = "playing"
                elif (responsible_team_id is not None
                      and occupant.team_id not in (responsible_team_id, support_id)):
                    status = "outside"
                else:
                    status = "ok"

                is_admin = g.tier == "admin"
                self_allowed = g.tier in {"member", "mv", "admin"} and viewer is not None and (
                    person_id is None or person_id == viewer.id
                )
                mv_game_team = (
                    responsible_team_id
                    if responsible_team_id in g.mv_team_ids
                    else None
                )
                mv_allowed = mv_game_team is not None and (
                    occupant is None or occupant.team_id == mv_game_team
                )
                editable = is_admin or (
                    not bool(d := parse_date(game.date)) or d >= today
                ) and (self_allowed or mv_allowed)
                if is_admin:
                    options = persons
                elif editable and mv_game_team is not None:
                    options = [
                        p for p in persons
                        if p["team_id"] == mv_game_team or p["id"] == viewer.id
                    ]
                elif editable and viewer is not None:
                    options = [persons_by_id[viewer.id]] if viewer.id in persons_by_id else []
                else:
                    options = []
                slots.append({
                    "label": label,
                    "role": role,
                    "slot": slot,
                    "person_id": person_id,
                    "person_name": occupant.name if occupant else "",
                    "person_team_name": occupant.team.name if occupant and occupant.team else "",
                    "status": status,
                    "editable": editable,
                    "options": options,
                })
            # Persons already assigned to a task of this game must not be
            # offered for the other tasks of the same game.
            taken_person_ids = {s["person_id"] for s in slots if s["person_id"] is not None}
            d = parse_date(game.date)
            return {
                "id": game.id,
                "nr": game.game_nr,
                "time": display_time(game.time),
                "day": game.day or "",
                "date": game.date or "",
                "ak": game.ak or "",
                "color": ak_color(game.ak),
                "home": game.home or "",
                "guest": game.guest or "",
                "hall": game.hall,
                "team_id": responsible_team_id,
                "playing_team_id": playing_team_id,
                "slots": slots,
                "taken_person_ids": taken_person_ids,
                "past": bool(d and d < today),
            }

        day_groups = []
        for game in games:
            view = game_view(game)
            d = parse_date(view["date"])
            month_label = f"{MONATE[d.month - 1]} {d.year}" if d else "Ohne Datum"
            if not day_groups or day_groups[-1]["date"] != view["date"]:
                day_groups.append({
                    "type": "day",
                    "month": month_label,
                    "day": view["day"],
                    "date": view["date"],
                    "games": [],
                })
            day_groups[-1]["games"].append(view)

        def with_month_headers(day_list):
            result = []
            current_month = None
            for day_group in day_list:
                if day_group["month"] != current_month:
                    result.append({"type": "month", "label": day_group["month"]})
                    current_month = day_group["month"]
                result.append(day_group)
            return result

        is_past = lambda dg: all(gm["past"] for gm in dg["games"])
        upcoming = [dg for dg in day_groups if not is_past(dg)]
        past = [dg for dg in day_groups if is_past(dg)]

        return {
            "upcoming": with_month_headers(upcoming),
            "past": with_month_headers(past),
            "persons": persons,
            "teams": teams,
            "support_id": support_id,
        }

    @app.route("/")
    def schedule():
        session = get_session()
        season_year = common.season_year_for(common.effective_today())
        data = build_schedule(session, season_year, g.viewer)
        return render_template(
            "schedule.html",
            upcoming=data["upcoming"],
            past=data["past"],
            persons=data["persons"],
            teams=data["teams"],
            support_id=data["support_id"],
            season=f"{season_year}/{str(season_year + 1)[-2:]}",
        )

    # ------------------------------------------------------------------
    # Statistics
    # ------------------------------------------------------------------

    @app.route("/statistik")
    def statistics():
        session = get_session()
        season_year = common.season_year_for(common.effective_today())
        today = common.effective_today()

        games = session.query(db.Game).filter(
            db.Game.season_year == season_year
        ).all()
        games.sort(key=db.game_sort_key)
        total_games = len(games)

        team_stats = []
        for team in db.get_all_teams(session):
            covered = sum(1 for gm in games if gm.team_id == team.id)
            share = round(100 * covered / total_games) if total_games else 0
            team_stats.append({
                "name": team.name,
                "is_support": team.is_support,
                "covered": covered,
                "share": share,
                "bar_width": max(share, 6) if covered else 0,
            })
        team_stats.sort(key=lambda t: (-t["covered"], t["name"]))

        season_game_ids = {gm.id for gm in games}
        person_stats = []
        for person in db.get_all_person_records(session):
            if person.account_status not in (db.ACCOUNT_ACTIVE, db.ACCOUNT_INACTIVE):
                continue
            assignments = [
                a for a in person.assignments if a.game_id in season_game_ids
            ]
            if not assignments:
                continue
            role_counts: dict[str, int] = {}
            for a in assignments:
                role_counts[a.role] = role_counts.get(a.role, 0) + 1
            person_stats.append({
                "name": person.name,
                "team_name": person.team.name if person.team else "",
                "jobs": len(assignments),
                "roles": sorted(role_counts.items(), key=lambda kv: (-kv[1], kv[0])),
            })
        person_stats.sort(key=lambda p: (-p["jobs"], p["name"]))

        gaps = []
        for game in games:
            d = parse_date(game.date)
            if d is None or d < today:
                continue
            missing_roles = db.missing_slots(game)
            if missing_roles:
                gaps.append({
                    "nr": game.game_nr,
                    "date": game.date,
                    "time": display_time(game.time),
                    "teams": f"{game.home or '?'} – {game.guest or '?'}",
                    "ak": game.ak or "",
                    "color": ak_color(game.ak),
                    "team_name": game.judge_team_name or "",
                    "missing": list(missing_roles.items()),
                })

        return render_template(
            "statistik.html",
            season=f"{season_year}/{str(season_year + 1)[-2:]}",
            total_games=total_games,
            team_stats=team_stats,
            person_stats=person_stats,
            gaps=gaps,
        )

    # ------------------------------------------------------------------
    # Person management
    # ------------------------------------------------------------------

    @app.route("/personen")
    def persons():
        session = get_session()
        visible_people = [
            person for person in db.get_all_person_records(session)
            if person.account_status in (db.ACCOUNT_ACTIVE, db.ACCOUNT_INACTIVE)
        ] if g.tier == "admin" else db.get_all_persons(session)
        all_persons = [
            {
                "id": p.id,
                "name": p.name,
                "email": p.email or "" if g.tier == "admin" or p.id == g.viewer.id else "",
                "phone": p.phone or "" if g.tier == "admin" or p.id == g.viewer.id else "",
                "team_id": p.team_id,
                "team_name": p.team.name if p.team else "",
                "status": p.account_status,
                "editable": g.tier == "admin" or p.id == g.viewer.id,
            }
            for p in visible_people
        ]
        pending = [
            p for p in db.get_all_person_records(session)
            if p.account_status == db.ACCOUNT_VERIFIED
            and (g.tier == "admin" or p.desired_team_id in g.mv_team_ids)
        ]
        return render_template(
            "persons.html",
            persons=all_persons,
            teams=team_options(session),
            pending=pending,
        )

    def _form_team_id(session) -> int | None:
        raw = request.form.get("team_id")
        if not raw:
            return None
        team = session.get(db.Team, int(raw))
        return team.id if team else None

    @app.post("/personen/add")
    def add_person():
        if g.tier != "admin":
            return api_error("Keine Berechtigung.", 403)
        name = (request.form.get("name") or "").strip()
        if not name:
            flash("Bitte einen Namen angeben.", "error")
            return redirect(url_for("persons"))
        email, phone, errors = _normalized_person_contacts()
        if errors:
            flash(next(iter(errors.values())), "error")
            return redirect(url_for("persons"))
        session = get_session()
        if (
            _contact_in_use("email", email)
            or _contact_in_use("sms", phone)
        ):
            flash(
                "E-Mail-Adresse oder Telefonnummer wird bereits verwendet.",
                "error",
            )
            return redirect(url_for("persons"))
        person = db.Person(name=name, email=email, phone=phone)
        person.team_id = _form_team_id(session)
        session.add(person)
        try:
            session.commit()
        except IntegrityError:
            session.rollback()
            flash(
                "E-Mail-Adresse oder Telefonnummer wird bereits verwendet.",
                "error",
            )
            return redirect(url_for("persons"))
        flash(f"'{name}' wurde angelegt.", "ok")
        return redirect(url_for("persons"))

    @app.post("/personen/<int:person_id>/edit")
    def edit_person(person_id: int):
        session = get_session()
        person = session.get(db.Person, person_id)
        if person is None:
            flash("Person nicht gefunden.", "error")
            return redirect(url_for("persons"))
        if g.tier != "admin" and person.id != g.viewer.id:
            return api_error("Keine Berechtigung.", 403)

        name = (request.form.get("name") or "").strip()
        email, phone, errors = _normalized_person_contacts()
        if errors:
            flash(next(iter(errors.values())), "error")
            return redirect(url_for("persons"))
        if (
            _contact_in_use("email", email, person.id)
            or _contact_in_use("sms", phone, person.id)
        ):
            flash(
                "E-Mail-Adresse oder Telefonnummer wird bereits verwendet.",
                "error",
            )
            return redirect(url_for("persons"))
        team_id = _form_team_id(session) if g.tier == "admin" else None

        if name:
            person.name = name
        person.email = email
        person.phone = phone
        if team_id is not None:
            person.team_id = team_id
        try:
            session.commit()
        except IntegrityError:
            session.rollback()
            flash(
                "E-Mail-Adresse oder Telefonnummer wird bereits verwendet.",
                "error",
            )
            return redirect(url_for("persons"))
        if not person.email and not person.phone:
            flash(
                "Gespeichert. Ohne Kontaktweg kannst du dich nicht erneut anmelden.",
                "error",
            )
        else:
            flash(f"Daten von '{person.name}' gespeichert.", "ok")
        return redirect(url_for("persons"))

    @app.post("/personen/<int:person_id>/delete")
    def delete_person(person_id: int):
        if g.tier != "admin":
            return api_error("Keine Berechtigung.", 403)
        session = get_session()
        person = session.get(db.Person, person_id)
        if person is not None:
            name = person.name
            db.delete_person(session, person, g.viewer, "admin")
            flash(f"'{name}' wurde gelöscht (inkl. Diensteinträge).", "ok")
        return redirect(url_for("persons"))

    @app.post("/personen/<int:person_id>/deactivate")
    def deactivate_person(person_id: int):
        if g.tier != "admin":
            return api_error("Keine Berechtigung.", 403)
        person = get_session().get(db.Person, person_id)
        if person is None:
            return api_error("Person nicht gefunden.", 404)
        db.deactivate_person(get_session(), person, g.viewer, "admin")
        flash(f"'{person.name}' wurde deaktiviert.", "ok")
        return redirect(url_for("persons"))

    @app.post("/personen/<int:person_id>/reactivate")
    def reactivate_person(person_id: int):
        if g.tier != "admin":
            return api_error("Keine Berechtigung.", 403)
        person = get_session().get(db.Person, person_id)
        if person is None:
            return api_error("Person nicht gefunden.", 404)
        db.reactivate_person(get_session(), person)
        flash(f"'{person.name}' wurde reaktiviert.", "ok")
        return redirect(url_for("persons"))

    # ------------------------------------------------------------------
    # Team management: exactly one Mannschaftsverantwortlicher per team,
    # who must be a member of that team. Everything else about teams is
    # derived automatically.
    # ------------------------------------------------------------------

    @app.post("/api/teams/<int:team_id>/mv")
    def api_team_mv(team_id: int):
        if g.tier != "admin":
            return api_error("Keine Berechtigung.", 403)
        session = get_session()
        team = session.get(db.Team, team_id)
        if team is None:
            return api_error("Mannschaft nicht gefunden.", 404)

        data = request.get_json(silent=True) or {}
        raw_person_id = data.get("person_id")
        if not raw_person_id:
            db.set_team_mv(session, team, None)
            return jsonify(ok=True)

        person = session.get(db.Person, int(raw_person_id))
        if person is None:
            return api_error("Person nicht gefunden.", 404)
        try:
            db.set_team_mv(session, team, person)
        except ValueError as exc:
            return api_error(str(exc))
        return jsonify(ok=True)

    # ------------------------------------------------------------------
    # JSON API for inline updates
    # ------------------------------------------------------------------

    def api_error(message: str, status: int = 400):
        return jsonify(ok=False, error=message), status

    def _assignment_request():
        data = request.get_json(silent=True) or {}
        session = get_session()
        try:
            game_id = int(data.get("game_id"))
        except (TypeError, ValueError):
            return data, session, None, None, None, api_error("Spiel nicht gefunden.", 404)
        game = session.get(db.Game, game_id)
        if game is None:
            return data, session, None, None, None, api_error("Spiel nicht gefunden.", 404)
        role = data.get("role")
        if role not in db.ROLE_SLOT_COUNT:
            return data, session, None, None, None, api_error("Unbekannter Dienst.")
        try:
            slot = int(data.get("slot") or 0)
        except (TypeError, ValueError):
            return data, session, None, None, None, api_error("Ungültiger Slot.")
        if not 0 <= slot < db.ROLE_SLOT_COUNT[role]:
            return data, session, None, None, None, api_error("Ungültiger Slot.")
        if g.tier != "admin":
            game_date = parse_date(game.date)
            if game_date is not None and game_date < common.effective_today():
                return data, session, None, None, None, api_error(
                    "Vergangene Spiele können nur Admins korrigieren.", 403
                )
        return data, session, game, role, slot, None

    def _may_manage_assignment(game: db.Game, person: db.Person) -> bool:
        if g.tier == "admin":
            return True
        if person.id == g.viewer.id:
            return True
        return (
            game.team_id in g.mv_team_ids
            and person.team_id == game.team_id
            and g.tier == "mv"
        )

    def _warning_for(game: db.Game, person: db.Person) -> str | None:
        playing_team = get_session().query(db.Team).filter(
            db.Team.name == (game.ak or "")
        ).first()
        if playing_team is not None and person.team_id == playing_team.id:
            return "Person spielt selbst in diesem Spiel."
        support = db.get_support_team(get_session())
        if game.team_id and person.team_id not in (
            game.team_id,
            support.id if support else None,
        ):
            return "Person gehört nicht zum verantwortlichen Team."
        return None

    def _conflict_response(exc: db.SlotConflictError):
        current = get_session().get(db.Person, exc.current_person_id)
        return jsonify(
            ok=False,
            code="conflict",
            error=str(exc),
            current_person_id=exc.current_person_id,
            current_person_name=current.name if current else None,
        ), 409

    @app.post("/api/assignment/claim")
    def api_assignment_claim():
        data, session_db, game, role, slot, error = _assignment_request()
        if error is not None:
            return error
        if "expected_person_id" not in data or data.get("expected_person_id") is not None:
            return api_error("Ein freier Platz muss erwartet werden.")
        try:
            raw_person_id = int(data.get("person_id", g.viewer.id))
        except (TypeError, ValueError):
            return api_error("Person nicht gefunden.", 404)
        person = session_db.get(db.Person, raw_person_id)
        if person is None:
            return api_error("Person nicht gefunden.", 404)
        if not _may_manage_assignment(game, person):
            return api_error("Keine Berechtigung für diese Einteilung.", 403)
        try:
            db.claim_slot(
                session_db, game, role, slot, None, person, g.viewer, g.tier
            )
            session_db.commit()
        except db.SlotConflictError as exc:
            session_db.rollback()
            return _conflict_response(exc)
        except IntegrityError:
            session_db.rollback()
            current = session_db.query(db.Assignment).filter_by(
                game_id=game.id, role=role, slot=slot
            ).first()
            return _conflict_response(
                db.SlotConflictError(current.person_id if current else None)
            )
        except ValueError as exc:
            session_db.rollback()
            return api_error(str(exc))
        return jsonify(ok=True, warning=_warning_for(game, person))

    @app.post("/api/assignment/release")
    def api_assignment_release():
        data, session_db, game, role, slot, error = _assignment_request()
        if error is not None:
            return error
        try:
            expected_id = int(data.get("expected_person_id"))
        except (TypeError, ValueError):
            return api_error("Die erwartete Person fehlt.")
        person = session_db.get(db.Person, expected_id)
        if person is None:
            return api_error("Person nicht gefunden.", 404)
        if not _may_manage_assignment(game, person):
            return api_error("Keine Berechtigung für diese Freigabe.", 403)
        try:
            db.release_slot(
                session_db, game, role, slot, expected_id, g.viewer, g.tier
            )
            session_db.commit()
        except db.SlotConflictError as exc:
            session_db.rollback()
            return _conflict_response(exc)
        return jsonify(ok=True)

    @app.post("/api/games/<int:game_id>/team")
    def api_game_team(game_id: int):
        if g.tier != "admin":
            return api_error("Keine Berechtigung.", 403)
        session = get_session()
        game = session.get(db.Game, game_id)
        if game is None:
            return api_error("Spiel nicht gefunden.", 404)
        data = request.get_json(silent=True) or {}
        raw_team_id = data.get("team_id")
        if raw_team_id:
            team = session.get(db.Team, int(raw_team_id))
            if team is None:
                return api_error("Mannschaft nicht gefunden.", 404)
            game.team_id = team.id
            game.jteam = team.name
        else:
            game.team_id = None
            game.jteam = None
        session.commit()
        return jsonify(ok=True)

    @app.route("/audit")
    def audit():
        if g.tier != "admin":
            return api_error("Keine Berechtigung.", 403)
        query = get_session().query(db.AssignmentAudit)
        game_id = request.args.get("game_id", type=int)
        person_id = request.args.get("person_id", type=int)
        if game_id is not None:
            query = query.filter(db.AssignmentAudit.game_id == game_id)
        if person_id is not None:
            query = query.filter(or_(
                db.AssignmentAudit.actor_person_id == person_id,
                db.AssignmentAudit.affected_person_id == person_id,
            ))
        entries = query.order_by(
            db.AssignmentAudit.changed_at.desc(), db.AssignmentAudit.id.desc()
        ).all()
        games = get_session().query(db.Game).all()
        games.sort(key=db.game_sort_key)
        return render_template(
            "audit.html",
            entries=entries,
            games=games,
            persons=db.get_all_person_records(get_session()),
            selected_game=game_id,
            selected_person=person_id,
        )

    return app


app = create_app()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
