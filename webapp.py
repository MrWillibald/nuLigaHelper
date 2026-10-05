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
import ipaddress
import json
import re
from datetime import datetime, timedelta

import common
import production as production_data
import auth_abuse
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
from sqlalchemy import or_, select, update
from sqlalchemy.exc import IntegrityError
from werkzeug.exceptions import RequestEntityTooLarge, SecurityError
from werkzeug.middleware.proxy_fix import ProxyFix

PRODUCTION_ENV = "production"
MAX_PRODUCTION_BODY_BYTES = 1024 * 1024
PRODUCTION_SECURITY_HEADERS = {
    "Content-Security-Policy": (
        "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; "
        "img-src 'self' data:; object-src 'none'; base-uri 'self'; "
        "frame-ancestors 'none'; form-action 'self'"
    ),
    "X-Content-Type-Options": "nosniff",
    "X-Frame-Options": "DENY",
    "Referrer-Policy": "strict-origin-when-cross-origin",
    "Permissions-Policy": "camera=(), microphone=(), geolocation=()",
}

_HOST_LABEL = re.compile(r"[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\Z")


def _production_mode() -> bool:
    value = os.environ.get("NULIGAHELPER_ENV", "").strip().lower()
    if value in {"", "local", "development"}:
        return False
    if value == PRODUCTION_ENV:
        return True
    raise RuntimeError(
        "NULIGAHELPER_ENV muss 'production' sein oder für den lokalen Modus fehlen."
    )


def _valid_exact_hostname(value: str) -> bool:
    if (
        not value
        or value != value.lower()
        or len(value) > 253
        or value.startswith(".")
        or value.endswith(".")
        or ":" in value
        or "/" in value
        or "*" in value
        or "," in value
    ):
        return False
    return all(_HOST_LABEL.fullmatch(label) for label in value.split("."))


def _production_settings() -> tuple[str, str, list[str]]:
    secret = os.environ.get("NULIGAHELPER_SECRET", "")
    database = os.environ.get("NULIGAHELPER_DB", "")
    raw_hosts = os.environ.get("NULIGAHELPER_TRUSTED_HOSTS", "")
    if not secret.strip():
        raise RuntimeError("NULIGAHELPER_SECRET muss in Produktion gesetzt sein.")
    if not database or not os.path.isabs(database):
        raise RuntimeError("NULIGAHELPER_DB muss in Produktion ein absoluter Pfad sein.")
    hosts = [host.strip() for host in raw_hosts.split(",")]
    if not raw_hosts or any(not _valid_exact_hostname(host) for host in hosts):
        raise RuntimeError(
            "NULIGAHELPER_TRUSTED_HOSTS muss exakte, kleingeschriebene Hostnamen "
            "ohne Schema, Port, Pfad oder Platzhalter enthalten."
        )
    if len(set(hosts)) != len(hosts):
        raise RuntimeError("NULIGAHELPER_TRUSTED_HOSTS darf keine Duplikate enthalten.")
    return secret, database, hosts


def _plain_wsgi_error(start_response, status: str, message: str):
    body = (message + "\n").encode("utf-8")
    headers = [("Content-Type", "text/plain; charset=utf-8"),
               ("Content-Length", str(len(body)))]
    headers.extend(PRODUCTION_SECURITY_HEADERS.items())
    start_response(status, headers)
    return [body]


class ProductionProxyBoundary:
    """Validate raw Caddy metadata before one-hop ProxyFix correction."""

    def __init__(self, app):
        self.app = app

    def __call__(self, environ, start_response):
        peer = environ.get("REMOTE_ADDR", "")
        try:
            if not ipaddress.ip_address(peer).is_loopback:
                raise ValueError
        except ValueError:
            return _plain_wsgi_error(start_response, "400 Bad Request", "Ungültige Proxy-Grenze.")

        forwarded = environ.get("HTTP_FORWARDED")
        client = environ.get("HTTP_X_FORWARDED_FOR", "")
        scheme = environ.get("HTTP_X_FORWARDED_PROTO", "")
        host = environ.get("HTTP_X_FORWARDED_HOST", "")
        unexpected = (
            environ.get("HTTP_X_FORWARDED_PORT"),
            environ.get("HTTP_X_FORWARDED_PREFIX"),
        )
        try:
            if forwarded or any(unexpected):
                raise ValueError
            if any("," in value for value in (client, scheme, host)):
                raise ValueError
            ipaddress.ip_address(client)
            if scheme != "https" or not _valid_exact_hostname(host):
                raise ValueError
        except ValueError:
            return _plain_wsgi_error(start_response, "400 Bad Request", "Ungültige Proxy-Metadaten.")
        return self.app(environ, start_response)

MONATE = [
    "Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember",
]

SLOT_LABELS = [
    ("Zeitnehmer", db.ROLE_TIMEKEEPER, 0),
    ("Sekretär", db.ROLE_SECRETARY, 0),
    ("Verkauf 1", db.ROLE_SALE, 0),
    ("Verkauf 2", db.ROLE_SALE, 1),
    ("Ordnungsdienst", db.ROLE_SECURITY, 0),
    ("Kasse", db.ROLE_CASH, 0),
    ("Reinigung 1", db.ROLE_CLEANING, 0),
    ("Reinigung 2", db.ROLE_CLEANING, 1),
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
    configured_path = os.environ.get("NULIGAHELPER_DB")
    if configured_path:
        return configured_path
    return common.load_config()["club"].get(
        "database", {}
    ).get("path", db.DEFAULT_DB_PATH)


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


def _ordered_person_options(
    persons: list[dict], responsible_team_id: int | None,
    support_team_id: int | None, playing_team_id: int | None,
) -> list[dict]:
    """Return per-game assignment options grouped by suitability."""
    options = []
    for person in persons:
        team_ids = set(person["team_ids"])
        if playing_team_id is not None and playing_team_id in team_ids:
            sort_group = 4
        elif responsible_team_id is not None and responsible_team_id in team_ids:
            sort_group = 1
        elif support_team_id is not None and support_team_id in team_ids:
            sort_group = 2
        else:
            sort_group = 3
        options.append({
            **person,
            "sort_group": sort_group,
            "sort_name": person["name"].casefold(),
        })
    return sorted(
        options,
        key=lambda person: (
            person["sort_group"], person["sort_name"], person["id"]
        ),
    )


def create_app() -> Flask:
    app = Flask(__name__)
    production = _production_mode()
    if production:
        from production_logging import configure
        configure()
        secret_key, database_path, trusted_hosts = _production_settings()
    else:
        secret_key = os.environ.get("NULIGAHELPER_SECRET")
        database_path = get_db_path()
        trusted_hosts = None
    if not secret_key:
        raise RuntimeError("NULIGAHELPER_SECRET muss gesetzt sein.")
    app.secret_key = secret_key
    app.config.update(
        PERMANENT_SESSION_LIFETIME=timedelta(hours=1),
        SESSION_REFRESH_EACH_REQUEST=True,
        SESSION_COOKIE_SAMESITE="Lax",
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_PATH="/",
        SESSION_COOKIE_DOMAIN=None,
        SESSION_COOKIE_SECURE=production,
        MAX_CONTENT_LENGTH=(MAX_PRODUCTION_BODY_BYTES if production else None),
        TRUSTED_HOSTS=trusted_hosts,
        NULIGAHELPER_PRODUCTION=production,
    )
    engine = db.make_engine(database_path)
    db.verify_db(engine)
    serializer = URLSafeTimedSerializer(secret_key, salt="nuligahelper-auth")
    raw_abuse_config = os.environ.get("NULIGAHELPER_AUTH_ABUSE_CONFIG")
    if raw_abuse_config:
        try:
            abuse_section = json.loads(raw_abuse_config)
        except json.JSONDecodeError as exc:
            raise RuntimeError("NULIGAHELPER_AUTH_ABUSE_CONFIG must be valid JSON") from exc
    elif production or os.environ.get("NULIGAHELPER_DB"):
        abuse_section = {}
    else:
        abuse_section = common.load_config().get("club", {}).get("auth_abuse", {})
    try:
        abuse_config = auth_abuse.load_config(abuse_section)
    except ValueError as exc:
        raise RuntimeError(f"Invalid auth_abuse configuration: {exc}") from exc
    abuse_service = auth_abuse.Service(engine, abuse_config, secret_key)
    abuse_service.cleanup()
    app.extensions["nuligahelper_auth_abuse"] = abuse_service

    if production:
        app.logger.setLevel(logging.INFO)
        corrected_app = ProxyFix(
            app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_port=0, x_prefix=0
        )
        app.wsgi_app = ProductionProxyBoundary(corrected_app)

    @app.before_request
    def public_readiness():
        if request.endpoint == "health" and request.method in {"GET", "HEAD"}:
            if production:
                request.host
            healthy = db.database_ready(database_path)
            if not healthy:
                logging.getLogger("nuligahelper.operations").warning(
                    "operation=health outcome=unavailable reason=database")
            return app.response_class("ok\n" if healthy else "unavailable\n",
                                      status=200 if healthy else 503,
                                      mimetype="text/plain", headers={"Cache-Control": "no-store"})

    app.add_url_rule("/healthz", "health", lambda: "", methods=["GET"])

    @app.before_request
    def enforce_trusted_host():
        if production:
            # Accessing the property invokes Werkzeug's TRUSTED_HOSTS check before
            # authentication, CSRF, or endpoint code can run.
            request.host

    @app.before_request
    def enforce_content_length():
        if production and request.content_length is not None:
            if request.content_length > MAX_PRODUCTION_BODY_BYTES:
                raise RequestEntityTooLarge()

    @app.errorhandler(413)
    def request_too_large(error):
        if request.path.startswith("/api/"):
            return jsonify(ok=False, code="request_too_large", error="Anfrage zu groß."), 413
        return render_template("message.html", message="Die Anfrage ist zu groß."), 413

    @app.after_request
    def production_response_headers(response):
        if production:
            for name, value in PRODUCTION_SECURITY_HEADERS.items():
                response.headers[name] = value
            if request.is_secure:
                response.headers["Strict-Transport-Security"] = "max-age=31536000"
            app.logger.info(
                "web request method=%s path=%s status=%s",
                request.method,
                str(request.url_rule) if request.url_rule else "unmatched",
                response.status_code,
            )
        return response

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
        if request.endpoint in {"impressum", "datenschutz"}:
            g.tier, g.mv_team_ids = "guest", set()
            return
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
        "health",
        "impressum",
        "datenschutz",
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
            "membership_label": db.membership_label,
            "person_label": db.person_label,
        }

    def legal_page(name):
        try:
            content = production_data.validate_legal(production_data.read_json(
                os.environ.get("NULIGAHELPER_LEGAL")))
        except production_data.ConfigurationError:
            return render_template("message.html",
                message="Diese Information ist derzeit nicht verfügbar."), 503
        return render_template("legal.html", page=content["pages"][name])

    app.add_url_rule("/impressum", "impressum", lambda: legal_page("impressum"))
    app.add_url_rule("/datenschutz", "datenschutz", lambda: legal_page("datenschutz"))

    @app.errorhandler(400)
    @app.errorhandler(403)
    @app.errorhandler(404)
    @app.errorhandler(405)
    @app.errorhandler(500)
    def public_error(error):
        if isinstance(error, SecurityError):
            return "Bad Request", 400
        return render_template("message.html", message="Die Anfrage konnte nicht bearbeitet werden."), error.code

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
            abuse_config,
        )

    def _reserve(rules: list[auth_abuse.Rule]) -> auth_abuse.Decision:
        client = _client_subject()
        if client is None:
            auth_abuse.LOGGER.error(
                "auth_abuse_storage_error action=client_attribution channel=any reason=invalid_proxy_metadata"
            )
            return auth_abuse.Decision(False, (), True)
        return abuse_service.reserve([
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
        return serializer.dumps({
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
            payload = serializer.loads(signed_challenge, max_age=15 * 60)
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
                abuse_service.digest("contact", channel, destination),
                abuse_service.digest("person", channel, str(person.id)),
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
            abuse_service.digest("contact", channel, destination),
            abuse_service.digest("person", channel, f"dummy:{destination}"),
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
        except Exception as exc:
            auth_abuse.LOGGER.error(
                "auth_delivery_failed action=account_message channel=preferred reason=%s",
                type(exc).__name__,
            )
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
        except Exception as exc:
            auth_abuse.LOGGER.error(
                "auth_delivery_failed action=account_message channel=%s reason=%s",
                channel, type(exc).__name__,
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
            f"{code} ist dein Code für die {action} bei nuLigaHelper des TuS Raubling Handball. "
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
            request_message=message,
            country_codes=COUNTRY_CODES,
            birth_date_max=common.effective_today().isoformat(),
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
        _notify_registration_approvers(person)
        return render_template(
            "message.html",
            message="Kontakt bestätigt. Die Freigabe steht noch aus.",
        )

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
                "Neue Registrierung",
                (
                    f"Hallo {approver.name},\n\n"
                    f"{person.name} hat den Kontakt bestätigt und wartet auf "
                    f"Freigabe für: {team_label}.\n\n"
                    "Bitte prüfe die Registrierung unter \"Helfer verwalten\"."
                ),
                (
                    f"Hallo {approver.name}, neue Registrierung von {person.name} "
                    f"für {team_label}. Bitte unter \"Helfer verwalten\" prüfen."
                ),
            )

    def _notify_registration_approved(person: db.Person) -> None:
        greeting = (
            "herzlich willkommen beim nuLigaHelper des TuS Raubling Handball!"
        )
        _safe_account_message(
            person,
            "Registrierung freigegeben",
            (
                f"Hallo {person.name},\n\n{greeting}\n\n"
                "Deine Registrierung wurde freigegeben. Du kannst dich jetzt "
                "anmelden und offene Dienste im Heimspielplan übernehmen."
            ),
            (
                f"Hallo {person.name}, {greeting} Deine Registrierung wurde "
                "freigegeben. Melde dich an und übernimm offene Dienste im "
                "Heimspielplan."
            ),
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
        if g.tier != "admin":
            return api_error("Keine Berechtigung.", 403)
        if decision == "approve":
            db.approve_person(get_session(), person)
        elif decision == "reject":
            person.account_status = db.ACCOUNT_REJECTED
            person.rejected_at = datetime.now()
        else:
            return api_error("Ungültige Entscheidung.")
        get_session().commit()
        if decision == "approve":
            _notify_registration_approved(person)
        return redirect(url_for("persons"))

    def person_options(session, records=None) -> list[dict]:
        return [
            {
                "id": p.id,
                "name": p.name,
                "team_ids": db.membership_team_ids(p),
                "team_label": db.membership_label(p),
            }
            for p in (db.get_all_persons(session) if records is None else records)
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

    def game_slot_scope(game: db.Game, occupant: db.Person | None) -> tuple[str | None, int | None]:
        """Return the viewer's editable scope and managed team for one slot."""
        if g.tier == "admin":
            return "all", None
        if g.viewer is None or g.tier not in {"member", "mv"}:
            return None, None
        game_date = parse_date(game.date)
        if game_date is not None and game_date < common.effective_today():
            return None, None
        mv_team_id = game.team_id if game.team_id in g.mv_team_ids else None
        if mv_team_id is not None and (
            occupant is None or db.has_team(occupant, mv_team_id)
        ):
            return "team", mv_team_id
        if occupant is None or occupant.id == g.viewer.id:
            return "self", None
        return None, None

    def block_slot_scope(block: db.DayBlock, occupant: db.Person | None) -> str | None:
        if g.tier == "admin":
            return "all"
        if g.viewer is None or g.tier not in {"member", "mv"}:
            return None
        block_date = parse_date(block.date)
        if block_date is not None and block_date < common.effective_today():
            return None
        return "self" if occupant is None or occupant.id == g.viewer.id else None

    def candidates_for_scope(persons: list[dict], scope: str, team_id: int | None):
        if scope == "all":
            return persons
        if scope == "team":
            return [
                p for p in persons
                if p["id"] == g.viewer.id or team_id in p["team_ids"]
            ]
        return [p for p in persons if p["id"] == g.viewer.id]

    def candidate_response(people: list[dict], slots: dict, **extra):
        response = jsonify(people=people, slots=slots, **extra)
        response.headers["Cache-Control"] = "private, no-store"
        return response

    @app.get("/api/games/<int:game_id>/candidates")
    def api_game_candidates(game_id: int):
        session_db = get_session()
        game = session_db.get(db.Game, game_id)
        if game is None:
            return api_error("Spiel nicht gefunden.", 404)
        assignments = {(a.role, a.slot): a for a in game.assignments}
        offered = set(db.offered_positions(game))
        slot_scopes = {}
        for role, count in db.ROLE_SLOT_COUNT.items():
            for slot in range(count):
                assignment = assignments.get((role, slot))
                if (role, slot) not in offered and assignment is None:
                    continue
                occupant = assignment.person if assignment else None
                scope, team_id = game_slot_scope(game, occupant)
                if scope is not None:
                    slot_scopes[f"{role}:{slot}"] = (scope, team_id, occupant.id if occupant else None)
        if not slot_scopes:
            return candidate_response([], {})

        # Load eligibility inputs once; never serialize private birth dates.
        person_records = {p.id: p for p in db.get_all_persons(session_db)}
        persons = person_options(session_db, person_records.values())
        assigned_ids = {a.person_id for a in game.assignments}
        allowed = {}
        used_ids = set()
        for key, (scope, team_id, occupant_id) in slot_scopes.items():
            role, raw_slot = key.rsplit(":", 1)
            retained = (role, int(raw_slot)) not in offered
            ids = [
                p["id"] for p in candidates_for_scope(persons, scope, team_id)
                if p["id"] == occupant_id or (
                    not retained
                    and
                    p["id"] not in assigned_ids
                    and db.claim_eligibility(game, role, person_records[p["id"]], int(raw_slot)) is None
                )
            ]
            used_ids.update(ids)
            allowed[key] = {"candidate_ids": ids, "occupant_id": occupant_id,
                            "release_only": retained}
        support = db.get_support_team(session_db)
        playing = session_db.scalar(select(db.Team).where(db.Team.name == (game.ak or "")))
        ordered = _ordered_person_options(
            [p for p in persons if p["id"] in used_ids],
            game.team_id, support.id if support else None, playing.id if playing else None,
        )
        people = []
        for person in ordered:
            team_ids = person["team_ids"]
            is_playing = playing is not None and playing.id in team_ids
            is_outside = (
                not is_playing and game.team_id is not None
                and game.team_id not in team_ids
                and (support is None or support.id not in team_ids)
            )
            people.append({
                "id": person["id"], "name": person["name"],
                "team_label": person["team_label"],
                "sort_group": person["sort_group"], "sort_name": person["sort_name"],
                "hint": "playing" if is_playing else "outside" if is_outside else "",
            })
        return candidate_response(people, allowed, staffing=db.staffing_status(game))

    @app.get("/api/blocks/<int:block_id>/candidates")
    def api_block_candidates(block_id: int):
        session_db = get_session()
        block = session_db.get(db.DayBlock, block_id)
        if block is None:
            return api_error("Tagesblock nicht gefunden.", 404)
        slot_scopes = {}
        for slot in range(db.block_capacity(block)):
            assignment = block.assignment_for_slot(slot)
            occupant = assignment.person if assignment else None
            scope = block_slot_scope(block, occupant)
            if scope is not None:
                slot_scopes[str(slot)] = (scope, occupant.id if occupant else None)
        if not slot_scopes:
            return candidate_response([], {}, **(
                {"block": block_view(block)}
                if block.phase == db.BLOCK_CAKE_DELIVERY else {}
            ))

        persons = person_options(session_db)
        assigned_ids = {assignment.person_id for assignment in block.assignments}
        allowed = {}
        used_ids = set()
        for key, (scope, occupant_id) in slot_scopes.items():
            ids = [
                p["id"] for p in candidates_for_scope(persons, scope, None)
                if p["id"] == occupant_id or p["id"] not in assigned_ids
            ]
            used_ids.update(ids)
            allowed[key] = {"candidate_ids": ids, "occupant_id": occupant_id}
        people = [
            {
                "id": p["id"], "name": p["name"], "team_label": p["team_label"],
                "sort_group": 0, "sort_name": p["name"].casefold(), "hint": "",
            }
            for p in persons if p["id"] in used_ids
        ]
        people.sort(key=lambda p: (p["sort_name"], p["id"]))
        return candidate_response(people, allowed, **(
            {"block": block_view(block)}
            if block.phase == db.BLOCK_CAKE_DELIVERY else {}
        ))

    def progress_data(filled: int, total: int) -> dict[str, int]:
        return {
            "filled": filled,
            "total": total,
            "percent": (filled * 100 + total // 2) // total if total else 0,
        }

    def block_view(block, full_date_games=None):
        calculated = db.calculated_block_time(block, full_date_games)
        home_date = parse_date(block.date)
        adjacent_date = (
            calculated.strftime("%d.%m.%Y")
            if calculated and calculated.date() != home_date else ""
        )
        slots = []
        for slot in range(db.block_capacity(block)):
            assignment = block.assignment_for_slot(slot)
            occupant = assignment.person if assignment else None
            slots.append({
                "label": db.block_slot_label(block, slot),
                "slot": slot,
                "person_id": occupant.id if occupant else None,
                "person_name": occupant.name if occupant else "",
                "person_team_label": db.membership_label(occupant) if occupant else "",
                "editable": block_slot_scope(block, occupant) is not None,
            })
        return {
            "id": block.id, "phase": block.phase, "label": block.label,
            "time": calculated.strftime("%H:%M") if calculated else "",
            "time_date": adjacent_date,
            "delivery_time": block.delivery_time,
            "cake_quantity": block.cake_quantity,
            "configured": db.block_is_configured(block),
            "slots": slots,
            "progress": progress_data(
                sum(slot["person_id"] is not None for slot in slots), len(slots)
            ),
            "past": bool(home_date and home_date < common.effective_today()),
        }

    def build_schedule(
        session, season_year: int, filters: dict[str, str] | None = None,
    ):
        filters = filters or {}
        today = common.effective_today()
        games = session.query(db.Game).filter(
            db.Game.season_year == season_year
        ).all()
        games.sort(key=db.game_sort_key)
        total_games = len(games)
        date_options = []
        seen_dates = set()
        for game in games:
            if game.date and game.date not in seen_dates:
                date_options.append({"date": game.date, "day": game.day or ""})
                seen_dates.add(game.date)

        teams = [
            {"id": t.id, "name": t.name, "is_support": t.is_support}
            for t in db.get_all_teams(session)
        ]
        support = db.get_support_team(session)
        support_id = support.id if support else None
        playing_team_by_ak = {t["name"]: t["id"] for t in teams}

        def selected_team(value: str) -> int | None:
            if not value:
                return None
            if not value.startswith("team-") or not value[5:].isdigit():
                return -1
            team_id = int(value[5:])
            return team_id if any(t["id"] == team_id for t in teams) else -1

        playing_filter = selected_team(filters.get("playing_team", ""))
        responsible_filter = selected_team(filters.get("responsible_team", ""))
        person_filter = " ".join(filters.get("person", "").split()).casefold()
        date_filter = filters.get("date", "")
        all_games = games

        def game_view(game):
            assignments = {(a.role, a.slot): a for a in game.assignments}
            required = set(db.required_positions(game))
            responsible_team_id = game.team_id
            # the age class of the game itself identifies the team that PLAYS
            playing_team_id = playing_team_by_ak.get(game.ak or "")
            slots = []
            for label, role, slot in SLOT_LABELS:
                assignment = assignments.get((role, slot))
                if (role, slot) not in required and assignment is None:
                    continue
                person_id = assignment.person_id if assignment is not None else None
                occupant = assignment.person if assignment is not None else None

                if person_id is None:
                    status = "none"
                elif playing_team_id and db.has_team(occupant, playing_team_id):
                    status = "playing"
                elif (responsible_team_id is not None
                      and not db.has_team(occupant, responsible_team_id)
                      and not db.has_team(occupant, support_id)):
                    status = "outside"
                else:
                    status = "ok"

                scope, _ = game_slot_scope(game, occupant)
                slots.append({
                    "label": label,
                    "role": role,
                    "slot": slot,
                    "required": (role, slot) in required,
                    "retained": (role, slot) not in required,
                    "person_id": person_id,
                    "person_name": occupant.name if occupant else "",
                    "person_team_label": db.membership_label(occupant) if occupant else "",
                    "status": status,
                    "editable": scope is not None,
                })
            d = parse_date(game.date)
            staffing = db.staffing_status(game)
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
                "spielfest": db.is_spielfest(game),
                "hall": game.hall,
                "team_id": responsible_team_id,
                "playing_team_id": playing_team_id,
                "slots": slots,
                "progress": progress_data(
                    staffing["required_filled"], staffing["required_total"]
                ),
                "classification_unresolved": staffing["classification_unresolved"],
                "past": bool(d and d < today),
                "eligibility_deficiencies": staffing["deficiencies"] if not d or d >= today else [],
            }

        day_groups = []
        dates = list(dict.fromkeys(game.date for game in all_games))
        for game_date in dates:
            if date_filter and game_date != date_filter:
                continue
            full_date_games = [game for game in all_games if game.date == game_date]
            team_games = [
                game for game in full_date_games
                if (playing_filter is None
                    or playing_team_by_ak.get(game.ak or "") == playing_filter)
                and (responsible_filter is None or game.team_id == responsible_filter)
            ]
            if not team_games:
                continue
            visible_games = team_games
            if person_filter:
                visible_games = [
                    game for game in team_games
                    if any(
                        person_filter in " ".join(a.person.name.split()).casefold()
                        for a in game.assignments
                    )
                ]
            blocks = [] if responsible_filter is not None else db.get_day_blocks(
                session, season_year, game_date or ""
            )
            if person_filter:
                blocks = [
                    block for block in blocks
                    if any(
                        person_filter in " ".join(a.person.name.split()).casefold()
                        for a in block.assignments
                    )
                ]
            if not visible_games and not blocks:
                continue
            views = [game_view(game) for game in visible_games]
            d = parse_date(game_date)
            month_label = f"{MONATE[d.month - 1]} {d.year}" if d else "Ohne Datum"
            block_views = {block.phase: block_view(block, full_date_games) for block in blocks}
            day_groups.append({
                "type": "day",
                "month": month_label,
                "day": full_date_games[0].day or "",
                "date": game_date,
                "past": bool(d and d < today),
                "games": views,
                "preparation": block_views.get(db.BLOCK_PREPARATION),
                "cake_delivery": block_views.get(db.BLOCK_CAKE_DELIVERY),
                "cleanup": block_views.get(db.BLOCK_CLEANUP),
            })

        def with_month_headers(day_list):
            result = []
            current_month = None
            for day_group in day_list:
                if day_group["month"] != current_month:
                    result.append({"type": "month", "label": day_group["month"]})
                    current_month = day_group["month"]
                result.append(day_group)
            return result

        upcoming = [dg for dg in day_groups if not dg["past"]]
        past = [dg for dg in day_groups if dg["past"]]

        return {
            "upcoming": with_month_headers(upcoming),
            "past": with_month_headers(past),
            "teams": teams,
            "support_id": support_id,
            "total_games": total_games,
            "date_options": date_options,
        }

    @app.route("/")
    def schedule():
        session = get_session()
        season_year = common.season_year_for(common.effective_today())
        filters = {
            "date": request.args.get("date", "").strip(),
            "playing_team": request.args.get("playing_team", "").strip(),
            "responsible_team": request.args.get("responsible_team", "").strip(),
            "person": request.args.get("person", "").strip(),
        }
        data = build_schedule(session, season_year, filters)
        return render_template(
            "schedule.html",
            upcoming=data["upcoming"],
            past=data["past"],
            teams=data["teams"],
            support_id=data["support_id"],
            total_games=data["total_games"],
            date_options=data["date_options"],
            selected_date_is_past=bool(
                (selected_date := parse_date(filters["date"]))
                and selected_date < common.effective_today()
            ),
            filters=filters,
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
        season_blocks = session.query(db.DayBlock).filter(
            db.DayBlock.season_year == season_year
        ).all()
        season_block_ids = {block.id for block in season_blocks}
        person_stats = []
        for person in db.get_all_person_records(session):
            if person.account_status not in (db.ACCOUNT_ACTIVE, db.ACCOUNT_INACTIVE):
                continue
            assignments = [
                a for a in person.assignments if a.game_id in season_game_ids
            ]
            block_assignments = [
                a for a in person.block_assignments if a.block_id in season_block_ids
            ]
            if not assignments and not block_assignments:
                continue
            role_counts: dict[str, int] = {}
            for a in assignments:
                role_counts[a.role] = role_counts.get(a.role, 0) + 1
            for a in block_assignments:
                label = a.block.label
                role_counts[label] = role_counts.get(label, 0) + 1
            person_stats.append({
                "name": person.name,
                "team_label": db.membership_label(person),
                "jobs": len(assignments) + len(block_assignments),
                "roles": sorted(role_counts.items(), key=lambda kv: (-kv[1], kv[0])),
            })
        person_stats.sort(key=lambda p: (-p["jobs"], p["name"]))

        gaps = []
        for game in games:
            d = parse_date(game.date)
            if d is not None and d < today:
                continue
            staffing = db.staffing_status(game)
            if not staffing["complete"]:
                gaps.append({
                    "kind": "game",
                    "nr": game.game_nr,
                    "date": game.date,
                    "time": display_time(game.time),
                    "teams": db.game_display_name(game),
                    "ak": game.ak or "",
                    "color": ak_color(game.ak),
                    "team_name": game.judge_team_name or "",
                    "missing": list(staffing["vacancies"].items()),
                    "deficiencies": staffing["deficiencies"],
                })

        for block in season_blocks:
            d = parse_date(block.date)
            if d is None or d < today:
                continue
            occupied = {assignment.slot for assignment in block.assignments}
            missing_count = sum(
                1 for slot in range(db.block_capacity(block)) if slot not in occupied
            )
            setup_needed = not db.block_is_configured(block)
            if missing_count or setup_needed:
                calculated = db.calculated_block_time(block)
                gaps.append({
                    "kind": "block",
                    "nr": "",
                    "date": block.date,
                    "time": calculated.strftime("%H:%M") if calculated else "Zeit offen",
                    "teams": block.label,
                    "ak": "Tagesdienst",
                    "color": "#0a1d4e",
                    "team_name": "–",
                    "missing": [(block.label, missing_count)] if missing_count else [],
                    "setup_needed": setup_needed,
                })
        gaps.sort(key=lambda gap: (
            parse_date(gap["date"]) or datetime.max.date(), gap["time"], gap["teams"]
        ))

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
        teams = team_options(session)
        all_records = db.get_all_person_records(session) if g.tier == "admin" else []
        missing_birth_dates = sum(p.birth_date is None for p in all_records)
        birth_date_filter = (
            "missing" if g.tier == "admin"
            and request.args.get("birth_date") == "missing" else ""
        )
        visible_people = [
            person for person in all_records
            if birth_date_filter or person.account_status in (db.ACCOUNT_ACTIVE, db.ACCOUNT_INACTIVE)
        ] if g.tier == "admin" else db.get_all_persons(session)
        if birth_date_filter:
            visible_people = [p for p in visible_people if p.birth_date is None]

        name_filter = (request.args.get("name") or "").strip()
        team_filter = request.args.get("team_id", type=int)
        status_filter = (request.args.get("status") or "").strip()
        if g.tier != "admin" or status_filter not in {
            db.ACCOUNT_ACTIVE,
            db.ACCOUNT_INACTIVE,
        }:
            status_filter = ""
        if name_filter:
            folded_name = name_filter.casefold()
            visible_people = [
                person for person in visible_people
                if folded_name in person.name.casefold()
            ]
        if team_filter is not None:
            visible_people = [
                person for person in visible_people
                if db.has_team(person, team_filter)
            ]
        if status_filter:
            visible_people = [
                person for person in visible_people
                if person.account_status == status_filter
            ]

        all_persons = [
            {
                "id": p.id,
                "name": p.name,
                "email": p.email or "" if g.tier == "admin" or p.id == g.viewer.id else "",
                "phone": p.phone or "" if g.tier == "admin" or p.id == g.viewer.id else "",
                "team_ids": db.membership_team_ids(p),
                "teams": [{"id": t.id, "name": t.name} for t in db.membership_teams(p)],
                "team_label": db.membership_label(p),
                "status": p.account_status,
                "editable": g.tier == "admin" or p.id == g.viewer.id,
                **({"birth_date": p.birth_date.isoformat() if p.birth_date else ""}
                   if g.tier == "admin" or p.id == g.viewer.id else {}),
                "mv_actions": [
                    {
                        "id": team_id,
                        "name": next(t["name"] for t in teams if t["id"] == team_id),
                        "member": db.has_team(p, team_id),
                        "locked": next(
                            t["mv_person_id"] for t in teams if t["id"] == team_id
                        ) == p.id,
                    }
                    for team_id in sorted(g.mv_team_ids)
                ] if g.tier == "mv" else [],
            }
            for p in visible_people
        ]
        pending = [
            p for p in db.get_all_person_records(session)
            if p.account_status == db.ACCOUNT_VERIFIED
        ] if g.tier == "admin" else []
        creation_teams = teams if g.tier == "admin" else [
            team for team in teams if team["id"] in g.mv_team_ids
        ]
        return render_template(
            "persons.html",
            persons=all_persons,
            teams=teams,
            creation_teams=creation_teams,
            pending=pending,
            name_filter=name_filter,
            team_filter=team_filter,
            status_filter=status_filter,
            birth_date_filter=birth_date_filter,
            missing_birth_dates=missing_birth_dates,
            birth_date_max=common.effective_today().isoformat(),
        )

    def _form_team_id(session) -> int | None:
        raw = request.form.get("team_id")
        if not raw:
            return None
        team = session.get(db.Team, int(raw))
        return team.id if team else None

    def _form_team_ids() -> list[int]:
        raw_values = request.form.getlist("team_ids")
        if any(not value.isdigit() for value in raw_values):
            raise ValueError("Ungültige Mannschaftsauswahl.")
        return [int(value) for value in raw_values]

    @app.post("/personen/add")
    def add_person():
        if g.tier not in {"admin", "mv"}:
            return api_error("Keine Berechtigung.", 403)
        name = (request.form.get("name") or "").strip()
        if not name:
            flash("Bitte einen Namen angeben.", "error")
            return redirect(url_for("persons"))
        session = get_session()
        if g.tier == "mv":
            try:
                team_id = _form_team_id(session)
            except (TypeError, ValueError):
                team_id = None
            if team_id is None:
                flash("Bitte eine gültige Mannschaft auswählen.", "error")
                return redirect(url_for("persons"))
            if team_id not in g.mv_team_ids:
                return api_error("Keine Berechtigung.", 403)
            team_ids = [team_id]
        else:
            try:
                team_ids = _form_team_ids()
            except ValueError as exc:
                flash(str(exc), "error")
                return redirect(url_for("persons"))
        email, phone, errors = _normalized_person_contacts()
        birth_date = None
        try:
            birth_date = db.validate_birth_date(request.form.get("birth_date"))
        except ValueError as exc:
            errors["birth_date"] = str(exc)
        if errors:
            flash(next(iter(errors.values())), "error")
            return redirect(url_for("persons"))
        if (
            _contact_in_use("email", email)
            or _contact_in_use("sms", phone)
        ):
            flash(
                "E-Mail-Adresse oder Telefonnummer wird bereits verwendet.",
                "error",
            )
            return redirect(url_for("persons"))
        person = db.Person(name=name, email=email, phone=phone, birth_date=birth_date)
        session.add(person)
        try:
            session.flush()
            db.replace_person_teams(session, person, team_ids)
            session.commit()
        except (IntegrityError, ValueError) as exc:
            session.rollback()
            flash(
                str(exc) if isinstance(exc, ValueError) else
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
        birth_date = person.birth_date
        if "birth_date" in request.form:
            raw_birth_date = (request.form.get("birth_date") or "").strip()
            if raw_birth_date or person.birth_date is not None:
                try:
                    birth_date = db.validate_birth_date(raw_birth_date)
                except ValueError as exc:
                    errors["birth_date"] = str(exc)
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
        if name:
            person.name = name
        person.email = email
        person.phone = phone
        person.birth_date = birth_date
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

    @app.post("/personen/<int:person_id>/teams")
    def edit_person_teams(person_id: int):
        if g.tier not in {"admin", "mv"}:
            return api_error("Keine Berechtigung.", 403)
        session = get_session()
        person = session.get(db.Person, person_id)
        if person is None:
            return api_error("Person nicht gefunden.", 404)
        try:
            selected_ids = set(_form_team_ids())
        except ValueError as exc:
            flash(str(exc), "error")
            return redirect(url_for("persons"))

        try:
            if g.tier == "admin":
                db.replace_person_teams(session, person, selected_ids)
            else:
                if not selected_ids.issubset(g.mv_team_ids):
                    return api_error("Keine Berechtigung.", 403)
                managed_teams = list(session.scalars(
                    select(db.Team).where(db.Team.id.in_(g.mv_team_ids))
                ))
                if {team.id for team in managed_teams} != g.mv_team_ids:
                    return api_error("Mannschaft nicht gefunden.", 404)
                for team in managed_teams:
                    should_be_member = team.id in selected_ids
                    if db.has_team(person, team) != should_be_member:
                        db.change_managed_team_membership(
                            session, g.viewer, person, team, add=should_be_member
                        )
            session.commit()
        except ValueError as exc:
            session.rollback()
            return api_error(str(exc), 403)
        flash(f"Mannschaften von '{person.name}' gespeichert.", "ok")
        return redirect(url_for("persons"))

    @app.post("/personen/<int:person_id>/teams/<int:team_id>/<action>")
    def change_roster_membership(person_id: int, team_id: int, action: str):
        if g.tier != "mv" or action not in {"add", "remove"}:
            return api_error("Keine Berechtigung.", 403)
        session = get_session()
        person = session.get(db.Person, person_id)
        team = session.get(db.Team, team_id)
        if person is None or team is None:
            return api_error("Person oder Mannschaft nicht gefunden.", 404)
        try:
            db.change_managed_team_membership(
                session, g.viewer, person, team, add=action == "add"
            )
            session.commit()
        except ValueError as exc:
            session.rollback()
            return api_error(str(exc), 403)
        flash(f"Mannschaftszuordnung von '{person.name}' wurde geändert.", "ok")
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
            and db.has_team(person, game.team_id)
            and g.tier == "mv"
        )

    def _warning_for(game: db.Game, person: db.Person) -> str | None:
        playing_team = get_session().query(db.Team).filter(
            db.Team.name == (game.ak or "")
        ).first()
        if playing_team is not None and db.has_team(person, playing_team.id):
            return "Person spielt selbst in diesem Spiel."
        support = db.get_support_team(get_session())
        if (game.team_id and not db.has_team(person, game.team_id)
                and not db.has_team(person, support.id if support else None)):
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

    def _assignment_unavailable_response(
        exc: db.AssignmentTemporarilyUnavailableError,
    ):
        return jsonify(
            ok=False,
            code="temporarily_unavailable",
            error=str(exc),
        ), 503

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
        except db.AssignmentTemporarilyUnavailableError as exc:
            session_db.rollback()
            logging.warning(
                "Assignment claim temporarily unavailable for game=%s role=%s "
                "slot=%s: %s",
                game.id,
                role,
                slot,
                exc,
            )
            return _assignment_unavailable_response(exc)
        except ValueError as exc:
            session_db.rollback()
            return api_error(str(exc))
        return jsonify(ok=True, warning=_warning_for(game, person), staffing=db.staffing_status(game))

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
        except db.AssignmentTemporarilyUnavailableError as exc:
            session_db.rollback()
            logging.warning(
                "Assignment release temporarily unavailable for game=%s role=%s "
                "slot=%s: %s",
                game.id,
                role,
                slot,
                exc,
            )
            return _assignment_unavailable_response(exc)
        return jsonify(ok=True, staffing=db.staffing_status(game))

    def _block_assignment_request():
        data = request.get_json(silent=True) or {}
        session_db = get_session()
        try:
            block_id = int(data.get("block_id"))
            slot = int(data.get("slot"))
        except (TypeError, ValueError):
            return data, session_db, None, None, api_error(
                "Tagesblock nicht gefunden.", 404
            )
        block = session_db.get(db.DayBlock, block_id)
        if block is None:
            return data, session_db, None, None, api_error(
                "Tagesblock nicht gefunden.", 404
            )
        if not 0 <= slot < db.block_capacity(block):
            return data, session_db, None, None, api_error("Ungültiger Slot.")
        block_date = parse_date(block.date)
        if (
            g.tier != "admin"
            and block_date is not None
            and block_date < common.effective_today()
        ):
            return data, session_db, None, None, api_error(
                "Vergangene Tagesblöcke können nur Admins korrigieren.", 403
            )
        return data, session_db, block, slot, None

    @app.post("/api/block-assignment/claim")
    def api_block_assignment_claim():
        data, session_db, block, slot, error = _block_assignment_request()
        if error is not None:
            return error
        if "expected_person_id" not in data or data.get("expected_person_id") is not None:
            return api_error("Ein freier Platz muss erwartet werden.")
        try:
            person_id = int(data.get("person_id", g.viewer.id))
        except (TypeError, ValueError):
            return api_error("Person nicht gefunden.", 404)
        person = session_db.get(db.Person, person_id)
        if person is None:
            return api_error("Person nicht gefunden.", 404)
        if g.tier != "admin" and person.id != g.viewer.id:
            return api_error("Keine Berechtigung für diese Einteilung.", 403)
        try:
            db.claim_block_slot(
                session_db, block, slot, None, person, g.viewer, g.tier
            )
            session_db.commit()
        except db.SlotConflictError as exc:
            session_db.rollback()
            return _conflict_response(exc)
        except db.AssignmentTemporarilyUnavailableError as exc:
            session_db.rollback()
            return _assignment_unavailable_response(exc)
        except ValueError as exc:
            session_db.rollback()
            return api_error(str(exc))
        if block.phase == db.BLOCK_CAKE_DELIVERY:
            return jsonify(ok=True, block_id=block.id, block=block_view(block))
        return jsonify(ok=True, block_id=block.id)

    @app.post("/api/block-assignment/release")
    def api_block_assignment_release():
        data, session_db, block, slot, error = _block_assignment_request()
        if error is not None:
            return error
        try:
            expected_id = int(data.get("expected_person_id"))
        except (TypeError, ValueError):
            return api_error("Die erwartete Person fehlt.")
        if g.tier != "admin" and expected_id != g.viewer.id:
            return api_error("Keine Berechtigung für diese Freigabe.", 403)
        try:
            db.release_block_slot(
                session_db, block, slot, expected_id, g.viewer, g.tier
            )
            session_db.commit()
        except db.SlotConflictError as exc:
            session_db.rollback()
            return _conflict_response(exc)
        except db.AssignmentTemporarilyUnavailableError as exc:
            session_db.rollback()
            return _assignment_unavailable_response(exc)
        except ValueError as exc:
            session_db.rollback()
            return api_error(str(exc))
        if block.phase == db.BLOCK_CAKE_DELIVERY:
            return jsonify(ok=True, block_id=block.id, block=block_view(block))
        return jsonify(ok=True, block_id=block.id)

    @app.post("/api/blocks/<int:block_id>/cake-settings")
    def api_cake_settings(block_id: int):
        if g.tier != "admin":
            return api_error("Nur Admins können die Kuchenlieferung konfigurieren.", 403)
        session_db = get_session()
        block = session_db.get(db.DayBlock, block_id)
        if block is None or block.phase != db.BLOCK_CAKE_DELIVERY:
            return api_error("Kuchenblock nicht gefunden.", 404)
        data = request.get_json(silent=True)
        if not isinstance(data, dict):
            return api_error("Ungültige Kucheneinstellungen.")
        quantity = data.get("cake_quantity")
        delivery_time = data.get("delivery_time")
        if type(quantity) is not int or quantity < 0:
            return api_error("Die Kuchenanzahl muss eine nichtnegative ganze Zahl sein.")
        if (
            not isinstance(delivery_time, str)
            or re.fullmatch(r"(?:[01][0-9]|2[0-3]):[0-5][0-9]", delivery_time) is None
        ):
            return api_error("Bitte eine gültige Lieferzeit im Format HH:MM angeben.")
        if not {"expected_delivery_time", "expected_cake_quantity"} <= data.keys():
            return api_error("Die erwarteten gespeicherten Einstellungen fehlen.")
        expected_time = data["expected_delivery_time"]
        expected_quantity = data["expected_cake_quantity"]
        if (
            (expected_time is not None and not isinstance(expected_time, str))
            or (expected_quantity is not None and type(expected_quantity) is not int)
        ):
            return api_error("Ungültige erwartete Einstellungen.")
        try:
            db.configure_cake_block(
                session_db, block, delivery_time, quantity,
                expected_delivery_time=expected_time,
                expected_cake_quantity=expected_quantity,
            )
            session_db.commit()
        except db.CakeConfigurationConflictError as exc:
            session_db.rollback()
            return jsonify(
                ok=False, code="stale_cake_settings",
                error="Die Kucheneinstellungen wurden inzwischen geändert. Bitte prüfe die gespeicherten Werte.",
                current_delivery_time=exc.current_delivery_time,
                current_cake_quantity=exc.current_cake_quantity,
                block=block_view(block),
            ), 409
        except db.AssignmentTemporarilyUnavailableError as exc:
            session_db.rollback()
            return _assignment_unavailable_response(exc)
        except ValueError as exc:
            session_db.rollback()
            return api_error(str(exc))
        return jsonify(ok=True, block=block_view(block))

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
        target = (request.args.get("target") or "").strip()
        person_id = request.args.get("person_id", type=int)
        if game_id is not None:
            query = query.filter(db.AssignmentAudit.game_id == game_id)
        elif target.startswith("game-") and target[5:].isdigit():
            query = query.filter(db.AssignmentAudit.game_id == int(target[5:]))
        elif target.startswith("block-") and target[6:].isdigit():
            query = query.filter(db.AssignmentAudit.block_id == int(target[6:]))
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
        blocks = get_session().query(db.DayBlock).all()
        blocks.sort(key=lambda block: (
            parse_date(block.date) or datetime.max.date(),
            db.BLOCK_PHASES.index(block.phase) if block.phase in db.BLOCK_PHASES else 99,
            block.id,
        ))
        return render_template(
            "audit.html",
            entries=entries,
            games=games,
            blocks=blocks,
            persons=db.get_all_person_records(get_session()),
            selected_game=game_id,
            selected_target=target or (f"game-{game_id}" if game_id is not None else ""),
            selected_person=person_id,
        )

    return app


app = create_app()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
