# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Production boundary, request hooks (host, body size, viewer,
# authentication, CSRF), response headers, health and legal pages.
# ---------------------------------------------------------------

import ipaddress
import logging
import os
import re
import secrets

from flask import (
    current_app,
    flash,
    g,
    jsonify,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from werkzeug.exceptions import RequestEntityTooLarge, SecurityError

import db
import production as production_data

from webapp.context import api_error, current_mv_team_ids, get_session, Registry, state, tier_for

routes = Registry()


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


@routes.before_request
def public_readiness():
    if request.endpoint == "health" and request.method in {"GET", "HEAD"}:
        if state().production:
            request.host
        healthy = db.database_ready(state().database_path)
        if not healthy:
            logging.getLogger("nuligahelper.operations").warning(
                "operation=health outcome=unavailable reason=database")
        return current_app.response_class("ok\n" if healthy else "unavailable\n",
                                  status=200 if healthy else 503,
                                  mimetype="text/plain", headers={"Cache-Control": "no-store"})


routes.add_url_rule("/healthz", "health", lambda: "", methods=["GET"])


@routes.before_request
def enforce_trusted_host():
    if state().production:
        # Accessing the property invokes Werkzeug's TRUSTED_HOSTS check before
        # authentication, CSRF, or endpoint code can run.
        request.host


@routes.before_request
def enforce_content_length():
    if state().production and request.content_length is not None:
        if request.content_length > MAX_PRODUCTION_BODY_BYTES:
            raise RequestEntityTooLarge()


@routes.errorhandler(413)
def request_too_large(error):
    if request.path.startswith("/api/"):
        return jsonify(ok=False, code="request_too_large", error="Anfrage zu groß."), 413
    return render_template("message.html", message="Die Anfrage ist zu groß."), 413


@routes.after_request
def production_response_headers(response):
    if state().production:
        for name, value in PRODUCTION_SECURITY_HEADERS.items():
            response.headers[name] = value
        if request.is_secure:
            response.headers["Strict-Transport-Security"] = "max-age=31536000"
        current_app.logger.info(
            "web request method=%s path=%s status=%s",
            request.method,
            str(request.url_rule) if request.url_rule else "unmatched",
            response.status_code,
        )
    return response


@routes.teardown_appcontext
def close_session(exception):
    session = g.pop("session", None)
    if session is not None:
        session.close()


@routes.before_request
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


@routes.before_request
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
        if request.method != "GET":
            flash("Deine Sitzung ist abgelaufen. Bitte melde dich erneut an.", "error")
        return redirect(url_for("login", next=request.path))
    if (
        g.viewer.account_status == db.ACCOUNT_VERIFIED
        and request.endpoint not in {"registration_status", "logout"}
    ):
        if is_json:
            return api_error("Die Registrierung ist noch nicht freigegeben.", 403)
        return redirect(url_for("registration_status"))
    return None


@routes.before_request
def csrf_protect():
    if request.method not in {"POST", "PUT", "PATCH", "DELETE"}:
        return None
    expected = session.get("csrf_token")
    supplied = request.headers.get("X-CSRF-Token") or request.form.get("csrf_token")
    if not expected or not supplied or not secrets.compare_digest(expected, supplied):
        if request.path.startswith("/api/"):
            return api_error("Ungültiges Sicherheitstoken.", 403)
        flash("Das Formular ist abgelaufen. Bitte lade die Seite neu.", "error")
        return render_template(
            "message.html", message="Bitte öffne das Formular erneut."
        ), 403
    return None


@routes.context_processor
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
        "feedback_auth_boundary": session.pop("feedback_auth_boundary", False),
        "overview_args": {"show_hidden": "1"} if g.get("show_hidden") else {},
    }


def legal_page(name):
    try:
        content = production_data.validate_legal(production_data.read_json(
            os.environ.get("NULIGAHELPER_LEGAL")))
    except production_data.ConfigurationError:
        return render_template("message.html",
            message="Diese Information ist derzeit nicht verfügbar."), 503
    return render_template("legal.html", page=content["pages"][name])


routes.add_url_rule("/impressum", "impressum", lambda: legal_page("impressum"))


routes.add_url_rule("/datenschutz", "datenschutz", lambda: legal_page("datenschutz"))


@routes.errorhandler(400)
@routes.errorhandler(403)
@routes.errorhandler(404)
@routes.errorhandler(405)
@routes.errorhandler(500)
def public_error(error):
    if isinstance(error, SecurityError):
        return "Bad Request", 400
    return render_template("message.html", message="Die Anfrage konnte nicht bearbeitet werden."), error.code


def register_routes(app) -> None:
    routes.register(app)
