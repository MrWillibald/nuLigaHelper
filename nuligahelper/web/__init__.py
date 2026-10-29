# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Web interface: home game schedule with inline task assignment,
# helper/team management and statistics. Designed to visually
# integrate with www.handball-raubling.de
#
# create_app() builds the Flask app; each topic module registers
# its routes and hooks through register_routes():
#   security     production boundary, auth guard, CSRF, health, legal
#   auth         login and registration (helpers in challenges)
#   schedule     schedule page, visibility, candidates, due-lock preflight
#   statistics   statistics and audit
#   persons      helper and team management
#   assignments  duty mutations and the non-JavaScript duty form
#   views        view-model builders shared by the pages
#
# Run locally:  ./venv/bin/python -m nuligahelper.web   (http://<pi-ip>:8080)
# Optional env: NULIGAHELPER_DB=/path/to/nuliga_helper.db
# ---------------------------------------------------------------

import json
import logging
import os
from datetime import timedelta
from pathlib import Path

from flask import Flask
from itsdangerous import URLSafeTimedSerializer
from werkzeug.middleware.proxy_fix import ProxyFix

from nuligahelper import auth_abuse
from nuligahelper import common
from nuligahelper import db
from nuligahelper.jobs import messages
from nuligahelper.jobs import notifier
from nuligahelper.web import assignments, auth, persons, schedule, security, statistics
from nuligahelper.web.context import AppState
from nuligahelper.web.security import (
    MAX_PRODUCTION_BODY_BYTES,
    PRODUCTION_SECURITY_HEADERS,
    ProductionProxyBoundary,
    _production_mode,
    _production_settings,
)
from nuligahelper.web.views import _ordered_person_options

# Kept importable from webapp for tests, fixtures and tools.
__all__ = [
    "MAX_PRODUCTION_BODY_BYTES", "PRODUCTION_SECURITY_HEADERS", "_ordered_person_options",
    "app", "common", "create_app", "db", "get_db_path", "messages", "notifier",
]

# Templates and static files live inside the web package.
PACKAGE_DIR = Path(__file__).resolve().parent

# Registration order is part of the contract: the security hooks (readiness,
# trusted host, body size, viewer, authentication, CSRF) run before the
# schedule module's due-lock preflight.
MODULES = (security, auth, schedule, statistics, persons, assignments)


def get_db_path() -> str:
    configured_path = os.environ.get("NULIGAHELPER_DB")
    if configured_path:
        return configured_path
    return common.load_config()["club"].get(
        "database", {}
    ).get("path", db.DEFAULT_DB_PATH)


def create_app() -> Flask:
    common.refuse_overrides_in_production()
    app = Flask(
        __name__,
        template_folder=str(PACKAGE_DIR / "templates"),
        static_folder=str(PACKAGE_DIR / "static"),
    )
    production = _production_mode()
    if production:
        from nuligahelper.ops.journal_logging import configure
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
    app.extensions["nuligahelper"] = AppState(
        engine=engine,
        serializer=serializer,
        abuse_service=abuse_service,
        abuse_config=abuse_config,
        database_path=database_path,
        production=production,
    )

    if production:
        app.logger.setLevel(logging.INFO)
        corrected_app = ProxyFix(
            app.wsgi_app, x_for=1, x_proto=1, x_host=1, x_port=0, x_prefix=0
        )
        app.wsgi_app = ProductionProxyBoundary(corrected_app)

    for module in MODULES:
        module.register_routes(app)
    return app


app = create_app()
