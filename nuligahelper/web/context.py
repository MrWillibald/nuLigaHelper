# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Shared request context: per-app state, database session, viewer
# tier and common JSON/HTML error responses.
# ---------------------------------------------------------------

from dataclasses import dataclass

from flask import (
    current_app,
    flash,
    g,
    jsonify,
    render_template,
    request,
)

from nuligahelper import db



@dataclass(frozen=True)
class AppState:
    """Per-app resources created once by create_app()."""

    engine: object
    serializer: object
    abuse_service: object
    abuse_config: object
    database_path: str
    production: bool


def state() -> AppState:
    return current_app.extensions["nuligahelper"]


class Registry:
    """Collect routes and request hooks of one module for later registration.

    Endpoint names stay the plain function names (no Blueprint prefix), so
    templates keep using url_for("schedule") etc. Hooks register in the order
    they are declared; create_app() registers modules in a fixed order.
    """

    def __init__(self):
        self._entries = []

    def route(self, rule, **options):
        def decorator(view):
            self._entries.append(("route", rule, view, options))
            return view
        return decorator

    def get(self, rule, **options):
        return self.route(rule, methods=["GET"], **options)

    def post(self, rule, **options):
        return self.route(rule, methods=["POST"], **options)

    def add_url_rule(self, rule, endpoint, view, **options):
        self._entries.append(("url_rule", rule, (endpoint, view), options))

    def _hook(self, kind):
        def decorator(function):
            self._entries.append((kind, None, function, {}))
            return function
        return decorator

    def before_request(self, function):
        return self._hook("before_request")(function)

    def after_request(self, function):
        return self._hook("after_request")(function)

    def teardown_appcontext(self, function):
        return self._hook("teardown_appcontext")(function)

    def context_processor(self, function):
        return self._hook("context_processor")(function)

    def errorhandler(self, code):
        def decorator(function):
            self._entries.append(("errorhandler", code, function, {}))
            return function
        return decorator

    def register(self, app) -> None:
        for kind, rule, target, options in self._entries:
            if kind == "route":
                app.add_url_rule(rule, target.__name__, target, **options)
            elif kind == "url_rule":
                endpoint, view = target
                app.add_url_rule(rule, endpoint, view, **options)
            elif kind == "errorhandler":
                app.register_error_handler(rule, target)
            else:
                getattr(app, kind)(target)


def get_session():
    if "session" not in g:
        g.session = db.Session(state().engine)
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


def api_error(message: str, status: int = 400, *, field: str | None = None):
    if not request.path.startswith("/api/"):
        flash(message, "error")
        return render_template("message.html", message="Bitte prüfe die Anfrage und deine Berechtigung."), status
    return jsonify(ok=False, error=message,
                   **({"field_errors": {field: message}} if field else {})), status


def _conflict_response(exc: db.SlotConflictError):
    current = get_session().get(db.Person, exc.current_person_id) if exc.current_person_id is not None else None
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
