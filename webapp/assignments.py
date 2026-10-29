# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Duty mutations: game and day-block claims, releases, replacement
# requests and takeovers, the non-JavaScript duty form, cake settings and
# responsible teams.
# ---------------------------------------------------------------

import logging
import re

from flask import (
    current_app,
    flash,
    g,
    jsonify,
    redirect,
    render_template,
    request,
    url_for,
)

import common
import db

from webapp.context import _assignment_unavailable_response, _conflict_response, api_error, get_session, Registry
from webapp.formatting import parse_date
from webapp.schedule import api_block_candidates, api_game_candidates
from webapp.views import _may_manage_assignment, block_view, duty_state

routes = Registry()


def _assignment_request():
    data = (request.get_json(silent=True) or {}) if request.is_json else request.form.to_dict()
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


@routes.post("/api/assignment/claim")
def api_assignment_claim():
    data, session_db, game, role, slot, error = _assignment_request()
    if error is not None:
        return error
    if "expected_person_id" not in data or data.get("expected_person_id") not in (None, ""):
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
    assignment = next((a for a in game.assignments if a.role == role and a.slot == slot), None)
    return jsonify(ok=True, warning=_warning_for(game, person), staffing=db.staffing_status(game),
                   assignment=duty_state(game, assignment))


@routes.post("/api/assignment/release")
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
    except ValueError as exc:
        session_db.rollback()
        return api_error(str(exc))
    return jsonify(ok=True, staffing=db.staffing_status(game))


def _block_assignment_request():
    data = (request.get_json(silent=True) or {}) if request.is_json else request.form.to_dict()
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


@routes.post("/api/block-assignment/claim")
def api_block_assignment_claim():
    data, session_db, block, slot, error = _block_assignment_request()
    if error is not None:
        return error
    if "expected_person_id" not in data or data.get("expected_person_id") not in (None, ""):
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
        return jsonify(ok=True, block_id=block.id, block=block_view(block),
                       assignment=duty_state(block, block.assignment_for_slot(slot)))
    return jsonify(ok=True, block_id=block.id, block=block_view(block),
                   assignment=duty_state(block, block.assignment_for_slot(slot)))


@routes.post("/api/block-assignment/release")
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
        return jsonify(ok=True, block_id=block.id, block=block_view(block),
                       assignment=duty_state(block, block.assignment_for_slot(slot)))
    return jsonify(ok=True, block_id=block.id, block=block_view(block),
                   assignment=duty_state(block, block.assignment_for_slot(slot)))


def replacement_change(is_block: bool, takeover: bool, cancel: bool = False):
    if is_block:
        data, session_db, container, slot, error = _block_assignment_request()
        role = None
    else:
        data, session_db, container, role, slot, error = _assignment_request()
    if error is not None:
        return error
    try:
        expected_id = int(data.get("expected_person_id"))
        incoming_id = int(data.get("person_id", g.viewer.id)) if takeover else expected_id
    except (ValueError, TypeError):
        return api_error("Die erwartete oder neue Person fehlt.")
    person = session_db.get(db.Person, incoming_id)
    if person is None:
        return api_error("Person nicht gefunden.", 404)
    if ((is_block and g.tier != "admin" and incoming_id != g.viewer.id)
            or (not is_block and not _may_manage_assignment(container, person))):
        return api_error("Keine Berechtigung für diese Einteilung.", 403)
    try:
        args = [session_db, container] + ([] if is_block else [role]) + [slot, expected_id]
        if takeover or cancel:
            token = data.get("expected_request_token")
            if not isinstance(token, str) or not token:
                return api_error("Die erwartete Ersatzsuche fehlt.")
            args += [token]
            if takeover:
                args += [person]
                func = db.takeover_block_slot if is_block else db.takeover_slot
            else:
                func = db.cancel_block_replacement if is_block else db.cancel_replacement
        else:
            func = db.request_block_replacement if is_block else db.request_replacement
        assignment = func(*args, actor=g.viewer, actor_tier=g.tier)
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
    extra = {"block_id": container.id, "block": block_view(container)} if is_block else {
        "staffing": db.staffing_status(container), "warning": _warning_for(container, person) if takeover else None,
    }
    return jsonify(ok=True, assignment=duty_state(container, assignment), **extra)


@routes.post("/api/assignment/request-replacement")
def api_assignment_request_replacement():
    return replacement_change(False, False)


@routes.post("/api/assignment/cancel-replacement")
def api_assignment_cancel_replacement():
    return replacement_change(False, False, cancel=True)


@routes.post("/api/assignment/takeover")
def api_assignment_takeover():
    return replacement_change(False, True)


@routes.post("/api/block-assignment/request-replacement")
def api_block_request_replacement():
    return replacement_change(True, False)


@routes.post("/api/block-assignment/cancel-replacement")
def api_block_cancel_replacement():
    return replacement_change(True, False, cancel=True)


@routes.post("/api/block-assignment/takeover")
def api_block_takeover():
    return replacement_change(True, True)


@routes.route("/dienst", methods=["GET", "POST"])
def assignment_form():
    """HTML equivalent: selection and cancellation never mutate assignments."""
    values = request.form if request.method == "POST" else request.args
    if request.method == "POST" and values.get("decision") == "cancel":
        return redirect(url_for("schedule"))
    is_block = bool(values.get("block_id"))
    try:
        target_id = int(values.get("block_id" if is_block else "game_id"))
        slot = int(values.get("slot"))
    except (TypeError, ValueError):
        return api_error("Dienst nicht gefunden.", 404)
    container = get_session().get(db.DayBlock if is_block else db.Game, target_id)
    if container is None:
        return api_error("Dienst nicht gefunden.", 404)
    role = None if is_block else values.get("role")
    if ((is_block and not 0 <= slot < db.block_capacity(container)) or
            (not is_block and (role not in db.ROLE_SLOT_COUNT or not 0 <= slot < db.ROLE_SLOT_COUNT[role]))):
        return api_error("Ungültiger Dienst.")
    action = values.get("action", "claim")
    if request.method == "POST" and (values.get("decision") == "confirm" or action in {"request-replacement", "cancel-replacement", "release"}):
        if action in {"claim", "takeover"} and values.get("decision") != "confirm":
            return api_error("Bitte bestätige die verbindliche Einteilung.")
        handler = {
            (False, "claim"): api_assignment_claim,
            (False, "takeover"): api_assignment_takeover,
            (False, "request-replacement"): api_assignment_request_replacement,
            (False, "cancel-replacement"): api_assignment_cancel_replacement,
            (False, "release"): api_assignment_release,
            (True, "claim"): api_block_assignment_claim,
            (True, "takeover"): api_block_takeover,
            (True, "request-replacement"): api_block_request_replacement,
            (True, "cancel-replacement"): api_block_cancel_replacement,
        }.get((is_block, action))
        if handler is None:
            return api_error("Ungültige Aktion.")
        response = current_app.make_response(handler())
        if not response.is_json:
            return response
        result = response.get_json()
        if not result.get("ok"):
            return render_template("message.html", message=result.get("error", "Dienst konnte nicht gespeichert werden.")), response.status_code
        flash("Ersatz gesucht. Die bisherige Person bleibt verantwortlich." if action == "request-replacement"
              else "Ersatzsuche beendet. Die Einteilung bleibt verbindlich." if action == "cancel-replacement"
              else "Bestehender Dienst freigegeben." if action == "release" else "Verbindliche Einteilung gespeichert.", "success")
        return redirect(url_for("schedule"))
    response = api_block_candidates(target_id) if is_block else api_game_candidates(target_id)
    response = current_app.make_response(response)
    if response.status_code != 200:
        return response
    payload = response.get_json()
    key = str(slot) if is_block else f"{role}:{slot}"
    state = payload["slots"].get(key)
    if state is None:
        return api_error("Dieser Dienst wird nicht angeboten.", 403)
    people = [p for p in payload["people"] if p["id"] in state["candidate_ids"]
              and p["id"] != state["occupant_id"]]
    confirming = request.method == "POST"
    selected = next((p for p in people if str(p["id"]) == values.get("person_id")), None)
    if confirming and selected is None:
        return api_error("Diese Person kann nicht eingeteilt werden.")
    target = {"block_id" if is_block else "game_id": target_id, "slot": slot}
    if not is_block:
        target["role"] = role
    if confirming:
        # Preserve the selection page's CAS expectations, not refreshed values.
        target.update(expected_person_id=values.get("expected_person_id", ""),
                      expected_request_token=values.get("expected_request_token", ""))
    else:
        target.update(expected_person_id=state["occupant_id"] or "",
                      expected_request_token=state["replacement_token"] or "")
    return render_template("assignment_form.html", target=target, state=state, people=people,
                           selected=selected, confirming=confirming,
                           action="takeover" if state["occupant_id"] else "claim",
                           label=db.block_slot_label(container, slot) if is_block else db.position_label(role, slot),
                           date=container.date)


@routes.post("/api/blocks/<int:block_id>/cake-settings")
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
        return api_error("Die Kuchenanzahl muss eine nichtnegative ganze Zahl sein.", field="cake_quantity")
    if (
        not isinstance(delivery_time, str)
        or re.fullmatch(r"(?:[01][0-9]|2[0-3]):[0-5][0-9]", delivery_time) is None
    ):
        return api_error("Bitte eine gültige Lieferzeit im Format HH:MM angeben.", field="delivery_time")
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


@routes.post("/api/games/<int:game_id>/team")
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


def register_routes(app) -> None:
    routes.register(app)
