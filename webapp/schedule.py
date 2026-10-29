# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Schedule page, visibility setting, candidate reads and the due-lock
# preflight before duty mutations.
# ---------------------------------------------------------------


from flask import (
    flash,
    g,
    jsonify,
    redirect,
    render_template,
    request,
    session,
    url_for,
)
from sqlalchemy import select

import common
import db

from webapp.context import _assignment_unavailable_response, api_error, get_session, Registry, state
from webapp.formatting import parse_date
from webapp.views import _ordered_person_options, block_slot_scope, block_view, build_schedule, candidates_for_scope, duty_state, game_slot_scope, may_request_block, may_request_game, person_options, visibility_scope

routes = Registry()


@routes.before_request
def persist_scheduled_duty_locks():
    # Materialize independently due locks before a requested mutation. Reads,
    # opening a confirmation, and Cancel remain strictly free of writes.
    endpoint = request.endpoint or ""
    api_write = request.method == "POST" and endpoint.startswith(("api_assignment_", "api_block_"))
    html_write = (endpoint == "assignment_form" and request.method == "POST"
                  and (request.form.get("decision") == "confirm"
                       or request.form.get("action") in {"request-replacement", "cancel-replacement", "release"})
                  and request.form.get("decision") != "cancel")
    if not (api_write or html_write) or g.viewer is None or g.tier == "guest":
        return None
    try:
        with db.Session(state().engine) as cutoff_session:
            db.lock_due_duties(cutoff_session)
            cutoff_session.commit()
    except db.AssignmentTemporarilyUnavailableError as exc:
        if request.path.startswith("/api/"):
            return _assignment_unavailable_response(exc)
        return api_error(str(exc), 503)
    get_session().rollback()
    return None


def candidate_response(people: list[dict], slots: dict, **extra):
    response = jsonify(people=people, slots=slots, **extra)
    response.headers["Cache-Control"] = "private, no-store"
    return response


@routes.get("/api/games/<int:game_id>/candidates")
def api_game_candidates(game_id: int):
    session_db = get_session()
    game = session_db.get(db.Game, game_id)
    if game is None:
        return api_error("Spiel nicht gefunden.", 404)
    assignments = {(a.role, a.slot): a for a in game.assignments}
    offered = set(db.offered_positions(game))
    person_records = {p.id: p for p in db.get_all_persons(session_db)}
    persons = person_options(session_db, person_records.values())
    assigned_ids = {a.person_id for a in game.assignments}
    allowed, used_ids = {}, set()
    for role, count in db.ROLE_SLOT_COUNT.items():
        for slot in range(count):
            assignment = assignments.get((role, slot))
            retained = (role, slot) not in offered
            if retained and assignment is None:
                continue
            occupant = assignment.person if assignment else None
            scope, team_id = game_slot_scope(game, occupant)
            state = duty_state(game, assignment)
            can_request = (not retained and not state["locked"]
                           and may_request_game(game, assignment))
            can_release = (retained and not state["locked"]
                           and may_request_game(game, assignment))
            ids = []
            if scope and not state["locked"] and not retained and (
                assignment is None or state["seeking_replacement"]
            ):
                ids = [
                    p["id"] for p in candidates_for_scope(persons, scope, team_id)
                    if p["id"] not in assigned_ids
                    and db.claim_eligibility(game, role, person_records[p["id"]], slot) is None
                ]
            # The occupant remains readable even if inactive or no longer eligible.
            if assignment is not None and (can_request or can_release):
                ids.insert(0, assignment.person_id)
            used_ids.update(ids)
            allowed[f"{role}:{slot}"] = {
                **state, "occupant_id": state["person_id"], "candidate_ids": ids,
                "can_request_replacement": can_request, "release_only": retained,
                "editable": bool(ids or can_release),
            }
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
        is_outside = (not is_playing and game.team_id is not None
                      and game.team_id not in team_ids
                      and (support is None or support.id not in team_ids))
        people.append({
            "id": person["id"], "name": person["name"], "team_label": person["team_label"],
            "sort_group": person["sort_group"], "sort_name": person["sort_name"],
            "hint": "playing" if is_playing else "outside" if is_outside else "",
        })
    return candidate_response(people, allowed, staffing=db.staffing_status(game))


@routes.get("/api/blocks/<int:block_id>/candidates")
def api_block_candidates(block_id: int):
    session_db = get_session()
    block = session_db.get(db.DayBlock, block_id)
    if block is None:
        return api_error("Tagesblock nicht gefunden.", 404)
    persons = person_options(session_db)
    assigned_ids = {a.person_id for a in block.assignments}
    allowed, used_ids = {}, set()
    for slot in range(db.block_capacity(block)):
        assignment = block.assignment_for_slot(slot)
        scope = block_slot_scope(block, assignment.person if assignment else None)
        state = duty_state(block, assignment)
        can_request = not state["locked"] and may_request_block(block, assignment)
        ids = []
        if scope and not state["locked"] and (assignment is None or state["seeking_replacement"]):
            ids = [p["id"] for p in candidates_for_scope(persons, scope, None)
                   if p["id"] not in assigned_ids]
        if assignment is not None and can_request:
            ids.insert(0, assignment.person_id)
        used_ids.update(ids)
        allowed[str(slot)] = {
            **state, "occupant_id": state["person_id"], "candidate_ids": ids,
            "can_request_replacement": can_request, "release_only": False,
            "editable": bool(ids),
        }
    people = [{"id": p["id"], "name": p["name"], "team_label": p["team_label"],
               "sort_group": 0, "sort_name": p["name"].casefold(), "hint": ""}
              for p in persons if p["id"] in used_ids]
    people.sort(key=lambda p: (p["sort_name"], p["id"]))
    return candidate_response(people, allowed, block=block_view(block))


@routes.post("/spielplan/sichtbarkeit")
def schedule_visibility():
    if g.tier != "admin":
        return api_error("Nur aktive Administratoren dürfen die Sichtbarkeit ändern.", 403)
    database_session = get_session()
    try:
        changed = db.update_schedule_visibility(
            database_session, actor_id=g.viewer.id,
            hide_from=request.form.get("hide_from"),
            expected_hide_from=request.form.get("expected_hide_from"),
        )
        database_session.commit()
    except (ValueError, db.AssignmentTemporarilyUnavailableError) as exc:
        database_session.rollback()
        session["visibility_error"] = str(exc)
        flash(str(exc), "error")
    else:
        flash("Sichtbarkeit gespeichert." if changed else "Sichtbarkeit unverändert.", "success" if changed else "info")
    return redirect(url_for("schedule", **(
        {"show_hidden": "1"} if request.form.get("show_hidden") == "1" else {}
    )), code=303)


@routes.route("/")
def schedule():
    database_session = get_session()
    season_year = common.season_year_for(common.effective_today())
    filters = {
        "date": request.args.get("date", "").strip(),
        "playing_team": request.args.get("playing_team", "").strip(),
        "responsible_team": request.args.get("responsible_team", "").strip(),
        "person": request.args.get("person", "").strip(),
    }
    scope = visibility_scope()
    data = build_schedule(database_session, season_year, filters,
                          scope["visibility_boundary"] if scope["visibility_restricted"] else None)
    return render_template(
        "schedule.html",
        **scope,
        visibility_error=session.pop("visibility_error", None),
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


def register_routes(app) -> None:
    routes.register(app)
