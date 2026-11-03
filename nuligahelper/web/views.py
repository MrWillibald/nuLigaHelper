# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# View-model builders: schedule cards, day blocks, duty state and
# candidate scopes. Templates only render what these return.
# ---------------------------------------------------------------


from flask import (
    g,
    request,
    url_for,
)

from nuligahelper import common
from nuligahelper import db

from nuligahelper.web.context import get_session
from nuligahelper.web.formatting import ak_color, display_time, MONATE, parse_date, SLOT_LABELS


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


def duty_state(container, assignment):
    locked = db.duty_is_locked(get_session(), container, assignment)
    return {
        "person_id": assignment.person_id if assignment else None,
        "person_name": assignment.person.name if assignment else "",
        "person_team_label": db.membership_label(assignment.person) if assignment else "",
        "locked": bool(assignment and locked),
        "late_claim_locked": bool(locked),
        "replacement_token": assignment.replacement_token if assignment else None,
        "seeking_replacement": bool(assignment and assignment.replacement_token and not locked),
        "can_request_replacement": bool(assignment and not locked and (
            (isinstance(container, db.Game) and (assignment.role, assignment.slot) in db.offered_positions(container)
             and may_request_game(container, assignment))
            or (isinstance(container, db.DayBlock) and may_request_block(container, assignment))
        )),
        "release_only": bool(assignment and isinstance(container, db.Game)
                             and (assignment.role, assignment.slot) not in db.offered_positions(container)),
    }


def game_slot_scope(game: db.Game, occupant: db.Person | None) -> tuple[str | None, int | None]:
    """Scope for the incoming helper, independently of outgoing ownership."""
    if g.tier == "admin":
        return "all", None
    if g.viewer is None or g.tier not in {"member", "mv"}:
        return None, None
    game_date = parse_date(game.date)
    if game_date is not None and game_date < common.effective_today():
        return None, None
    mv_team_id = game.team_id if game.team_id in g.mv_team_ids else None
    if mv_team_id is not None:
        return "team", mv_team_id
    return "self", None


def block_slot_scope(block: db.DayBlock, occupant: db.Person | None) -> str | None:
    if g.tier == "admin":
        return "all"
    if g.viewer is None or g.tier not in {"member", "mv"}:
        return None
    block_date = parse_date(block.date)
    if block_date is not None and block_date < common.effective_today():
        return None
    return "self"


def may_request_game(game, assignment):
    return bool(assignment and game_slot_scope(game, assignment.person)[0]
                and _may_manage_assignment(game, assignment.person))


def may_request_block(block, assignment):
    return bool(assignment and block_slot_scope(block, assignment.person)
                and (g.tier == "admin" or assignment.person_id == g.viewer.id))


def candidates_for_scope(persons: list[dict], scope: str, team_id: int | None):
    if scope == "all":
        return persons
    if scope == "team":
        return [
            p for p in persons
            if p["id"] == g.viewer.id or team_id in p["team_ids"]
        ]
    return [p for p in persons if p["id"] == g.viewer.id]


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
            "description": db.TASK_DESCRIPTIONS[block.phase],
            "slot": slot,
            "person_id": occupant.id if occupant else None,
            "person_name": occupant.name if occupant else "",
            "person_team_label": db.membership_label(occupant) if occupant else "",
            **duty_state(block, assignment),
            "can_request_replacement": bool(not db.duty_is_locked(get_session(), block, assignment)
                                            and may_request_block(block, assignment)),
            "editable": bool(block_slot_scope(block, occupant) is not None
                             and (assignment is None or not db.duty_is_locked(get_session(), block, assignment))
                             ),
        })
    progress = progress_data(
        sum(slot["person_id"] is not None for slot in slots), len(slots)
    )
    missing = progress["total"] - progress["filled"]
    return {
        "id": block.id, "phase": block.phase, "label": block.label,
        "description": db.TASK_DESCRIPTIONS[block.phase],
        "time": calculated.strftime("%H:%M") if calculated else "",
        "time_date": adjacent_date,
        "delivery_time": block.delivery_time,
        "cake_quantity": block.cake_quantity,
        "configured": db.block_is_configured(block),
        "slots": slots,
        "progress": progress,
        "vacancies": {block.label: missing} if missing else {},
        "past": bool(home_date and home_date < common.effective_today()),
    }


def build_schedule(
    session, season_year: int, filters: dict[str, str] | None = None,
    hide_from=None,
):
    filters = filters or {}
    today = common.effective_today()
    games = session.query(db.Game).filter(
        db.Game.season_year == season_year
    ).all()
    games.sort(key=db.game_sort_key)
    total_games = len(games)
    games = [game for game in games if db.schedule_date_visible(game.date, hide_from)]
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
                "description": db.TASK_DESCRIPTIONS[role],
                "role": role,
                "slot": slot,
                "required": (role, slot) in required,
                "retained": (role, slot) not in required,
                "person_id": person_id,
                "person_name": occupant.name if occupant else "",
                "person_team_label": db.membership_label(occupant) if occupant else "",
                "status": status,
                **duty_state(game, assignment),
                "can_request_replacement": bool((role, slot) in required
                                                and not db.duty_is_locked(get_session(), game, assignment)
                                                and may_request_game(game, assignment)),
                "editable": bool(scope is not None
                                 and (assignment is None or not db.duty_is_locked(get_session(), game, assignment))
                                 and ((role, slot) in required or may_request_game(game, assignment))),
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
            "vacancies": {role: staffing["vacancies"][role]
                          for role in db.ROLE_SLOT_COUNT if staffing["vacancies"].get(role)},
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


def visibility_scope():
    boundary = db.schedule_visibility_boundary(get_session())
    g.show_hidden = g.tier == "admin" and request.args.get("show_hidden") == "1"
    toggle_args = ({key: request.args[key] for key in
                    ("date", "playing_team", "responsible_team", "person")
                    if key in request.args} if request.endpoint == "schedule" else {})
    if not g.show_hidden:
        toggle_args["show_hidden"] = "1"
    return {
        "visibility_boundary": boundary,
        "visibility_restricted": boundary is not None and not g.show_hidden,
        "show_hidden": g.show_hidden,
        "visibility_toggle_url": url_for(request.endpoint, **toggle_args),
    }


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
