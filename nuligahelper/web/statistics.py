# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Statistics page and admin assignment audit.
# ---------------------------------------------------------------

from datetime import datetime

from flask import (
    g,
    render_template,
    request,
)
from sqlalchemy import or_

from nuligahelper import common
from nuligahelper import db

from nuligahelper.web.context import api_error, get_session, Registry
from nuligahelper.web.formatting import ak_color, display_time, parse_date
from nuligahelper.web.views import visibility_scope

routes = Registry()


def build_statistics(session, season_year: int, today, hide_from=None) -> dict:
    """Season distribution, personal counts and outstanding staffing gaps.

    Team and person figures cover the whole season; ``hide_from`` only limits
    the outstanding-gaps list (shared home-game-day visibility).
    """
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
            label = db.role_label(a.role)
            role_counts[label] = role_counts.get(label, 0) + 1
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
        if not db.schedule_date_visible(game.date, hide_from):
            continue
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
                "missing": [(db.role_label(role), count)
                            for role, count in staffing["vacancies"].items()],
                "deficiencies": staffing["deficiencies"],
            })

    for block in season_blocks:
        if not db.schedule_date_visible(block.date, hide_from):
            continue
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

    return {
        "total_games": total_games,
        "team_stats": team_stats,
        "person_stats": person_stats,
        "gaps": gaps,
    }


@routes.route("/statistik")
def statistics():
    session = get_session()
    season_year = common.season_year_for(common.effective_today())
    today = common.effective_today()
    scope = visibility_scope()
    hide_from = scope["visibility_boundary"] if scope["visibility_restricted"] else None
    data = build_statistics(session, season_year, today, hide_from)
    return render_template(
        "statistik.html",
        **scope,
        season=f"{season_year}/{str(season_year + 1)[-2:]}",
        **data,
    )


@routes.route("/audit")
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


def register_routes(app) -> None:
    routes.register(app)
