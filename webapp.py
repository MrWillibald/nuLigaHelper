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
from datetime import datetime

import common
import db
from flask import (
    Flask,
    flash,
    g,
    jsonify,
    redirect,
    render_template,
    request,
    url_for,
)

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


def _person_team(persons: list[dict], person_id: int) -> int | None:
    return next((p["team_id"] for p in persons if p["id"] == person_id), None)


def create_app() -> Flask:
    app = Flask(__name__)
    app.secret_key = os.environ.get("NULIGAHELPER_SECRET", "tus-raubling-nuligahelper")
    engine = db.make_engine(get_db_path())
    db.init_db(engine)

    @app.teardown_appcontext
    def close_session(exception):
        session = g.pop("session", None)
        if session is not None:
            session.close()

    def get_session():
        if "session" not in g:
            g.session = db.Session(engine)
        return g.session

    @app.context_processor
    def inject_globals():
        return {"active_page": request.path}

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
                    for p in sorted(t.persons, key=lambda p: p.name)
                ],
            }
            for t in db.get_all_teams(session)
        ]

    # ------------------------------------------------------------------
    # Schedule overview
    # ------------------------------------------------------------------

    def build_schedule(session, season_year: int):
        today = common.effective_today()
        games = session.query(db.Game).filter(
            db.Game.season_year == season_year
        ).all()
        games.sort(key=db.game_sort_key)

        persons = person_options(session)
        teams = team_options(session)
        support = db.get_support_team(session)
        support_id = support.id if support else None
        playing_team_by_ak = {t["name"]: t["id"] for t in teams}

        def game_view(game):
            sales = game.assignments_by_role(db.ROLE_SALE)
            responsible_team_id = game.team_id
            # the age class of the game itself identifies the team that PLAYS
            playing_team_id = playing_team_by_ak.get(game.ak or "")
            slots = []
            sale_idx = 0
            for label, role in SLOT_LABELS:
                if role == db.ROLE_SALE:
                    assignment = sales[sale_idx] if sale_idx < len(sales) else None
                    slot = sale_idx
                    sale_idx += 1
                else:
                    assignment = game.assignment_by_role(role)
                    slot = 0
                person_id = assignment.person_id if assignment is not None else None

                if person_id is None:
                    status = "none"
                elif playing_team_id and _person_team(persons, person_id) == playing_team_id:
                    status = "playing"
                elif (responsible_team_id is not None
                      and _person_team(persons, person_id) not in (responsible_team_id, support_id)):
                    status = "outside"
                else:
                    status = "ok"

                slots.append({
                    "label": label,
                    "role": role,
                    "slot": slot,
                    "person_id": person_id,
                    "status": status,
                })
            d = parse_date(game.date)
            return {
                "id": game.id,
                "nr": game.game_nr,
                "time": (game.time or "").split()[0],
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
        data = build_schedule(session, season_year)
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
        for person in db.get_all_persons(session):
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
                    "time": (game.time or "").split()[0],
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
        all_persons = [
            {
                "id": p.id,
                "name": p.name,
                "email": p.email or "",
                "phone": p.phone or "",
                "team_id": p.team_id,
            }
            for p in db.get_all_persons(session)
        ]
        return render_template(
            "persons.html",
            persons=all_persons,
            teams=team_options(session),
        )

    def _form_team_id(session) -> int | None:
        raw = request.form.get("team_id")
        if not raw:
            return None
        team = session.get(db.Team, int(raw))
        return team.id if team else None

    @app.post("/personen/add")
    def add_person():
        name = (request.form.get("name") or "").strip()
        email = (request.form.get("email") or "").strip() or None
        phone = (request.form.get("phone") or "").strip() or None
        if not name:
            flash("Bitte einen Namen angeben.", "error")
            return redirect(url_for("persons"))
        session = get_session()
        existing = session.query(db.Person).filter(db.Person.name == name).first()
        if existing is not None:
            existing.email = email or existing.email
            existing.phone = phone or existing.phone
            existing.team_id = _form_team_id(session) or existing.team_id
            flash(f"Daten von '{name}' aktualisiert.", "ok")
        else:
            person = db.get_or_create_person(session, name, email, phone)
            person.team_id = _form_team_id(session) or person.team_id
            flash(f"'{name}' wurde angelegt.", "ok")
        session.commit()
        return redirect(url_for("persons"))

    @app.post("/personen/<int:person_id>/edit")
    def edit_person(person_id: int):
        session = get_session()
        person = session.get(db.Person, person_id)
        if person is None:
            flash("Person nicht gefunden.", "error")
            return redirect(url_for("persons"))
        name = (request.form.get("name") or "").strip()
        if name:
            person.name = name
        person.email = (request.form.get("email") or "").strip() or None
        person.phone = (request.form.get("phone") or "").strip() or None
        team_id = _form_team_id(session)
        if team_id is not None:
            person.team_id = team_id
        session.commit()
        flash(f"Daten von '{person.name}' gespeichert.", "ok")
        return redirect(url_for("persons"))

    @app.post("/personen/<int:person_id>/delete")
    def delete_person(person_id: int):
        session = get_session()
        person = session.get(db.Person, person_id)
        if person is not None:
            name = person.name
            db.delete_person(session, person)
            flash(f"'{name}' wurde gelöscht (inkl. Diensteinträge).", "ok")
        return redirect(url_for("persons"))

    # ------------------------------------------------------------------
    # Team management: exactly one Mannschaftsverantwortlicher per team,
    # who must be a member of that team. Everything else about teams is
    # derived automatically.
    # ------------------------------------------------------------------

    @app.post("/api/teams/<int:team_id>/mv")
    def api_team_mv(team_id: int):
        session = get_session()
        team = session.get(db.Team, team_id)
        if team is None:
            return api_error("Mannschaft nicht gefunden.", 404)

        data = request.get_json(silent=True) or {}
        raw_person_id = data.get("person_id")
        if not raw_person_id:
            team.mv_person_id = None
            session.commit()
            return jsonify(ok=True)

        person = session.get(db.Person, int(raw_person_id))
        if person is None:
            return api_error("Person nicht gefunden.", 404)
        if person.team_id != team.id:
            return api_error(
                f"{person.name} ist kein Mitglied der Mannschaft {team.name}."
            )
        team.mv_person_id = person.id
        session.commit()
        return jsonify(ok=True)

    # ------------------------------------------------------------------
    # JSON API for inline updates
    # ------------------------------------------------------------------

    def api_error(message: str, status: int = 400):
        return jsonify(ok=False, error=message), status

    @app.post("/api/assignment")
    def api_assignment():
        data = request.get_json(silent=True) or {}
        session = get_session()

        game = session.get(db.Game, data.get("game_id"))
        if game is None:
            return api_error("Spiel nicht gefunden.", 404)

        role = data.get("role")
        if role not in db.ROLE_SLOT_COUNT:
            return api_error("Unbekannter Dienst.")
        slot_count = db.ROLE_SLOT_COUNT[role]
        try:
            slot = int(data.get("slot") or 0)
        except (TypeError, ValueError):
            return api_error("Ungültiger Slot.")
        if not 0 <= slot < slot_count:
            return api_error("Ungültiger Slot.")

        person_id = data.get("person_id")
        if person_id is not None:
            person = session.get(db.Person, int(person_id))
            if person is None:
                return api_error("Person nicht gefunden.", 404)

        current = [a.person_id for a in game.assignments_by_role(role)]
        desired = (current + [None] * slot_count)[:slot_count]
        desired[slot] = int(person_id) if person_id is not None else None

        filled = [p for p in desired if p is not None]
        if len(filled) != len(set(filled)):
            return api_error("Diese Person ist für den Dienst bereits eingetragen.")

        db.set_role_assignments(session, game, role, filled)
        return jsonify(ok=True)

    @app.post("/api/games/<int:game_id>/team")
    def api_game_team(game_id: int):
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

    return app


app = create_app()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
