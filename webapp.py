# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Web interface: home game schedule with inline task assignment
# and helper (person) management. Designed to visually integrate
# with www.handball-raubling.de
#
# Run locally:  ./venv/bin/python webapp.py   (http://<pi-ip>:8080)
# Optional env: NULIGAHELPER_DB=/path/to/nuliga_helper.db
# ---------------------------------------------------------------

import os
from datetime import datetime

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

import common
import db

MONATE = [
    "Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember",
]

SLOT_LABELS = [
    ("MV Kampfgericht", db.ROLE_MV),
    ("Zeitnehmer", db.ROLE_TIMEKEEPER),
    ("Sekretär", db.ROLE_SECRETARY),
    ("Verkauf 1", db.ROLE_SALE),
    ("Verkauf 2", db.ROLE_SALE),
    ("Ordnungsdienst", db.ROLE_SECURITY),
    ("Reinigung", db.ROLE_CLEANING),
]

ROLE_SLOT_COUNT = {
    db.ROLE_MV: 1,
    db.ROLE_TIMEKEEPER: 1,
    db.ROLE_SECRETARY: 1,
    db.ROLE_SALE: 2,
    db.ROLE_SECURITY: 1,
    db.ROLE_CLEANING: 1,
}


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

    # ------------------------------------------------------------------
    # Schedule overview
    # ------------------------------------------------------------------

    def build_schedule(session, season_year: int):
        today = common.effective_today()
        games = session.query(db.Game).filter(
            db.Game.season_year == season_year
        ).all()
        games.sort(key=db.game_sort_key)

        persons = db.get_all_persons(session)
        person_opts = [{"id": p.id, "name": p.name} for p in persons]

        def game_view(game):
            sales = game.assignments_by_role(db.ROLE_SALE)
            slots = []
            sale_idx = 0
            for label, role in SLOT_LABELS:
                if role == db.ROLE_SALE:
                    person_id = sales[sale_idx].person_id if sale_idx < len(sales) else None
                    slot = sale_idx
                    sale_idx += 1
                else:
                    assignment = game.assignment_by_role(role)
                    person_id = assignment.person_id if assignment is not None else None
                    slot = 0
                slots.append(
                    {"label": label, "role": role, "slot": slot, "person_id": person_id}
                )
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
                "jteam": game.jteam or "",
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
            "persons": person_opts,
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
            season=f"{season_year}/{str(season_year + 1)[-2:]}",
        )

    # ------------------------------------------------------------------
    # Person management
    # ------------------------------------------------------------------

    @app.route("/personen")
    def persons():
        session = get_session()
        all_persons = [
            {"id": p.id, "name": p.name, "email": p.email or "", "phone": p.phone or ""}
            for p in db.get_all_persons(session)
        ]
        return render_template("persons.html", persons=all_persons)

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
            flash(f"Daten von '{name}' aktualisiert.", "ok")
        else:
            db.get_or_create_person(session, name, email, phone)
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
        if role not in ROLE_SLOT_COUNT:
            return api_error("Unbekannter Dienst.")
        slot_count = ROLE_SLOT_COUNT[role]
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

    @app.post("/api/games/<int:game_id>/jteam")
    def api_jteam(game_id: int):
        session = get_session()
        game = session.get(db.Game, game_id)
        if game is None:
            return api_error("Spiel nicht gefunden.", 404)
        data = request.get_json(silent=True) or {}
        game.jteam = (data.get("value") or "").strip()[:120] or None
        session.commit()
        return jsonify(ok=True)

    return app


app = create_app()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
