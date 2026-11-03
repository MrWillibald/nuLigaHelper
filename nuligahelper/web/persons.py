# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Helper management: roster page, person maintenance, memberships
# and team MV appointments.
# ---------------------------------------------------------------


from flask import (
    flash,
    g,
    jsonify,
    redirect,
    render_template,
    request,
    url_for,
)
from sqlalchemy import select
from sqlalchemy.exc import IntegrityError

from nuligahelper import common
from nuligahelper import db

from nuligahelper.web.challenges import _contact_in_use, _normalized_person_contacts
from nuligahelper.web.context import api_error, get_session, Registry
from nuligahelper.web.views import team_options

routes = Registry()


@routes.route("/personen")
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

    active_filters = []
    if name_filter:
        active_filters.append(f"Name: {name_filter}")
    selected_team = next((team for team in teams if team["id"] == team_filter), None)
    if team_filter is not None:
        active_filters.append(
            f"Mannschaft: {selected_team['name']}" if selected_team else "Unbekannte Mannschaft"
        )
    if status_filter:
        active_filters.append(f"Status: {'Aktiv' if status_filter == db.ACCOUNT_ACTIVE else 'Inaktiv'}")
    if birth_date_filter:
        active_filters.append("Geburtsdatum: Fehlt")

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
            "edit_values": g.get("person_edit_values", {})
            if g.get("error_person_id") == p.id
            and (g.tier == "admin" or p.id == g.viewer.id) else {},
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
        active_filters=active_filters,
        unknown_team_filter=team_filter is not None and selected_team is None,
        missing_birth_dates=missing_birth_dates,
        birth_date_max=common.effective_today().isoformat(),
        field_errors=g.get("person_field_errors", {}),
        error_person_id=g.get("error_person_id"),
        new_values=g.get("new_person_values", {}),
    )


def person_field_errors(errors, person_id=None):
    # Render only after the caller's authorization checks. Never flash private values.
    g.person_field_errors = errors
    g.error_person_id = person_id
    if person_id is None:
        g.new_person_values = {field: request.form.get(field, "")
                               for field in ("name", "email", "phone", "birth_date")}
    else:
        g.person_edit_values = {field: request.form[field]
                                for field in ("name", "email", "phone", "birth_date")
                                if field in request.form}
    return persons(), 400


def _contact_collision_errors(email, phone, person_id=None):
    errors = {}
    if _contact_in_use("email", email, person_id):
        errors["email"] = "E-Mail-Adresse wird bereits verwendet."
    if _contact_in_use("sms", phone, person_id):
        errors["phone"] = "Telefonnummer wird bereits verwendet."
    return errors


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


@routes.post("/personen/add")
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
        return person_field_errors(errors)
    errors = _contact_collision_errors(email, phone)
    if errors:
        return person_field_errors(errors)
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
    flash(f"'{name}' wurde angelegt.", "success")
    return redirect(url_for("persons"))


@routes.post("/personen/<int:person_id>/edit")
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
        return person_field_errors(errors, person.id)
    errors = _contact_collision_errors(email, phone, person.id)
    if errors:
        return person_field_errors(errors, person.id)
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
            "Daten gespeichert. Ohne Kontaktweg ist keine erneute Anmeldung möglich.",
            "warning",
        )
    else:
        flash(f"Daten von '{person.name}' gespeichert.", "success")
    return redirect(url_for("persons"))


@routes.post("/personen/<int:person_id>/teams")
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
    flash(f"Mannschaften von '{person.name}' gespeichert.", "success")
    return redirect(url_for("persons"))


@routes.post("/personen/<int:person_id>/teams/<int:team_id>/<action>")
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
    flash(f"Mannschaftszuordnung von '{person.name}' wurde geändert.", "success")
    return redirect(url_for("persons"))


@routes.post("/personen/<int:person_id>/delete")
def delete_person(person_id: int):
    if g.tier != "admin":
        return api_error("Keine Berechtigung.", 403)
    session = get_session()
    person = session.get(db.Person, person_id)
    if person is not None:
        name = person.name
        try:
            db.delete_person(session, person, g.viewer, "admin")
        except ValueError as exc:
            session.rollback()
            return api_error(str(exc), 409)
        flash(f"'{name}' wurde gelöscht.", "success")
    return redirect(url_for("persons"))


@routes.post("/personen/<int:person_id>/deactivate")
def deactivate_person(person_id: int):
    if g.tier != "admin":
        return api_error("Keine Berechtigung.", 403)
    person = get_session().get(db.Person, person_id)
    if person is None:
        return api_error("Person nicht gefunden.", 404)
    try:
        db.deactivate_person(get_session(), person, g.viewer, "admin")
    except ValueError as exc:
        get_session().rollback()
        return api_error(str(exc), 409)
    flash(f"'{person.name}' wurde deaktiviert.", "success")
    return redirect(url_for("persons"))


@routes.post("/personen/<int:person_id>/reactivate")
def reactivate_person(person_id: int):
    if g.tier != "admin":
        return api_error("Keine Berechtigung.", 403)
    person = get_session().get(db.Person, person_id)
    if person is None:
        return api_error("Person nicht gefunden.", 404)
    db.reactivate_person(get_session(), person)
    flash(f"'{person.name}' wurde reaktiviert.", "success")
    return redirect(url_for("persons"))


@routes.post("/api/teams/<int:team_id>/mv")
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


def register_routes(app) -> None:
    routes.register(app)
