# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Small CLI to manage persons, games and assignments until the
# web interface is available.
#
# Examples:
#   python manage_db.py init
#   python manage_db.py add-person "Max Mustermann" --email max@example.com --phone +4917012345678
#   python manage_db.py list-games
#   python manage_db.py search-person "Max"
#   python manage_db.py list-games --number 12034
#   python manage_db.py assign 42 Zeitnehmer 7
#   python manage_db.py unassign 42 Verkauf 7
#   python manage_db.py grant-admin 7
#   python manage_db.py set-jteam 42 "Damen 1"
# ---------------------------------------------------------------

import argparse
import json
import contact_validation as contacts
import os

import common
import db


def get_db_path(args) -> str:
    if args.db:
        return args.db
    config_path = os.path.join(os.path.dirname(__file__), "config.json")
    if os.path.exists(config_path):
        with open(config_path, encoding="utf-8") as f:
            club_cfg = json.load(f)["club"]
        return club_cfg.get("database", {}).get("path", db.DEFAULT_DB_PATH)
    return db.DEFAULT_DB_PATH


def open_session(args):
    engine = db.make_engine(get_db_path(args))
    db.init_db(engine)
    return db.Session(engine), engine


def cmd_init(args):
    _, engine = open_session(args)
    print(f"Database initialized at {engine.url}")


def _resolve_team(session, name: str) -> db.Team:
    """Find an existing team by name; teams cannot be created manually."""
    for team in db.get_all_teams(session):
        if team.name.lower() == name.strip().lower():
            return team
    available = ", ".join(t.name for t in db.get_all_teams(session)) or "-"
    raise SystemExit(
        f"Unknown team '{name}'. Available: {available}. "
        "Teams are derived from the scraped game plan."
    )


def _resolve_person(session, person_id: int) -> db.Person:
    person = session.get(db.Person, person_id)
    if person is None:
        raise SystemExit(f"Person ID {person_id} not found")
    return person


def _resolve_game(session, game_id: int) -> db.Game:
    game = session.get(db.Game, game_id)
    if game is None:
        raise SystemExit(f"Game ID {game_id} not found")
    return game


def cmd_add_person(args):
    session, _ = open_session(args)
    team = _resolve_team(session, args.team) if args.team else None
    try:
        email = contacts.normalize_email(args.email)
        phone = contacts.normalize_phone(args.phone)
    except contacts.ContactValidationError as exc:
        raise SystemExit(exc.message) from exc
    person = db.Person(name=args.name.strip(), email=email, phone=phone)
    session.add(person)
    if team is not None:
        person.team_id = team.id
    elif person.team_id is None:
        support = db.get_support_team(session)
        if support is not None:
            person.team_id = support.id
    session.commit()
    team_name = person.team.name if person.team else "-"
    print(
        f"Person created: ID {person.id} {person.name} "
        f"(team={team_name}, email={person.email}, phone={person.phone})"
    )


def cmd_list_teams(args):
    session, _ = open_session(args)
    for t in db.get_all_teams(session):
        suffix = " (Support)" if t.is_support else ""
        print(f"{t.name}{suffix:<12} members={len(t.persons)} games={len(t.games)}")


def cmd_list_persons(args):
    session, _ = open_session(args)
    for p in session.query(db.Person).order_by(db.Person.name):
        team_name = p.team.name if p.team else "-"
        print(f"ID {p.id:<5} {p.name:<30} team={team_name:<20} email={p.email or '-':<35} phone={p.phone or '-'}")


def cmd_contact_preflight(args):
    """Report contact normalization issues without modifying any records."""
    session, _ = open_session(args)
    report = contacts.analyze_existing_contacts(
        db.get_all_person_records(session)
    )
    if report.clean:
        print("Contact preflight: no issues found")
        return
    for item in report.invalid:
        print(
            f"INVALID person={item['person_id']} field={item['field']} "
            f"value={item['value']!r}: {item['error']}"
        )
    for item in report.changed:
        print(
            f"CHANGED person={item['person_id']} field={item['field']} "
            f"{item['value']!r} -> {item['normalized']!r}"
        )
    for item in report.collisions:
        person_ids = ",".join(
            str(person["person_id"]) for person in item["people"]
        )
        print(
            f"COLLISION field={item['field']} normalized={item['normalized']!r} "
            f"people={person_ids}"
        )


def cmd_search_person(args):
    session, _ = open_session(args)
    needle = args.name.strip().lower()
    matches = [
        person for person in db.get_all_person_records(session)
        if needle in person.name.lower()
    ]
    for person in matches:
        team_name = person.team.name if person.team else "-"
        print(f"ID {person.id:<5} {person.name} ({team_name})")


def cmd_list_games(args):
    session, _ = open_session(args)
    query = session.query(db.Game).filter(db.Game.season_year == args.season)
    if args.date:
        query = query.filter(db.Game.date == args.date)
    if args.number is not None:
        query = query.filter(db.Game.game_nr == args.number)
    games = query.all()
    games.sort(key=db.game_sort_key)
    for g in games:
        assignments = ", ".join(f"{a.role}: {a.person.name}" for a in g.assignments) or "-"
        print(
            f"ID {g.id:<5} Nr.{g.game_nr:<7} {g.date or '?'} {g.time or '?':<8} {g.ak or '?':<5} "
            f"{g.home or '?'} - {g.guest or '?'}\n"
            f"         Verantwortlich: {g.jteam or '-'} | {assignments}"
        )


def cmd_assign(args):
    session, _ = open_session(args)
    game = _resolve_game(session, args.game_id)
    person = _resolve_person(session, args.person_id)
    try:
        db.assign_person(session, game, person, args.role, actor_tier="cli")
    except ValueError as exc:
        session.rollback()
        raise SystemExit(str(exc))
    session.commit()
    print(f"{person.name} assigned to game ID {game.id} (Nr.{game.game_nr}) as {args.role}")


def cmd_unassign(args):
    session, _ = open_session(args)
    game = _resolve_game(session, args.game_id)
    person = _resolve_person(session, args.person_id)
    if db.unassign_person(session, game, person, args.role, actor_tier="cli"):
        session.commit()
        print(f"{person.name} removed from game ID {game.id} (Nr.{game.game_nr}, {args.role})")
    else:
        print("No such assignment found")


def cmd_set_jteam(args):
    session, _ = open_session(args)
    game = _resolve_game(session, args.game_id)
    if args.team:
        team = _resolve_team(session, args.team)
        game.team_id = team.id
        game.jteam = None
    else:
        game.team_id = None
    session.commit()
    team_name = game.team.name if game.team else "-"
    print(f"Responsible team of game ID {game.id} (Nr.{game.game_nr}) set to '{team_name}'")


def cmd_set_mv(args):
    session, _ = open_session(args)
    team = _resolve_team(session, args.team)
    person = _resolve_person(session, args.person_id)
    try:
        db.set_team_mv(session, team, person)
    except ValueError as exc:
        raise SystemExit(str(exc))
    print(f"{person.name} is now MV of '{team.name}'")


def cmd_grant_admin(args):
    session, _ = open_session(args)
    person = _resolve_person(session, args.person_id)
    if person.account_status != db.ACCOUNT_ACTIVE:
        raise SystemExit("Only an active person can be granted admin rights")
    person.is_admin = True
    session.commit()
    print(f"Admin rights granted to ID {person.id} ({person.name})")


def build_parser():
    parser = argparse.ArgumentParser(description="nuLigaHelper database management CLI")
    parser.add_argument("--db", help="Path to SQLite file (default from config.json)")
    parser.add_argument(
        "--season", type=int,
        default=common.season_year_for(common.effective_today()),
        help="Season start year (default: current season)",
    )
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("init").set_defaults(func=cmd_init)

    p = sub.add_parser("add-person", help="Create a person")
    p.add_argument("name")
    p.add_argument("--team", help="Existing team name (default: support team)")
    p.add_argument("--email")
    p.add_argument("--phone")
    p.set_defaults(func=cmd_add_person)

    sub.add_parser("list-teams").set_defaults(func=cmd_list_teams)

    sub.add_parser("list-persons").set_defaults(func=cmd_list_persons)

    sub.add_parser(
        "contact-preflight",
        help="Report invalid, changed and colliding contact values",
    ).set_defaults(func=cmd_contact_preflight)

    p = sub.add_parser("search-person", help="Find person IDs by display name")
    p.add_argument("name")
    p.set_defaults(func=cmd_search_person)

    p = sub.add_parser("list-games", help="List games and their assignments")
    p.add_argument("--date", help="Only games on this date (dd.mm.yyyy)")
    p.add_argument("--number", type=int, help="Only games with this display number")
    p.set_defaults(func=cmd_list_games)

    p = sub.add_parser("assign", help="Assign a person to a game role")
    p.add_argument("game_id", type=int)
    p.add_argument("role", choices=[
        db.ROLE_TIMEKEEPER, db.ROLE_SECRETARY,
        db.ROLE_SALE, db.ROLE_SECURITY, db.ROLE_CLEANING,
    ])
    p.add_argument("person_id", type=int)
    p.set_defaults(func=cmd_assign)

    p = sub.add_parser("unassign", help="Remove an assignment")
    p.add_argument("game_id", type=int)
    p.add_argument("role", choices=[
        db.ROLE_TIMEKEEPER, db.ROLE_SECRETARY,
        db.ROLE_SALE, db.ROLE_SECURITY, db.ROLE_CLEANING,
    ])
    p.add_argument("person_id", type=int)
    p.set_defaults(func=cmd_unassign)

    p = sub.add_parser("set-mv", help="Set the Mannschaftsverantwortlicher of a team")
    p.add_argument("team", help="Team name")
    p.add_argument("person_id", type=int, help="Person ID (must be a member of the team)")
    p.set_defaults(func=cmd_set_mv)

    p = sub.add_parser("grant-admin", help="Bootstrap or grant administrator rights")
    p.add_argument("person_id", type=int)
    p.set_defaults(func=cmd_grant_admin)

    p = sub.add_parser("set-jteam", help="Set the team providing judges for a game")
    p.add_argument("game_id", type=int)
    p.add_argument("team")
    p.set_defaults(func=cmd_set_jteam)

    return parser


if __name__ == "__main__":
    args = build_parser().parse_args()
    args.func(args)
