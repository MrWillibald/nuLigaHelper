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
#   python manage_db.py assign 12034 Zeitnehmer "Max Mustermann"
#   python manage_db.py unassign 12034 Verkauf "Max Mustermann"
#   python manage_db.py set-jteam 12034 "Damen 1"
# ---------------------------------------------------------------

import argparse
import json
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


def cmd_add_person(args):
    session, _ = open_session(args)
    person = db.get_or_create_person(session, args.name, args.email, args.phone)
    session.commit()
    print(f"Person saved: {person.name} (email={person.email}, phone={person.phone})")


def cmd_list_persons(args):
    session, _ = open_session(args)
    for p in session.query(db.Person).order_by(db.Person.name):
        print(f"{p.name:<30} email={p.email or '-':<35} phone={p.phone or '-'}")


def cmd_list_games(args):
    session, _ = open_session(args)
    query = session.query(db.Game).filter(db.Game.season_year == args.season)
    if args.date:
        query = query.filter(db.Game.date == args.date)
    games = query.all()
    games.sort(key=db.game_sort_key)
    for g in games:
        assignments = ", ".join(f"{a.role}: {a.person.name}" for a in g.assignments) or "-"
        print(
            f"Nr.{g.game_nr:<7} {g.date or '?'} {g.time or '?':<8} {g.ak or '?':<5} "
            f"{g.home or '?'} - {g.guest or '?'}\n"
            f"         Kampfgericht-Team: {g.jteam or '-'} | {assignments}"
        )


def cmd_assign(args):
    session, _ = open_session(args)
    game = db.get_game(session, args.season, args.game_nr)
    if game is None:
        raise SystemExit(f"Game {args.game_nr} not found in season {args.season}")
    person = db.get_or_create_person(session, args.person, args.email, args.phone)
    db.assign_person(session, game, person, args.role)
    session.commit()
    print(f"{person.name} assigned to game {args.game_nr} as {args.role}")


def cmd_unassign(args):
    session, _ = open_session(args)
    game = db.get_game(session, args.season, args.game_nr)
    if game is None:
        raise SystemExit(f"Game {args.game_nr} not found in season {args.season}")
    person = db.get_or_create_person(session, args.person)
    if db.unassign_person(session, game, person, args.role):
        session.commit()
        print(f"{person.name} removed from game {args.game_nr} ({args.role})")
    else:
        print("No such assignment found")


def cmd_set_jteam(args):
    session, _ = open_session(args)
    game = db.get_game(session, args.season, args.game_nr)
    if game is None:
        raise SystemExit(f"Game {args.game_nr} not found in season {args.season}")
    game.jteam = args.team
    session.commit()
    print(f"Judge team of game {args.game_nr} set to '{args.team}'")


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

    p = sub.add_parser("add-person", help="Create or update a person")
    p.add_argument("name")
    p.add_argument("--email")
    p.add_argument("--phone")
    p.set_defaults(func=cmd_add_person)

    sub.add_parser("list-persons").set_defaults(func=cmd_list_persons)

    p = sub.add_parser("list-games", help="List games and their assignments")
    p.add_argument("--date", help="Only games on this date (dd.mm.yyyy)")
    p.set_defaults(func=cmd_list_games)

    p = sub.add_parser("assign", help="Assign a person to a game role")
    p.add_argument("game_nr", type=int)
    p.add_argument("role", choices=[
        db.ROLE_MV, db.ROLE_TIMEKEEPER, db.ROLE_SECRETARY,
        db.ROLE_SALE, db.ROLE_SECURITY, db.ROLE_CLEANING,
    ])
    p.add_argument("person")
    p.add_argument("--email", help="Contact data used when creating a new person")
    p.add_argument("--phone")
    p.set_defaults(func=cmd_assign)

    p = sub.add_parser("unassign", help="Remove an assignment")
    p.add_argument("game_nr", type=int)
    p.add_argument("role", choices=[
        db.ROLE_MV, db.ROLE_TIMEKEEPER, db.ROLE_SECRETARY,
        db.ROLE_SALE, db.ROLE_SECURITY, db.ROLE_CLEANING,
    ])
    p.add_argument("person")
    p.set_defaults(func=cmd_unassign)

    p = sub.add_parser("set-jteam", help="Set the team providing judges for a game")
    p.add_argument("game_nr", type=int)
    p.add_argument("team")
    p.set_defaults(func=cmd_set_jteam)

    return parser


if __name__ == "__main__":
    args = build_parser().parse_args()
    args.func(args)
