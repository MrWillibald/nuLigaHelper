# ---------------------------------------------------------------
#                          nuLigaHelper – tests
# ---------------------------------------------------------------
# Shared helpers for all test modules.
#
# Importing this module makes the project root importable and points
# NULIGAHELPER_DB to a throwaway SQLite file, so the webapp module
# can be imported safely without touching real data.
# ---------------------------------------------------------------

import json
import os
import sys
import tempfile
import traceback

PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_DIR not in sys.path:
    sys.path.insert(0, PROJECT_DIR)

_TEST_DIR = tempfile.mkdtemp(prefix="nuligahelper-test-")
os.environ["NULIGAHELPER_DB"] = os.path.join(_TEST_DIR, "webapp.db")

SEASON = 2026


def make_engine():
    """Return a fresh in-file SQLite engine with all tables created."""
    from sqlalchemy import create_engine
    import db

    path = os.path.join(_TEST_DIR, f"{next(tempfile._get_candidate_names())}.db")
    engine = create_engine(f"sqlite:///{path}")
    db.init_db(engine)
    return engine


def Session(engine):
    """Shortcut so tests do not need to import SQLAlchemy themselves."""
    from sqlalchemy.orm import Session
    return Session(engine)


def app_db_engine():
    """Engine for the database the webapp uses (NULIGAHELPER_DB target)."""
    from sqlalchemy import create_engine
    return create_engine(f"sqlite:///{os.environ['NULIGAHELPER_DB']}")


def load_club_config() -> dict:
    """Load the sample club configuration (texts etc.) shipped with the repo."""
    with open(os.path.join(PROJECT_DIR, "config_template.json"), encoding="utf-8") as f:
        return json.load(f)["club"]


def sample_games() -> list[dict]:
    """
    A small synthetic home-game plan (no network access needed).

    Dates are intentionally unsorted and span the year boundary to cover
    chronological ordering; game 1001 gets assignments in other tests.
    """
    base = {"day": "Sa", "hall": 280340, "home": "TuS Raubling", "score": ""}
    return [
        {**base, "date": "03.10.2026", "time": "17:30", "game_nr": 1003,
         "ak": "BL F", "guest": "TSV Brannenburg"},
        {**base, "date": "05.09.2026", "time": "15:00", "game_nr": 1001,
         "ak": "BL mD", "guest": "SBC Traunstein"},
        {**base, "day": "So", "hall": 280345, "date": "14.03.2027", "time": "12:00",
         "game_nr": 2001, "ak": "GE", "home": "TuS Raubling", "guest": "Turnier"},
        {**base, "date": "28.09.2026", "time": "09:00", "game_nr": 1002,
         "ak": "BL M", "guest": "SV Anzing"},
        {**base, "date": "03.10.2026", "time": "10:00", "game_nr": 1004,
         "ak": "BK wD", "guest": "HT München"},
        {**base, "date": "01.11.2026", "time": "18:00", "game_nr": 1005,
         "ak": "BL mC", "guest": "TSV Übersee"},
    ]


def sync_sample_games(session) -> list[dict]:
    """Insert the sample games via the regular sync and return them."""
    import db

    games = sample_games()
    db.sync_games(session, games, SEASON)
    return games


def run_all(scope: dict) -> None:
    """Run every ``test_*`` function in *scope*; usable when executed directly."""
    failures = 0
    for name, func in sorted(scope.items()):
        if not name.startswith("test_") or not callable(func):
            continue
        try:
            func()
            print(f"  PASS {name}")
        except AssertionError as exc:
            failures += 1
            print(f"  FAIL {name}: {exc or 'assertion failed'}")
        except Exception:
            failures += 1
            print(f"  ERROR {name}")
            traceback.print_exc()
    if failures:
        print(f"\n{failures} test(s) failed")
        sys.exit(1)
    print("all tests passed")
