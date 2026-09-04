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
os.environ.pop("NULIGAHELPER_ENV", None)
os.environ.pop("NULIGAHELPER_TRUSTED_HOSTS", None)
os.environ["NULIGAHELPER_DB"] = os.path.join(_TEST_DIR, "webapp.db")
os.environ["NULIGAHELPER_SECRET"] = "test-only-secret-not-for-production"

SEASON = 2026


def make_engine():
    """Return a fresh in-file SQLite engine with all tables created."""
    import db

    path = os.path.join(_TEST_DIR, f"{next(tempfile._get_candidate_names())}.db")
    engine = db.make_engine(path)
    db.init_db(engine)
    return engine


def Session(engine):
    """Shortcut so tests do not need to import SQLAlchemy themselves."""
    from sqlalchemy.orm import Session
    return Session(engine)


def app_db_engine():
    """Engine for the database the webapp uses (NULIGAHELPER_DB target)."""
    import db

    return db.make_engine(os.environ["NULIGAHELPER_DB"])


def sign_in(client, person_id: int, csrf_token: str = "test-csrf-token") -> str:
    """Create an authenticated test session and return its CSRF token."""
    with client.session_transaction() as browser_session:
        browser_session["person_id"] = person_id
        browser_session["csrf_token"] = csrf_token
        browser_session.permanent = True
    return csrf_token


def csrf_headers(token: str = "test-csrf-token") -> dict:
    return {"X-CSRF-Token": token}


def csrf_data(data: dict | None = None, token: str = "test-csrf-token") -> dict:
    return {**(data or {}), "csrf_token": token}


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
    games = [
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
    for game in games:
        game["source_key"] = f"test:{game['game_nr']}"
    return games


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
