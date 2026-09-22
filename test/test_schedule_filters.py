"""Public schedule filtering and past-game presentation with synthetic data."""

import os
import re
import tempfile
from datetime import date
from unittest.mock import patch

import helpers as h
import db
import webapp


_previous_db = os.environ["NULIGAHELPER_DB"]
_db_path = os.path.join(h._TEST_DIR, f"schedule-{next(tempfile._get_candidate_names())}.db")
os.environ["NULIGAHELPER_DB"] = _db_path
try:
    db.initialize_db(db.make_engine(_db_path))
    app = webapp.create_app()
finally:
    os.environ["NULIGAHELPER_DB"] = _previous_db
ENGINE = db.make_engine(_db_path)
guest = app.test_client()

with h.Session(ENGINE) as session:
    games = h.sample_games()
    later = next(game for game in games if game["game_nr"] == "1001").copy()
    later.update(game_nr="9001", date="01.11.2026",
                 guest="Zweiter Gegner")
    games.append(later)
    db.sync_games(session, games, h.SEASON)
    playing = session.query(db.Team).filter_by(name="BL mD").one()
    support = db.get_support_team(session)
    first = db.Person(name="Alex Test", teams=[playing], email="first@example.test")
    second = db.Person(name="Alex Test", teams=[support], email="second@example.test")
    unassigned = db.Person(name="Private Roster", teams=[support],
                           email="private@example.test")
    session.add_all([first, second, unassigned])
    session.flush()
    first_game = session.query(db.Game).filter_by(game_nr="1001").one()
    later_game = session.query(db.Game).filter_by(game_nr="9001").one()
    first_game.team_id = support.id
    later_game.team_id = support.id
    db.assign_person(session, first_game, first, db.ROLE_TIMEKEEPER)
    db.assign_person(session, later_game, second, db.ROLE_TIMEKEEPER)
    session.commit()
    PLAYING_ID, SUPPORT_ID, FIRST_ID = playing.id, support.id, first.id


def _page(query="/", today=date(2026, 10, 15), browser=guest):
    with patch.object(webapp.common, "effective_today", return_value=today):
        response = browser.get(query)
    assert response.status_code == 200
    return response.get_data(as_text=True)


def test_01_filters_combine_and_keep_chronological_groups():
    page = _page(
        f"/?playing_team=team-{PLAYING_ID}&responsible_team=team-{SUPPORT_ID}"
        "&person=alex%20test"
    )
    assert "Nr. 1001" in page and "Nr. 9001" in page
    past_start = page.index('<details class="past-details">')
    assert page.index("Nr. 9001") < past_start < page.index("Nr. 1001"), (
        "matching upcoming and past games must stay in their own sections"
    )
    assert "Nr. 1002" not in page and "Nr. 1004" not in page
    assert "September 2026" in page and "November 2026" in page
    assert "Oktober 2026" not in page, "empty month headings must be removed"
    assert f'value="team-{PLAYING_ID}" selected' in page
    assert f'value="team-{SUPPORT_ID}" selected' in page


def test_02_unknown_team_and_unassigned_responsibility_do_not_broaden_results():
    for value in ("nonsense", "team-999999", "team-"):
        page = _page(f"/?playing_team={value}")
        assert "Keine Spiele für diese Filter gefunden." in page
        assert "Nr. 1001" not in page
    page = _page(f"/?responsible_team=team-{SUPPORT_ID}")
    assert "Nr. 1001" in page and "Nr. 9001" in page
    assert "Nr. 1002" not in page, "games without a responsible team must not match"


def test_03_name_search_matches_both_people_with_same_name_and_no_roster():
    page = _page("/?person=AlEx%20TeSt")
    assert "Nr. 1001" in page and "Nr. 9001" in page
    assert "Private Roster" not in page and "private@example.test" not in page
    assert "first@example.test" not in page and "second@example.test" not in page
    assert f'value="{FIRST_ID}"' not in page and 'data-role="' not in page
    assert "person_id" not in page and 'id="game-' not in page
    assert 'value="AlEx TeSt"' in page, "the search value must survive a GET refresh"
    assert 'name="playing_team"' in page and 'name="responsible_team"' in page
    signed_in = app.test_client()
    h.sign_in(signed_in, FIRST_ID)
    member_page = _page("/?person=Alex", browser=signed_in)
    assert "Nr. 1001" in member_page and "Nr. 9001" in member_page
    assert 'value="Alex"' in member_page


def test_04_past_section_is_collapsed_and_counts_filtered_days():
    page = _page("/?person=Alex")
    assert re.search(r"<details class=\"past-details\">\s*<summary>"
                     r"Vergangene Spieltage anzeigen \(1\)", page)
    assert 'class="day-block past"' in page
    assert "Nr. 1001" in page and "Nr. 9001" in page
    assert 'class="day-block past"' not in _page("/?person=Zweiter")
    assert "Nr. 9001" in _page("/?person=Alex")


def test_05_empty_results_clear_action_and_distinct_empty_season():
    page = _page("/?person=Private%20Roster")
    assert "Keine Spiele für diese Filter gefunden." in page
    assert '<a href="/">Filter löschen</a>' in page
    assert "Nr. 1001" not in page
    assert "Nr. 1001" in _page(), "the clear-filter URL must restore games"
    no_season = _page("/", today=date(2027, 7, 1))
    assert "Noch keine Spiele vorhanden." in no_season
    assert "Keine Spiele für diese Filter gefunden." not in no_season


def test_06_today_stays_upcoming_and_footer_links_are_shared():
    page = _page(today=date(2026, 9, 5))
    upcoming = page.split('<details class="past-details">')[0]
    assert "Nr. 1001" in upcoming, "a game on effective today is not past"
    assert 'href="/impressum"' in page
    assert 'href="/datenschutz"' in page
    login = _page("/login")
    assert 'href="/impressum"' in login and 'href="/datenschutz"' in login


if __name__ == "__main__":
    h.run_all(dict(globals()))
