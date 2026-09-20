"""Offline scraper tests for canonical game numbers and SPF aggregation."""

import os

import helpers as h
import scraper


def _fixture(name):
    with open(
        os.path.join(h.PROJECT_DIR, "test", "fixtures", name), encoding="utf-8"
    ) as source:
        return source.read()


def _spf_row(number, date, time, hall=280340, age_group="SPF Mini"):
    return {
        "day": "Sa", "date": date, "time": time, "hall": hall,
        "game_nr": str(number), "ak": age_group, "home": "Team A",
        "guest": "Team B", "score": "",
    }


def test_ordinary_numbers_ignore_meeting_links_and_stay_textual():
    html = _fixture("nuliga_duplicate_games.html").replace(
        '<td><a href="/meeting?meeting=102&amp;foo=y">555</a></td>',
        '<td><a href="/meeting?meeting=102&amp;foo=y">556</a></td>',
    )
    games = scraper.parse_home_games(html, ["280340"])
    assert [game["game_nr"] for game in games] == ["555", "556"]
    assert all("source_key" not in game for game in games)


def test_duplicate_ordinary_game_number_is_rejected():
    try:
        scraper.parse_home_games(
            _fixture("nuliga_duplicate_games.html"), ["280340"]
        )
    except ValueError as exc:
        assert "Mehrdeutige Spielidentität" in str(exc)
        assert "555" in str(exc)
    else:
        raise AssertionError("ordinary duplicate game numbers must be rejected")


def test_seven_spf_matches_collapse_to_one_deterministic_event():
    rows = [
        _spf_row(number, "28.11.2026", time)
        for number, time in zip(
            (1, 3, 4, 6, 7, 9, 10),
            ("09:00", "09:40", "10:00", "10:40", "11:00", "11:40", "12:00"),
        )
    ]
    first = scraper.collapse_spielfeste(rows)
    second = scraper.collapse_spielfeste(list(reversed(rows)))
    assert first == second == [{
        "day": "Sa", "date": "28.11.2026", "time": "09:00", "hall": 280340,
        "game_nr": "SPF:2026-11-28:spf mini", "ak": "SPF Mini",
        "home": "Spielfest", "guest": "", "score": "",
    }]


def test_reused_spf_match_numbers_on_another_date_make_another_event():
    rows = [
        _spf_row(1, "28.11.2026", "09:00"),
        _spf_row(1, "12.12.2026", "09:00"),
    ]
    games = scraper.collapse_spielfeste(rows)
    assert {game["game_nr"] for game in games} == {
        "SPF:2026-11-28:spf mini", "SPF:2026-12-12:spf mini",
    }


def test_spf_groups_distinguish_full_age_group():
    rows = [
        _spf_row(1, "28.11.2026", "09:00", age_group="spf Mini"),
        _spf_row(2, "28.11.2026", "10:00", age_group="SPF E-Jugend"),
    ]
    assert len(scraper.collapse_spielfeste(rows)) == 2


def test_spf_conflicting_halls_and_invalid_values_are_rejected():
    conflicts = [
        _spf_row(1, "28.11.2026", "09:00", hall=280340),
        _spf_row(2, "28.11.2026", "09:40", hall=280345),
    ]
    for rows in (
        conflicts,
        [_spf_row(1, "invalid", "09:00")],
        [_spf_row(1, "28.11.2026", "invalid")],
    ):
        try:
            scraper.collapse_spielfeste(rows)
        except ValueError as exc:
            assert "Spielfest" in str(exc)
        else:
            raise AssertionError("malformed SPF groups must be rejected")


if __name__ == "__main__":
    h.run_all(dict(globals()))
