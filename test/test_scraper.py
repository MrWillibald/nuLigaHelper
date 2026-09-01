"""Offline scraper tests for stable nuLiga source identity."""

import os

import helpers as h
import scraper


def _fixture(name):
    with open(os.path.join(h.PROJECT_DIR, "test", "fixtures", name), encoding="utf-8") as source:
        return source.read()


def test_duplicate_numbers_use_distinct_meeting_ids():
    games = scraper.parse_home_games(
        _fixture("nuliga_duplicate_games.html"), ["280340"]
    )
    assert [game["game_nr"] for game in games] == [555, 555]
    assert [game["date"] for game in games] == ["05.09.2026", "05.09.2026"]
    assert [game["source_key"] for game in games] == ["meeting:101", "meeting:102"]


def test_fallback_ignores_mutable_schedule_fields():
    game = {
        "game_nr": 555, "ak": "GE", "home": "TuS Raubling", "guest": "Team A",
        "date": "05.09.2026", "time": "10:00", "hall": 280340, "score": "",
    }
    shifted = {
        **game, "date": "06.09.2026", "time": "12:30", "hall": 280345,
        "score": "12:10",
    }
    assert scraper._fallback_source_key(game) == scraper._fallback_source_key(shifted)


def test_fallback_collision_reports_both_rows():
    html = _fixture("nuliga_duplicate_games.html")
    html = html.replace(
        '<a href="/meeting?foo=x&amp;meeting=101">555</a>', "555"
    ).replace(
        '<a href="/meeting?meeting=102&amp;foo=y">555</a>', "555"
    ).replace("Team B", "Team A")
    try:
        scraper.parse_home_games(html, ["280340"])
    except ValueError as exc:
        message = str(exc)
        assert "Mehrdeutige Spielidentität" in message
        assert message.count("Team A") >= 2
    else:
        raise AssertionError("indistinguishable fallback rows must be rejected")


if __name__ == "__main__":
    h.run_all(dict(globals()))
