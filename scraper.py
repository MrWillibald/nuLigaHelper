# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Scraper for the BHV/nuLiga Hallenspielplan (home games only)
# ---------------------------------------------------------------

import io
import hashlib
import logging
from urllib.parse import parse_qs, urlsplit

import pandas as pd
import requests

NULIGA_URL = (
    "https://bhv-handball.liga.nu/cgi-bin/WebObjects/nuLigaHBDE.woa/wa/clubMeetings"
)

GAME_FIELDS = [
    "source_key", "day", "date", "time", "hall", "game_nr", "ak", "home", "guest", "score",
]


def fetch_home_games(config: dict, season_year: int) -> list[dict]:
    """
    Scrape all home games of the club from the nuLiga Hallenspielplan.

    Returns a list of dicts with keys:
    source_key, day, date, time, hall, game_nr, ak, home, guest, score
    """
    logging.info("Read current home game plan from BHV Hallenspielplan website")

    season_part1 = str(season_year)
    season_part2 = str(season_year + 1)
    parameters = {
        "club": config["clubId"],
        "searchType": "1",
        "searchTimeRangeFrom": "01.09." + season_part1,
        "searchTimeRangeTo": "01.07." + season_part2,
        "onlyHomeMeetings": "false",
    }
    result = requests.post(NULIGA_URL, data=parameters)

    # pandas >= 3 rejects raw bytes (treated as file path), so decode first
    return parse_home_games(result.content.decode("utf-8"), config["hallIds"])


def _meeting_key(row) -> str | None:
    meeting_ids = set()
    for value in row:
        if not isinstance(value, tuple) or not value[1]:
            continue
        query = parse_qs(urlsplit(value[1]).query)
        for key, values in query.items():
            if key.casefold() == "meeting":
                meeting_ids.update(v for v in values if v)
    if len(meeting_ids) > 1:
        raise ValueError(f"Mehrere nuLiga-IDs in einer Spielzeile: {sorted(meeting_ids)}")
    return f"meeting:{next(iter(meeting_ids))}" if meeting_ids else None


def _fallback_source_key(game: dict) -> str:
    def normalized(value) -> str:
        return " ".join(str(value).split()).casefold()

    identity = "\x1f".join(
        normalized(game[field]) for field in ("game_nr", "ak", "home", "guest")
    )
    return f"fallback:{hashlib.sha256(identity.encode('utf-8')).hexdigest()}"


def parse_home_games(html_text: str, hall_ids: list[str]) -> list[dict]:
    """Parse a nuLiga result page without performing network access."""
    html = io.StringIO(html_text)
    table = pd.read_html(
        html, header=0, attrs={"class": "result-set"}, extract_links="body"
    )[0]
    meeting_keys = table.apply(_meeting_key, axis=1)
    table = table.map(lambda value: value[0] if isinstance(value, tuple) else value)

    # Drop obsolete columns and rename
    table.drop(table.columns[[9, 10, 11]], axis=1, inplace=True)
    table.columns = [
        "day", "date", "time", "hall", "game_nr", "ak", "home", "guest", "score",
    ]

    table["hall"] = table["hall"].astype(str)
    for column in ("day", "date"):
        table[column] = table[column].replace(r"^\s*$", pd.NA, regex=True).ffill()

    # Keep only games in own halls
    mask = table["hall"].apply(
        lambda game: any(hall in game for hall in hall_ids)
    )
    table = table[mask]

    # Drop "spielfrei" rows (no game number)
    table = table[table["game_nr"].notna()]

    # Normalize types and whitespace (NA-safe, pandas 2.x and 3.x)
    def _clean(value) -> str:
        return "" if pd.isna(value) else str(value).strip()

    games = []
    for rec in table[GAME_FIELDS[1:]].to_dict("records"):
        game = {
            "day": _clean(rec["day"]),
            "date": _clean(rec["date"]),
            "time": _clean(rec["time"]),
            "hall": int(rec["hall"]),
            "game_nr": int(rec["game_nr"]),
            "ak": _clean(rec["ak"]),
            "home": _clean(rec["home"]),
            "guest": _clean(rec["guest"]),
            "score": _clean(rec["score"]),
        }
        games.append(game)

    # to_dict() drops the source row index, so assign keys in filtered-row order.
    filtered_keys = list(meeting_keys.loc[table.index])
    for game, meeting_key in zip(games, filtered_keys):
        game["source_key"] = meeting_key or _fallback_source_key(game)

    by_key = {}
    for game in games:
        previous = by_key.get(game["source_key"])
        if previous is not None:
            raise ValueError(
                "Mehrdeutige Spielidentität "
                f"{game['source_key']}: {previous!r} / {game!r}"
            )
        by_key[game["source_key"]] = game
    logging.info(f"Current home game plan loaded: {len(games)} home games")
    return games
