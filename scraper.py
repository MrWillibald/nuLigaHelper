# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Scraper for the BHV/nuLiga Hallenspielplan (home games only)
# ---------------------------------------------------------------

import io
import logging

import pandas as pd
import requests

NULIGA_URL = (
    "https://bhv-handball.liga.nu/cgi-bin/WebObjects/nuLigaHBDE.woa/wa/clubMeetings"
)

GAME_FIELDS = ["day", "date", "time", "hall", "game_nr", "ak", "home", "guest", "score"]


def fetch_home_games(config: dict, season_year: int) -> list[dict]:
    """
    Scrape all home games of the club from the nuLiga Hallenspielplan.

    Returns a list of dicts with keys:
    day, date, time, hall, game_nr, ak, home, guest, score
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
    html = io.StringIO(result.content.decode("utf-8"))
    table = pd.read_html(html, header=0, attrs={"class": "result-set"})[0]

    # Drop obsolete columns and rename
    table.drop(table.columns[[9, 10, 11]], axis=1, inplace=True)
    table.columns = [
        "day", "date", "time", "hall", "game_nr", "ak", "home", "guest", "score",
    ]

    table["hall"] = table["hall"].astype(str)
    table[["day", "date"]] = table[["day", "date"]].ffill()

    # Keep only games in own halls
    mask = table["hall"].apply(
        lambda game: any(hall in game for hall in config["hallIds"])
    )
    table = table[mask]

    # Drop "spielfrei" rows (no game number)
    table = table[table["game_nr"].notna()]

    # Normalize types and whitespace (NA-safe, pandas 2.x and 3.x)
    def _clean(value) -> str:
        return "" if pd.isna(value) else str(value).strip()

    games = []
    for rec in table[GAME_FIELDS].to_dict("records"):
        games.append({
            "day": _clean(rec["day"]),
            "date": _clean(rec["date"]),
            "time": _clean(rec["time"]),
            "hall": int(rec["hall"]),
            "game_nr": int(rec["game_nr"]),
            "ak": _clean(rec["ak"]),
            "home": _clean(rec["home"]),
            "guest": _clean(rec["guest"]),
            "score": _clean(rec["score"]),
        })
    logging.info(f"Current home game plan loaded: {len(games)} home games")
    return games
