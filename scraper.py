# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Scraper for the BHV/nuLiga Hallenspielplan (home games only)
# ---------------------------------------------------------------

import io
import logging
import re
from datetime import datetime

import pandas as pd
import requests

NULIGA_URL = (
    "https://bhv-handball.liga.nu/cgi-bin/WebObjects/nuLigaHBDE.woa/wa/clubMeetings"
)

GAME_FIELDS = [
    "day", "date", "time", "hall", "game_nr", "ak", "home", "guest", "score",
]


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
    return parse_home_games(result.content.decode("utf-8"), config["hallIds"])


def normalize_age_group(value: str) -> str:
    """Return the stable normalized age-group text used by SPF identities."""
    return " ".join(value.split()).casefold()


def is_spielfest_age_group(value: str | None) -> bool:
    """Whether an age group contains the standalone marker ``SPF``."""
    return bool(re.search(r"(?<!\w)spf(?!\w)", value or "", re.IGNORECASE))


def spielfest_game_number(date_text: str, age_group: str) -> str:
    """Build the canonical pseudo number for one SPF date/age-group event."""
    try:
        iso_date = datetime.strptime(date_text, "%d.%m.%Y").date().isoformat()
    except ValueError as exc:
        raise ValueError(f"Ungültiges Spielfest-Datum: {date_text!r}") from exc
    normalized_age_group = normalize_age_group(age_group)
    if not normalized_age_group:
        raise ValueError("Leere Spielfest-Altersklasse.")
    return f"SPF:{iso_date}:{normalized_age_group}"


def collapse_spielfeste(games: list[dict]) -> list[dict]:
    """Collapse individual SPF matches into one task-bearing event per date/AK."""
    ordinary: list[dict] = []
    groups: dict[tuple[str, str], list[dict]] = {}
    for game in games:
        if not is_spielfest_age_group(game.get("ak")):
            ordinary.append(game)
            continue
        key = (game.get("date", ""), normalize_age_group(game.get("ak", "")))
        groups.setdefault(key, []).append(game)

    collapsed: list[dict] = []
    for (date_text, normalized_ak), rows in groups.items():
        pseudo_number = spielfest_game_number(date_text, normalized_ak)
        halls = {row.get("hall") for row in rows}
        days = {" ".join(str(row.get("day", "")).split()) for row in rows}
        if len(halls) != 1 or None in halls or len(days) != 1 or "" in days:
            raise ValueError(
                f"Widersprüchliche Spielfest-Daten {pseudo_number}: {rows!r}"
            )
        timed_rows = []
        for row in rows:
            try:
                parsed_time = datetime.strptime(row.get("time", ""), "%H:%M").time()
            except ValueError as exc:
                raise ValueError(
                    f"Ungültige Spielfest-Zeit {pseudo_number}: {rows!r}"
                ) from exc
            timed_rows.append((parsed_time, row))
        earliest = min(timed_rows, key=lambda item: item[0])[1]
        collapsed.append({
            "day": next(iter(days)),
            "date": date_text,
            "time": earliest["time"],
            "hall": next(iter(halls)),
            "game_nr": pseudo_number,
            "ak": " ".join(str(rows[0]["ak"]).split()),
            "home": "Spielfest",
            "guest": "",
            "score": "",
        })
    return ordinary + collapsed


def parse_home_games(html_text: str, hall_ids: list[str]) -> list[dict]:
    """Parse a nuLiga result page without performing network access."""
    html = io.StringIO(html_text)
    table = pd.read_html(html, header=0, attrs={"class": "result-set"})[0]
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
    for rec in table[GAME_FIELDS].to_dict("records"):
        game = {
            "day": _clean(rec["day"]),
            "date": _clean(rec["date"]),
            "time": _clean(rec["time"]),
            "hall": int(rec["hall"]),
            "game_nr": str(int(rec["game_nr"])),
            "ak": _clean(rec["ak"]),
            "home": _clean(rec["home"]),
            "guest": _clean(rec["guest"]),
            "score": _clean(rec["score"]),
        }
        games.append(game)

    games = collapse_spielfeste(games)
    by_key: dict[str, dict] = {}
    for game in games:
        previous = by_key.get(game["game_nr"])
        if previous is not None:
            raise ValueError(
                "Mehrdeutige Spielidentität "
                f"{game['game_nr']}: {previous!r} / {game!r}"
            )
        by_key[game["game_nr"]] = game
    logging.info(f"Current home game plan loaded: {len(games)} home games")
    return games
