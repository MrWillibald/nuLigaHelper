# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Merging scraped games into the database and the resulting sync events.
# ---------------------------------------------------------------

from __future__ import annotations

import logging
from dataclasses import dataclass, field

from sqlalchemy import select, update
from sqlalchemy.orm import Session

from db.models import DayBlock, Game
from db.people import get_or_create_team
from db.day_blocks import reconcile_day_blocks
from db.locks import enforce_due_duty_locks


@dataclass
class ShiftEvent:
    game_id: int
    game_nr: str
    old_date: str
    old_time: str
    new_date: str
    new_time: str


@dataclass
class RefereeEvent:
    game_id: int
    game_nr: str
    date: str
    time: str


@dataclass
class GameEvent:
    game_id: int
    game_nr: str
    ak: str


@dataclass
class SyncEvents:
    shifts: list[ShiftEvent] = field(default_factory=list)
    referee_alerts: list[RefereeEvent] = field(default_factory=list)
    new_games: list[GameEvent] = field(default_factory=list)
    removed_games: list[GameEvent] = field(default_factory=list)


def sync_games(session: Session, scraped: list[dict], season_year: int) -> SyncEvents:
    """
    Merge freshly scraped game data into the database.

    - New games are inserted.
    - Existing games get their scraped fields updated; date/time changes and
      newly missing referees ("§77") are reported via the returned events.
    - Games no longer present in the scrape stay untouched but are logged.

    `scraped` items must be dicts with keys:
    day, date, time, hall, game_nr, ak, home, guest, score
    """
    events = SyncEvents()
    game_numbers = [str(rec.get("game_nr", "")).strip() for rec in scraped]
    if any(not number for number in game_numbers):
        raise ValueError("Jedes Spiel benötigt eine Spielnummer.")
    duplicate_keys = sorted({key for key in game_numbers if game_numbers.count(key) > 1})
    if duplicate_keys:
        conflicts = [rec for rec in scraped if str(rec.get("game_nr", "")).strip() in duplicate_keys]
        raise ValueError(
            f"Mehrdeutige Spielidentität {duplicate_keys}: {conflicts!r}"
        )
    # Serialize the preflight with claims before inspecting any disappearing
    # container. A previously loaded empty relationship cannot authorize removal.
    with session.no_autoflush:
        session.execute(update(Game).where(Game.season_year == season_year).values(id=Game.id)
                        .execution_options(synchronize_session=False))
    session.flush()
    active_dates = {str(rec.get("date") or "").strip() for rec in scraped if rec.get("date")}
    disappearing = list(session.scalars(select(DayBlock).where(
        DayBlock.season_year == season_year
    ).execution_options(populate_existing=True)))
    for block in disappearing:
        session.expire(block, ["assignments"])
    if any(block.date not in active_dates and block.assignments for block in disappearing):
        raise ValueError("Ein Heimspieltag mit verbindlichen Einteilungen kann nicht entfernt werden.")
    existing = {
        g.game_nr: g
        for g in session.scalars(select(Game).where(Game.season_year == season_year))
    }
    new_games: list[Game] = []

    for rec in scraped:
        rec = {**rec, "game_nr": str(rec["game_nr"]).strip()}
        game_nr = rec["game_nr"]
        game = existing.get(game_nr)

        if game is None:
            game = Game(season_year=season_year, **rec)
            session.add(game)
            new_games.append(game)
            logging.info(f"New game {game_nr} added to database")
            continue

        if (game.date != rec["date"]) or (game.time != rec["time"]):
            events.shifts.append(
                ShiftEvent(
                    game_id=game.id,
                    game_nr=game_nr,
                    old_date=game.date or "",
                    old_time=game.time or "",
                    new_date=rec["date"],
                    new_time=rec["time"],
                )
            )
        if ("§77" in rec["score"]) and ("§77" not in (game.score or "")):
            events.referee_alerts.append(
                RefereeEvent(
                    game_id=game.id,
                    game_nr=game_nr,
                    date=rec["date"],
                    time=rec["time"],
                )
            )

        enforce_due_duty_locks(session, game)
        for f in ("day", "date", "time", "hall", "ak", "home", "guest", "score"):
            setattr(game, f, rec[f])

    scraped_keys = set(game_numbers)
    for game_nr in sorted(set(existing) - scraped_keys):
        game = existing[game_nr]
        events.removed_games.append(
            GameEvent(game.id, game.game_nr, game.ak or "")
        )
        logging.warning(f"Game {game.game_nr} not contained in online plan anymore")

    # Teams mirror the scraped age classes ("ak") so they are always available
    for ak in sorted({rec["ak"] for rec in scraped if rec["ak"]}):
        get_or_create_team(session, ak)

    session.flush()
    reconcile_day_blocks(
        session,
        season_year,
        {
            str(record.get("date") or "").strip()
            for record in scraped
            if str(record.get("date") or "").strip()
        },
    )
    events.new_games.extend(
        GameEvent(game.id, game.game_nr, game.ak or "")
        for game in new_games
    )
    session.commit()
    logging.info(
        f"Sync completed: {len(events.new_games)} new, {len(events.shifts)} shifted, "
        f"{len(events.referee_alerts)} referee alerts"
    )
    return events
