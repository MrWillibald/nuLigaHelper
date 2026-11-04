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

from nuligahelper.db import age_eligibility, clock
from nuligahelper.db.models import DayBlock, Game
from nuligahelper.db.people import get_or_create_team
from nuligahelper.db.audit import _audit_block_assignment
from nuligahelper.db.day_blocks import block_date_valid, block_slot_label, reconcile_day_blocks
from nuligahelper.db.locks import enforce_due_duty_locks, release_unbound_game_locks


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
class CancelledBlockDuty:
    """Snapshot of a removed duty; its assignment no longer exists when notifying."""
    person_id: int
    block_date: str
    task_label: str


@dataclass
class SyncEvents:
    shifts: list[ShiftEvent] = field(default_factory=list)
    referee_alerts: list[RefereeEvent] = field(default_factory=list)
    new_games: list[GameEvent] = field(default_factory=list)
    removed_games: list[GameEvent] = field(default_factory=list)
    cancelled_block_duties: list[CancelledBlockDuty] = field(default_factory=list)


def _cancel_date_block_duties(
    session: Session, season_year: int, active_dates: set[str]
) -> tuple[list[CancelledBlockDuty], set[str]]:
    """Remove block duties of cancelled dates; open and past dates are retained."""
    with session.no_autoflush:
        session.execute(update(DayBlock).where(DayBlock.season_year == season_year)
                        .values(id=DayBlock.id).execution_options(synchronize_session=False))
    session.flush()
    blocks = list(session.scalars(select(DayBlock).where(
        DayBlock.season_year == season_year
    ).execution_options(populate_existing=True)))
    today = clock.duty_now().date()
    cancelled: list[CancelledBlockDuty] = []
    retained: set[str] = set()
    for block in blocks:
        # A previously loaded empty relationship cannot decide the outcome.
        session.expire(block, ["assignments"])
        if block.date in active_dates or not block.assignments:
            continue
        block_date = age_eligibility.parse_game_date(block.date)
        if block_date is None or block_date < today:
            retained.add(block.date)
            continue
        for assignment in sorted(block.assignments, key=lambda item: item.slot):
            if assignment.replacement_token is not None:
                _audit_block_assignment(session, assignment, "request_close", None, "system")
            _audit_block_assignment(session, assignment, "date_remove", None, "system")
            cancelled.append(CancelledBlockDuty(
                assignment.person_id, block.date, block_slot_label(block, assignment.slot),
            ))
            person = assignment.person
            block.assignments.remove(assignment)
            session.flush()
            session.expire(person, ["block_assignments"])
    return cancelled, retained


def sync_games(session: Session, scraped: list[dict], season_year: int) -> SyncEvents:
    """
    Merge freshly scraped game data into the database.

    - New games are inserted.
    - Existing games get their scraped fields updated; date/time changes and
      newly missing referees ("§77") are reported via the returned events.
    - Games no longer present in the scrape stay untouched but are logged.
    - Day blocks exist for calendar dates only. Duties of a current or future date
      without any scraped game are removed and returned for notification; occupied
      blocks of open or past dates are retained. An empty scrape is refused.
    - Saved duty locks of every stored game of the season are released when the
      game's current date no longer justifies them (later not-yet-due or open
      date), including games moved before this rule existed.

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
    if not scraped and session.scalar(
        select(Game.id).where(Game.season_year == season_year).limit(1)
    ) is not None:
        raise ValueError("Der Spielplan enthält keine Spiele; es wird nichts geändert.")
    # Serialize with claims before any stored game or block is changed.
    with session.no_autoflush:
        session.execute(update(Game).where(Game.season_year == season_year).values(id=Game.id)
                        .execution_options(synchronize_session=False))
    session.flush()
    active_dates = {
        date for date in (str(rec.get("date") or "").strip() for rec in scraped)
        if block_date_valid(date)
    }
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

    session.flush()
    released = sum(release_unbound_game_locks(session, game) for game in existing.values())
    if released:
        logging.info(f"Released {released} duty locks of postponed or undated games")

    # Teams mirror the scraped age classes ("ak") so they are always available
    for ak in sorted({rec["ak"] for rec in scraped if rec["ak"]}):
        get_or_create_team(session, ak)

    session.flush()
    cancelled, retained_dates = _cancel_date_block_duties(session, season_year, active_dates)
    events.cancelled_block_duties.extend(cancelled)
    reconcile_day_blocks(session, season_year, active_dates, retained_dates)
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
