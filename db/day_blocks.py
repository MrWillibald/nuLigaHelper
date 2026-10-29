# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Dated day blocks: capacity, labels, calculated times and reconciliation.
# ---------------------------------------------------------------

from __future__ import annotations

import logging
import re
from datetime import datetime, timedelta
from typing import TYPE_CHECKING

from sqlalchemy import select, update
from sqlalchemy.orm import Session

from db.roles import (
    BLOCK_CAKE_DELIVERY,
    BLOCK_CLEANUP,
    BLOCK_PHASE_LABELS,
    BLOCK_PHASES,
    BLOCK_PREPARATION,
    BLOCK_SLOT_COUNT,
)
from db.models import DayBlock
from db.games import get_games_on_date

if TYPE_CHECKING:
    from db.models import Game


def get_day_blocks(
    session: Session, season_year: int, block_date: str
) -> list[DayBlock]:
    """Return a date's blocks in preparation/cake/cleanup order."""
    blocks = list(
        session.scalars(
            select(DayBlock).where(
                DayBlock.season_year == season_year, DayBlock.date == block_date
            )
        )
    )
    order = {phase: index for index, phase in enumerate(BLOCK_PHASES)}
    return sorted(blocks, key=lambda block: (order.get(block.phase, 99), block.id))


def _valid_cake_delivery_time(value: str | None) -> bool:
    return isinstance(value, str) and bool(
        re.fullmatch(r"(?:[01][0-9]|2[0-3]):[0-5][0-9]", value)
    )


def block_is_configured(block: DayBlock) -> bool:
    """Distinguish an explicit zero-cake request from missing setup."""
    if block.phase in (BLOCK_PREPARATION, BLOCK_CLEANUP):
        return True
    return (
        block.phase == BLOCK_CAKE_DELIVERY
        and type(block.cake_quantity) is int
        and block.cake_quantity >= 0
        and _valid_cake_delivery_time(block.delivery_time)
    )


def block_capacity(block: DayBlock) -> int:
    """Return configured positions, without inventing a cake quantity."""
    if block.phase in (BLOCK_PREPARATION, BLOCK_CLEANUP):
        return BLOCK_SLOT_COUNT
    if block.phase == BLOCK_CAKE_DELIVERY and block_is_configured(block):
        return block.cake_quantity
    return 0


def block_slot_label(block_or_phase: DayBlock | str, slot: int) -> str:
    phase = block_or_phase.phase if isinstance(block_or_phase, DayBlock) else block_or_phase
    # Display labels are also used for durable history after capacity changes.
    # A cake position keeps its original name even after it no longer exists.
    if (
        phase not in BLOCK_PHASE_LABELS
        or type(slot) is not int
        or slot < 0
        or (phase != BLOCK_CAKE_DELIVERY and slot >= BLOCK_SLOT_COUNT)
    ):
        raise ValueError("Ungültiger Aufgabenplatz.")
    return f"{BLOCK_PHASE_LABELS[phase]} {slot + 1}"


def calculated_block_time(
    block: DayBlock, games: list[Game] | None = None
) -> datetime | None:
    """Return the saved cake time or calculated bookend, retaining dates."""
    try:
        base_date = datetime.strptime(block.date, "%d.%m.%Y")
    except (TypeError, ValueError):
        return None
    if block.phase == BLOCK_CAKE_DELIVERY:
        if not _valid_cake_delivery_time(block.delivery_time):
            return None
        hour, minute = map(int, block.delivery_time.split(":"))
        return base_date.replace(hour=hour, minute=minute)
    if games is None:
        session = Session.object_session(block)
        if session is None:
            return None
        games = get_games_on_date(session, block.date)
        games = [game for game in games if game.season_year == block.season_year]
    candidates = []
    for game in games:
        token = (game.time or "").split()
        match = re.fullmatch(r"(\d{1,2}):(\d{2})", token[0] if token else "")
        if not match:
            continue
        hour, minute = map(int, match.groups())
        if hour > 23 or minute > 59:
            continue
        candidates.append(base_date.replace(hour=hour, minute=minute))
    if not candidates:
        return None
    if block.phase == BLOCK_PREPARATION:
        return min(candidates) - timedelta(minutes=90)
    if block.phase == BLOCK_CLEANUP:
        return max(candidates) + timedelta(minutes=60)
    return None


def reconcile_day_blocks(
    session: Session, season_year: int, active_dates: set[str]
) -> None:
    """Ensure dated phases and remove only genuinely empty vanished blocks."""
    with session.no_autoflush:
        session.execute(update(DayBlock).where(DayBlock.season_year == season_year)
                        .values(id=DayBlock.id).execution_options(synchronize_session=False))
    session.flush()
    existing = list(session.scalars(select(DayBlock).where(
        DayBlock.season_year == season_year
    ).execution_options(populate_existing=True)))
    for block in existing:
        session.expire(block, ["assignments"])
    if any(block.date not in active_dates and block.assignments for block in existing):
        raise ValueError("Ein Heimspieltag mit verbindlichen Einteilungen kann nicht entfernt werden.")
    by_key = {(block.date, block.phase): block for block in existing}
    for block_date in sorted(active_dates):
        for phase in BLOCK_PHASES:
            if (block_date, phase) not in by_key:
                session.add(
                    DayBlock(season_year=season_year, date=block_date, phase=phase)
                )
    session.flush()
    for block in existing:
        if block.date in active_dates:
            continue
        logging.info(
            "day_block_removed season_year=%s date=%s phase=%s removed_assignments=%s",
            block.season_year,
            block.date,
            block.phase,
            0,
        )
        session.delete(block)
    session.flush()
