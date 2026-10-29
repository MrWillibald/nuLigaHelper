# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Duty clock and the 09:00 day-before cutoff. Patch db.clock.duty_now in
# tests; other modules call clock.duty_now() through this module.
# ---------------------------------------------------------------

from __future__ import annotations

from datetime import datetime, timedelta
from typing import TYPE_CHECKING
from zoneinfo import ZoneInfo

from sqlalchemy import select
from sqlalchemy.orm import Session

import age_eligibility

from db.models import DutyCutoff

if TYPE_CHECKING:
    from db.models import Assignment, BlockAssignment, DayBlock, Game


def _cutoff_time(session: Session, container: Game | DayBlock) -> datetime | None:
    return session.scalar(select(DutyCutoff.locked_at).where(
        DutyCutoff.season_year == container.season_year,
        DutyCutoff.date == container.date,
    ))


def duty_now() -> datetime:
    """Berlin wall time with the shared effective date for debug and test clocks."""
    import common
    current = datetime.now(ZoneInfo("Europe/Berlin"))
    effective = common.effective_today()
    return current.replace(year=effective.year, month=effective.month, day=effective.day)


def duty_cutoff_at(container_or_date: Game | DayBlock | str) -> datetime | None:
    value = container_or_date if isinstance(container_or_date, str) else container_or_date.date
    scheduled = age_eligibility.parse_game_date(value)
    if scheduled is None:
        return None
    before = scheduled - timedelta(days=1)
    return datetime(before.year, before.month, before.day, 9, tzinfo=ZoneInfo("Europe/Berlin"))


def duty_is_locked(
    session: Session, container: Game | DayBlock,
    assignment: Assignment | BlockAssignment | None = None,
) -> bool:
    """The fixed Berlin cutoff applies even when no reminder job has run."""
    cutoff = duty_cutoff_at(container)
    return bool((assignment is not None and assignment.locked_at is not None)
                or _cutoff_time(session, container) is not None
                or (cutoff is not None and duty_now() >= cutoff))


def _claim_lock_time(session: Session, container: Game | DayBlock) -> datetime | None:
    saved = _cutoff_time(session, container)
    if saved is not None:
        return saved
    cutoff = duty_cutoff_at(container)
    return cutoff.replace(tzinfo=None) if cutoff is not None and duty_now() >= cutoff else None
