# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Persisting due day-before locks and closing open replacement requests.
# ---------------------------------------------------------------

from __future__ import annotations

import time

from sqlalchemy import select, update
from sqlalchemy.exc import OperationalError
from sqlalchemy.orm import Session

from nuligahelper.db import age_eligibility

from nuligahelper.db import clock
from nuligahelper.db import storage
from nuligahelper.db.models import DayBlock, DutyCutoff, Game
from nuligahelper.db.storage import _wait_for_assignment_retry
from nuligahelper.db.clock import duty_cutoff_at, duty_is_locked
from nuligahelper.db.audit import _audit_assignment, _audit_block_assignment


def enforce_due_duty_locks(session: Session, container: Game | DayBlock):
    """Stage due cohort locks; request boundaries commit this before mutation."""
    if duty_is_locked(session, container):
        return lock_day_before_batch(session, container.season_year, container.date)
    return None


def lock_due_duties(session: Session) -> int:
    """Stage all due dated cohorts for a separately committed request preflight."""
    cohorts = set(session.execute(select(Game.season_year, Game.date)).all())
    cohorts.update(session.execute(select(DayBlock.season_year, DayBlock.date)).all())
    locked = 0
    current = clock.duty_now()
    for season_year, duty_date in cohorts:
        cutoff = duty_cutoff_at(duty_date)
        if cutoff is not None and current >= cutoff:
            lock_day_before_batch(session, season_year, duty_date)
            locked += 1
    return locked


def lock_day_before_batch(session: Session, season_year: int, date: str) -> DutyCutoff | None:
    """Persist a due fixed cutoff for the whole dated cohort; caller commits."""
    cutoff = duty_cutoff_at(date)
    if cutoff is not None and clock.duty_now() < cutoff:
        return None
    if age_eligibility.parse_game_date(date) is None:
        raise ValueError("Ungültiges Datum für die Vortagserinnerung.")
    deadline = time.monotonic() + storage.SQLITE_TIMEOUT_SECONDS
    while True:
        try:
            with session.no_autoflush:
                session.execute(update(DutyCutoff).where(
                    DutyCutoff.season_year == season_year, DutyCutoff.date == date,
                ).values(id=DutyCutoff.id).execution_options(synchronize_session=False))
            session.flush()
            marker = session.scalars(select(DutyCutoff).where(
                DutyCutoff.season_year == season_year, DutyCutoff.date == date,
            ).execution_options(populate_existing=True)).first()
            if marker is None:
                marker = DutyCutoff(season_year=season_year, date=date, locked_at=cutoff.replace(tzinfo=None))
                session.add(marker)
                session.flush()
            games = list(session.scalars(select(Game).where(
                Game.season_year == season_year, Game.date == date,
            ).execution_options(populate_existing=True)))
            blocks = list(session.scalars(select(DayBlock).where(
                DayBlock.season_year == season_year, DayBlock.date == date,
            ).execution_options(populate_existing=True)))
            for container in [*games, *blocks]:
                session.expire(container, ["assignments"])
                audit = _audit_assignment if isinstance(container, Game) else _audit_block_assignment
                for assignment in container.assignments:
                    session.refresh(assignment)
                    if assignment.replacement_token is not None:
                        assignment.replacement_token = None
                        audit(session, assignment, "request_close", None, "system")
                    if assignment.locked_at is None:
                        assignment.locked_at = marker.locked_at
                        audit(session, assignment, "lock", None, "system")
            session.flush()
            return marker
        except OperationalError as exc:
            session.rollback()
            _wait_for_assignment_retry(deadline, exc)
