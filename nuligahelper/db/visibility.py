# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Shared home-game-day visibility boundary for overview pages.
# ---------------------------------------------------------------

from __future__ import annotations

import time
from datetime import date

from sqlalchemy import update
from sqlalchemy.exc import OperationalError
from sqlalchemy.orm import Session

from nuligahelper.db import age_eligibility

from nuligahelper.db import storage
from nuligahelper.db.errors import SQLiteInitializationError
from nuligahelper.db.models import ACCOUNT_ACTIVE, Person, ScheduleVisibilitySetting
from nuligahelper.db.storage import _wait_for_assignment_retry


def schedule_visibility_boundary(session: Session) -> date | None:
    """Read the initialized singleton without repairing missing state on GET."""
    setting = session.get(ScheduleVisibilitySetting, 1, populate_existing=True)
    if setting is None:
        raise SQLiteInitializationError("Schedule visibility setting is missing; verify the schema.")
    return setting.hide_from


def schedule_date_visible(scraped_date: str | None, hide_from: date | None) -> bool:
    """Compare calendar dates; unresolved dates remain visible for correction."""
    scheduled = age_eligibility.parse_game_date(scraped_date)
    return hide_from is None or scheduled is None or scheduled < hide_from


def update_schedule_visibility(
    session: Session, *, actor_id: int, hide_from: str, expected_hide_from: str,
) -> bool:
    """Serialize, recheck active admin and compare saved state; caller commits."""
    from nuligahelper import common

    deadline = time.monotonic() + storage.SQLITE_TIMEOUT_SECONDS
    while True:
        try:
            with session.no_autoflush:
                session.execute(update(ScheduleVisibilitySetting).where(
                    ScheduleVisibilitySetting.id == 1,
                ).values(id=ScheduleVisibilitySetting.id)
                    .execution_options(synchronize_session=False))
            saved = schedule_visibility_boundary(session)
            actor = session.get(Person, actor_id, populate_existing=True)
            if actor is None or actor.account_status != ACCOUNT_ACTIVE or not actor.is_admin:
                raise ValueError("Nur aktive Administratoren dürfen die Sichtbarkeit ändern.")
            if not isinstance(expected_hide_from, str) or not isinstance(hide_from, str):
                raise ValueError("Das Formular ist unvollständig. Bitte lade die Seite neu.")
            if expected_hide_from != (saved.isoformat() if saved else ""):
                raise ValueError("Die Sichtbarkeit wurde zwischenzeitlich geändert. Bitte prüfe den gespeicherten Wert.")
            if hide_from == expected_hide_from:
                return False
            try:
                proposed = date.fromisoformat(hide_from) if hide_from else None
                if proposed and proposed.isoformat() != hide_from:
                    raise ValueError
            except ValueError:
                raise ValueError("Bitte ein gültiges Datum für „Heimspieltage ausblenden ab“ angeben.") from None
            if proposed is not None and proposed <= common.effective_today():
                raise ValueError("„Heimspieltage ausblenden ab“ muss in der Zukunft liegen.")
            session.get(ScheduleVisibilitySetting, 1).hide_from = proposed
            session.flush()
            return True
        except OperationalError as exc:
            session.rollback()
            _wait_for_assignment_retry(deadline, exc)
