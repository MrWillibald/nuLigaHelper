# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Slot lookup, row serialization and authorization shared by game and
# block duty writers.
# ---------------------------------------------------------------

from __future__ import annotations

from sqlalchemy import select, update
from sqlalchemy.orm import Session

import age_eligibility

from db.roles import BLOCK_CAKE_DELIVERY, BLOCK_PHASES, ROLE_SLOT_COUNT, offered_positions
from db.errors import DutyLockedError
from db.models import (
    ACCOUNT_ACTIVE,
    Assignment,
    BlockAssignment,
    DayBlock,
    Game,
    Person,
    person_teams,
    Team,
)
from db.day_blocks import block_capacity, block_is_configured
from db.clock import duty_is_locked


def _stored_slot(
    session: Session, game_id: int, role: str, slot: int
) -> Assignment | None:
    return session.scalars(
        select(Assignment).where(
            Assignment.game_id == game_id,
            Assignment.role == role,
            Assignment.slot == slot,
        ).execution_options(populate_existing=True)
    ).first()


def _validate_slot(role: str, slot: int) -> None:
    if role not in ROLE_SLOT_COUNT or slot < 0 or slot >= ROLE_SLOT_COUNT[role]:
        raise ValueError("Ungültiger Aufgabenplatz.")


def _slot_assignment(game: Game, role: str, slot: int) -> Assignment | None:
    return next(
        (a for a in game.assignments if a.role == role and a.slot == slot), None
    )


def _lock_game(session: Session, game_id: int) -> Game:
    """Obtain writer serialization before reading any duty decision."""
    with session.no_autoflush:
        session.execute(
            update(Game).where(Game.id == game_id).values(id=Game.id)
            .execution_options(synchronize_session=False)
        )
    session.flush()
    stored = session.get(Game, game_id, populate_existing=True)
    if stored is None:
        raise ValueError("Spiel wurde nicht gefunden.")
    return stored


def _authorize_duty_write(
    session: Session, container: Game | DayBlock, person_id: int,
    actor_id: int | None, actor_tier: str,
) -> Person | None:
    """Recheck current account and tier/team authority inside writer serialization."""
    if actor_tier in {"system", "cli"} and actor_id is None:
        return None
    actor = session.get(Person, actor_id, populate_existing=True) if actor_id else None
    if actor is None or actor.account_status != ACCOUNT_ACTIVE:
        raise ValueError("Für diese Änderung ist eine aktive Anmeldung erforderlich.")
    if actor_tier not in {"member", "mv", "admin"}:
        raise ValueError("Keine Berechtigung für diese Änderung.")
    if actor_tier == "admin":
        if not actor.is_admin:
            raise ValueError("Keine Administratorberechtigung für diese Änderung.")
        return actor
    import common
    duty_date = age_eligibility.parse_game_date(container.date)
    if duty_date is not None and duty_date < common.effective_today():
        raise ValueError("Vergangene Einteilungen können nicht geändert werden.")
    if actor.id == person_id:
        return actor
    if actor_tier == "mv" and isinstance(container, Game) and container.team_id:
        managed = session.scalar(select(Team.id).join(
            person_teams, person_teams.c.team_id == Team.id
        ).where(
            Team.id == container.team_id,
            Team.mv_person_id == actor.id,
            person_teams.c.person_id == person_id,
        ))
        own_membership = session.scalar(select(person_teams.c.person_id).where(
            person_teams.c.team_id == container.team_id,
            person_teams.c.person_id == actor.id,
        ))
        if managed is not None and own_membership is not None:
            return actor
    raise ValueError("Keine Berechtigung für diese Einteilung.")


def _validate_retirement(
    session: Session, container: Game | DayBlock,
    assignment: Assignment | BlockAssignment, action: str,
    actor_id: int | None, actor_tier: str,
) -> None:
    """Only an obsolete, unlocked game duty may be explicitly retired."""
    if action != "release":
        raise ValueError("Ungültiger Grund für die Freigabe.")
    if duty_is_locked(session, container, assignment):
        raise DutyLockedError()
    if isinstance(container, Game) and (assignment.role, assignment.slot) not in offered_positions(container):
        return  # Explicit release-only retirement of an obsolete game duty.
    raise ValueError("Die Einteilung bleibt verbindlich. Bitte Ersatz suchen lassen.")


def _validate_block_slot(block: DayBlock, slot: int) -> None:
    if (
        block.phase not in BLOCK_PHASES
        or type(slot) is not int
        or not 0 <= slot < block_capacity(block)
    ):
        if block.phase == BLOCK_CAKE_DELIVERY and not block_is_configured(block):
            raise ValueError("Die Kuchenlieferung muss zuerst eingerichtet werden.")
        raise ValueError("Ungültiger Aufgabenplatz.")


def _lock_day_block(session: Session, block_id: int) -> DayBlock:
    """Read saved capacity only after obtaining the shared SQLite writer lock."""
    with session.no_autoflush:
        session.execute(
            update(DayBlock).where(DayBlock.id == block_id).values(id=DayBlock.id)
            .execution_options(synchronize_session=False)
        )
    session.flush()
    stored = session.get(DayBlock, block_id, populate_existing=True)
    if stored is None:
        raise ValueError("Tagesblock wurde nicht gefunden.")
    return stored


def _stored_block_slot(
    session: Session, block_id: int, slot: int
) -> BlockAssignment | None:
    return session.scalars(
        select(BlockAssignment).where(
            BlockAssignment.block_id == block_id,
            BlockAssignment.slot == slot,
        ).execution_options(populate_existing=True)
    ).first()
