# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Day-block duty writes and cake-delivery configuration.
# ---------------------------------------------------------------

from __future__ import annotations

import time
from typing import TYPE_CHECKING

from sqlalchemy import select
from sqlalchemy.exc import IntegrityError, OperationalError
from sqlalchemy.orm import Session

from db import storage
from db.roles import BLOCK_CAKE_DELIVERY
from db.errors import CakeConfigurationConflictError, SlotConflictError
from db.models import ACCOUNT_ACTIVE, BlockAssignment, Person
from db.storage import _wait_for_assignment_retry
from db.day_blocks import _valid_cake_delivery_time
from db.clock import _claim_lock_time
from db.audit import _audit_block_assignment
from db.slots import (
    _authorize_duty_write,
    _lock_day_block,
    _stored_block_slot,
    _validate_block_slot,
    _validate_retirement,
)
from db.locks import enforce_due_duty_locks

if TYPE_CHECKING:
    from db.models import DayBlock


def configure_cake_block(
    session: Session,
    block: DayBlock,
    delivery_time: str,
    cake_quantity: int,
    expected_delivery_time: str | None,
    expected_cake_quantity: int | None,
) -> DayBlock:
    """Save cake settings atomically with claims; authorization is the caller's job."""
    if not _valid_cake_delivery_time(delivery_time):
        raise ValueError("Bitte eine gültige Lieferzeit im Format HH:MM angeben.")
    if type(cake_quantity) is not int or not 0 <= cake_quantity <= 2**63 - 1:
        raise ValueError("Die Kuchenanzahl muss eine nichtnegative ganze Zahl sein.")
    block_id = block.id
    deadline = time.monotonic() + storage.SQLITE_TIMEOUT_SECONDS
    while True:
        try:
            stored_block = _lock_day_block(session, block_id)
            enforce_due_duty_locks(session, stored_block)
            if stored_block.phase != BLOCK_CAKE_DELIVERY:
                raise ValueError("Nur Kuchenlieferungen haben Lieferzeit und Kuchenanzahl.")
            if (
                stored_block.delivery_time != expected_delivery_time
                or stored_block.cake_quantity != expected_cake_quantity
            ):
                raise CakeConfigurationConflictError(
                    stored_block.delivery_time, stored_block.cake_quantity
                )
            removed_occupant = session.scalars(
                select(BlockAssignment).where(
                    BlockAssignment.block_id == block_id,
                    BlockAssignment.slot >= cake_quantity,
                )
            ).first()
            if removed_occupant is not None:
                raise ValueError(
                    "Die Kuchenanzahl kann keine belegten Kuchenplätze entfernen."
                )
            stored_block.delivery_time = delivery_time
            stored_block.cake_quantity = cake_quantity
            session.flush()
            session.expire(stored_block, ["assignments"])
            return stored_block
        except OperationalError as exc:
            session.rollback()
            _wait_for_assignment_retry(deadline, exc)


def claim_block_slot(
    session: Session,
    block: DayBlock,
    slot: int,
    expected_person_id: int | None,
    person: Person,
    actor: Person | None = None,
    actor_tier: str = "system",
) -> BlockAssignment:
    """Claim one day-block slot with bounded compare-and-swap semantics."""
    if type(slot) is not int or slot < 0:
        raise ValueError("Ungültiger Aufgabenplatz.")
    block_id, person_id = block.id, person.id
    actor_id = actor.id if actor else None
    deadline = time.monotonic() + storage.SQLITE_TIMEOUT_SECONDS
    while True:
        try:
            stored_block = _lock_day_block(session, block_id)
            enforce_due_duty_locks(session, stored_block)
            _validate_block_slot(stored_block, slot)
            current = _stored_block_slot(session, block_id, slot)
            current_id = current.person_id if current else None
            if current_id != expected_person_id:
                raise SlotConflictError(current_id)
            if current is not None:
                _authorize_duty_write(session, stored_block, person_id, actor_id, actor_tier)
                if current.person_id == person_id:
                    return current
                raise SlotConflictError(current.person_id)
            stored_person = session.get(Person, person_id, populate_existing=True)
            if stored_person is None:
                raise ValueError("Tagesblock oder Person wurde nicht gefunden.")
            if stored_person.account_status != ACCOUNT_ACTIVE:
                raise ValueError("Diese Person kann nicht eingeteilt werden.")
            stored_actor = _authorize_duty_write(
                session, stored_block, person_id, actor_id, actor_tier
            )
            other = session.scalars(
                select(BlockAssignment).where(
                    BlockAssignment.block_id == block_id,
                    BlockAssignment.person_id == person_id,
                )
            ).first()
            if other is not None:
                raise ValueError(
                    f"{stored_person.name} ist für diesen Tagesblock bereits als "
                    f"'{stored_block.label}' eingeteilt."
                )
            assignment = BlockAssignment(
                block=stored_block, person=stored_person, slot=slot,
                locked_at=_claim_lock_time(session, stored_block),
            )
            session.add(assignment)
            session.flush()
            _audit_block_assignment(
                session,
                assignment,
                "claim",
                stored_actor,
                actor_tier,
            )
            if assignment.locked_at is not None:
                _audit_block_assignment(session, assignment, "lock", None, "system")
            session.flush()
            return assignment
        except OperationalError as exc:
            session.rollback()
            _wait_for_assignment_retry(deadline, exc)
        except IntegrityError:
            session.rollback()
            current = _stored_block_slot(session, block_id, slot)
            current_id = current.person_id if current else None
            if current_id != expected_person_id:
                raise SlotConflictError(current_id) from None
            other = session.scalars(
                select(BlockAssignment).where(
                    BlockAssignment.block_id == block_id,
                    BlockAssignment.person_id == person_id,
                )
            ).first()
            if other is not None:
                stored_person = session.get(Person, person_id)
                raise ValueError(
                    f"{stored_person.name if stored_person else 'Die Person'} ist für "
                    "diesen Tagesblock bereits eingeteilt."
                ) from None
            raise


def release_block_slot(
    session: Session,
    block: DayBlock,
    slot: int,
    expected_person_id: int | None,
    actor: Person | None = None,
    actor_tier: str = "system",
    action: str = "release",
) -> Person | None:
    """Compatibility refusal endpoint; occupied block duties cannot be released."""
    if type(slot) is not int or slot < 0:
        raise ValueError("Ungültiger Aufgabenplatz.")
    block_id, actor_id = block.id, actor.id if actor else None
    deadline = time.monotonic() + storage.SQLITE_TIMEOUT_SECONDS
    while True:
        try:
            stored_block = _lock_day_block(session, block_id)
            enforce_due_duty_locks(session, stored_block)
            current = _stored_block_slot(session, block_id, slot)
            current_id = current.person_id if current else None
            if current_id != expected_person_id:
                raise SlotConflictError(current_id)
            if current is None:
                _validate_block_slot(stored_block, slot)
                return None
            _authorize_duty_write(session, stored_block, current.person_id, actor_id, actor_tier)
            _validate_retirement(session, stored_block, current, action, actor_id, actor_tier)
        except OperationalError as exc:
            session.rollback()
            _wait_for_assignment_retry(deadline, exc)
