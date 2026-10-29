# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Game duty state machine: claim, release, replacement request, cancel
# and atomic takeover (also for day blocks via _replacement_change).
# ---------------------------------------------------------------

from __future__ import annotations

import secrets
import time
from typing import TYPE_CHECKING

from sqlalchemy import delete, select, update
from sqlalchemy.exc import IntegrityError, OperationalError
from sqlalchemy.orm import Session

from task_labels import role_label

from db import storage
from db.roles import ROLE_SLOT_COUNT, offered_positions
from db.errors import (
    AgeEligibilityError,
    AssignmentTemporarilyUnavailableError,
    DutyLockedError,
    SlotConflictError,
)
from db.models import (
    ACCOUNT_ACTIVE,
    Assignment,
    AssignmentAudit,
    BlockAssignment,
    Game,
    Person,
)
from db.storage import _wait_for_assignment_retry
from db.clock import _claim_lock_time, duty_is_locked
from db.audit import _audit_assignment, _audit_block_assignment, _game_snapshot
from db.staffing import claim_eligibility
from db.slots import (
    _authorize_duty_write,
    _lock_day_block,
    _lock_game,
    _slot_assignment,
    _stored_block_slot,
    _stored_slot,
    _validate_block_slot,
    _validate_retirement,
    _validate_slot,
)
from db.locks import enforce_due_duty_locks

if TYPE_CHECKING:
    from db.models import DayBlock


def _replacement_change(
    session: Session, container: Game | DayBlock, role: str | None, slot: int,
    expected_person_id: int | None, expected_request_token: str | None,
    person: Person | None, actor: Person | None, actor_tier: str,
    *, cancel: bool = False,
) -> Assignment | BlockAssignment:
    is_game = isinstance(container, Game)
    if is_game:
        _validate_slot(role, slot)
    elif type(slot) is not int or slot < 0:
        raise ValueError("Ungültiger Aufgabenplatz.")
    container_id, actor_id = container.id, actor.id if actor else None
    person_id = person.id if person else None
    deadline = time.monotonic() + storage.SQLITE_TIMEOUT_SECONDS
    while True:
        try:
            stored = (_lock_game(session, container_id) if is_game
                      else _lock_day_block(session, container_id))
            enforce_due_duty_locks(session, stored)
            current = (_stored_slot(session, container_id, role, slot) if is_game
                       else _stored_block_slot(session, container_id, slot))
            current_id = current.person_id if current else None
            if current_id != expected_person_id or current is None:
                raise SlotConflictError(current_id)
            if (person_id is not None or cancel) and (
                not expected_request_token or current.replacement_token != expected_request_token
            ):
                raise SlotConflictError(current_id)
            if duty_is_locked(session, stored, current):
                raise DutyLockedError()
            if is_game:
                if (role, slot) not in offered_positions(stored):
                    raise ValueError("Diese Aufgabe wird für dieses Spiel nicht angeboten.")
            else:
                _validate_block_slot(stored, slot)
            stored_actor = _authorize_duty_write(
                session, stored, person_id if person_id is not None else current_id,
                actor_id, actor_tier,
            )
            audit = _audit_assignment if is_game else _audit_block_assignment
            if cancel:
                current.replacement_token = None
                audit(session, current, "request_close", stored_actor, actor_tier)
                session.flush()
                return current
            if person_id is None:
                if current.replacement_token is None:
                    current.replacement_token = secrets.token_hex(24)
                    audit(session, current, "request_open", stored_actor, actor_tier)
                    session.flush()
                return current
            replacement = session.get(Person, person_id, populate_existing=True)
            if replacement is None or replacement.account_status != ACCOUNT_ACTIVE:
                raise ValueError("Diese Person kann nicht eingeteilt werden.")
            assignment_type = Assignment if is_game else BlockAssignment
            owner_column = Assignment.game_id if is_game else BlockAssignment.block_id
            other = session.scalars(select(assignment_type).where(
                owner_column == container_id, assignment_type.person_id == person_id,
            )).first()
            if other is not None:
                raise ValueError("Diese Person ist für dieses Spiel oder diesen Tagesblock bereits eingeteilt.")
            if is_game:
                session.expire(stored, ["assignments"])
                for sibling in stored.assignments:
                    session.refresh(sibling.person, ["birth_date"])
                reason = claim_eligibility(stored, role, replacement, slot)
                if reason is not None:
                    raise AgeEligibilityError(reason)
            audit(session, current, "takeover_out", stored_actor, actor_tier)
            audit(session, current, "request_close", stored_actor, actor_tier)
            current.person = replacement
            current.person_id = person_id
            current.replacement_token = None
            audit(session, current, "takeover_in", stored_actor, actor_tier)
            session.flush()
            session.expire(stored, ["assignments"])
            return current
        except OperationalError as exc:
            session.rollback()
            _wait_for_assignment_retry(deadline, exc)
        except IntegrityError:
            session.rollback()
            raise SlotConflictError(None) from None


def request_replacement(session, game, role, slot, expected_person_id,
                        actor=None, actor_tier="system"):
    return _replacement_change(session, game, role, slot, expected_person_id,
                               None, None, actor, actor_tier)


def request_block_replacement(session, block, slot, expected_person_id,
                              actor=None, actor_tier="system"):
    return _replacement_change(session, block, None, slot, expected_person_id,
                               None, None, actor, actor_tier)


def cancel_replacement(session, game, role, slot, expected_person_id,
                       expected_request_token, actor=None, actor_tier="system"):
    """End this request generation while retaining the assigned responsibility."""
    return _replacement_change(session, game, role, slot, expected_person_id,
                               expected_request_token, None, actor, actor_tier, cancel=True)


def cancel_block_replacement(session, block, slot, expected_person_id,
                             expected_request_token, actor=None, actor_tier="system"):
    return _replacement_change(session, block, None, slot, expected_person_id,
                               expected_request_token, None, actor, actor_tier, cancel=True)


def takeover_slot(session, game, role, slot, expected_person_id,
                  expected_request_token, person, actor=None, actor_tier="system"):
    return _replacement_change(session, game, role, slot, expected_person_id,
                               expected_request_token, person, actor, actor_tier)


def takeover_block_slot(session, block, slot, expected_person_id,
                        expected_request_token, person, actor=None, actor_tier="system"):
    return _replacement_change(session, block, None, slot, expected_person_id,
                               expected_request_token, person, actor, actor_tier)


def claim_slot(
    session: Session,
    game: Game,
    role: str,
    slot: int,
    expected_person_id: int | None,
    person: Person,
    actor: Person | None = None,
    actor_tier: str = "system",
) -> Assignment:
    """Claim one slot using bounded, freshly revalidated compare-and-swap."""
    _validate_slot(role, slot)
    game_id = game.id
    person_id = person.id
    actor_id = actor.id if actor else None
    deadline = time.monotonic() + storage.SQLITE_TIMEOUT_SECONDS

    while True:
        try:
            # Acquire the SQLite writer lock BEFORE reading collective coverage.
            # A stale WAL read snapshot fails here and retries from a fresh one.
            with session.no_autoflush:
                session.execute(
                    update(Game).where(Game.id == game_id).values(id=Game.id)
                    .execution_options(synchronize_session=False)
                )
            session.flush()
            stored_game = session.get(Game, game_id, populate_existing=True)
            if stored_game is None:
                raise ValueError("Spiel wurde nicht gefunden.")
            enforce_due_duty_locks(session, stored_game)
            current = _stored_slot(session, game_id, role, slot)
            current_id = current.person_id if current else None
            if current_id != expected_person_id:
                raise SlotConflictError(current_id)
            if current is not None:
                _authorize_duty_write(session, stored_game, person_id, actor_id, actor_tier)
                if current.person_id == person_id:
                    return current
                raise SlotConflictError(current.person_id)

            stored_game = session.get(Game, game_id, populate_existing=True)
            stored_person = session.get(Person, person_id, populate_existing=True)
            if stored_game is None or stored_person is None:
                raise ValueError("Spiel oder Person wurde nicht gefunden.")
            if (role, slot) not in offered_positions(stored_game):
                raise ValueError("Diese Aufgabe wird für dieses Spiel nicht angeboten.")
            if stored_person.account_status != ACCOUNT_ACTIVE:
                raise ValueError("Diese Person kann nicht eingeteilt werden.")
            stored_actor = _authorize_duty_write(
                session, stored_game, person_id, actor_id, actor_tier
            )
            other = session.scalars(
                select(Assignment).where(
                    Assignment.game_id == game_id,
                    Assignment.person_id == person_id,
                )
            ).first()
            if other is not None:
                raise ValueError(
                    f"{stored_person.name} ist für dieses Spiel bereits als "
                    f"'{role_label(other.role)}' eingeteilt."
                )

            session.expire(stored_game, ["assignments"])
            # Refresh sibling sellers too: the identity map may predate a
            # separately committed profile correction or assignment mutation.
            for sibling in stored_game.assignments:
                session.refresh(sibling.person, ["birth_date"])
            reason = claim_eligibility(stored_game, role, stored_person, slot)
            if reason is not None:
                raise AgeEligibilityError(reason)

            assignment = Assignment(
                game=stored_game,
                person=stored_person,
                role=role,
                slot=slot,
                locked_at=_claim_lock_time(session, stored_game),
            )
            session.add(assignment)
            session.flush()
            _audit_assignment(
                session, assignment, "claim", stored_actor, actor_tier
            )
            if assignment.locked_at is not None:
                _audit_assignment(session, assignment, "lock", None, "system")
            session.flush()
            return assignment
        except OperationalError as exc:
            session.rollback()
            _wait_for_assignment_retry(deadline, exc)
        except IntegrityError as exc:
            session.rollback()
            current = _stored_slot(session, game_id, role, slot)
            current_id = current.person_id if current else None
            if current_id != expected_person_id:
                raise SlotConflictError(current_id) from None
            other = session.scalars(
                select(Assignment).where(
                    Assignment.game_id == game_id,
                    Assignment.person_id == person_id,
                )
            ).first()
            if other is not None:
                stored_person = session.get(Person, person_id)
                name = stored_person.name if stored_person else "Die Person"
                raise ValueError(
                    f"{name} ist für dieses Spiel bereits als '{role_label(other.role)}' eingeteilt."
                ) from None
            raise


def release_slot(
    session: Session,
    game: Game,
    role: str,
    slot: int,
    expected_person_id: int | None,
    actor: Person | None = None,
    actor_tier: str = "system",
    action: str = "release",
) -> Person | None:
    """Release one slot using bounded, freshly revalidated compare-and-swap."""
    _validate_slot(role, slot)
    game_id = game.id
    actor_id = actor.id if actor else None
    actor_name = actor.name if actor else "System"
    deadline = time.monotonic() + storage.SQLITE_TIMEOUT_SECONDS

    while True:
        try:
            stored_game = _lock_game(session, game_id)
            enforce_due_duty_locks(session, stored_game)
            current = _stored_slot(session, game_id, role, slot)
            current_id = current.person_id if current else None
            if current_id != expected_person_id:
                raise SlotConflictError(current_id)
            if current is None:
                return None

            person = current.person
            _authorize_duty_write(session, stored_game, current.person_id, actor_id, actor_tier)
            _validate_retirement(session, stored_game, current, action, actor_id, actor_tier)
            audit = AssignmentAudit(
                actor_person_id=actor_id,
                actor_tier=actor_tier,
                action=action,
                affected_person_id=current.person_id,
                game_id=current.game_id,
                role=role_label(current.role),
                slot=current.slot,
                actor_name=actor_name,
                affected_person_name=person.name,
                game_snapshot=_game_snapshot(current.game),
            )
            result = session.execute(
                delete(Assignment).where(
                    Assignment.id == current.id,
                    Assignment.game_id == game_id,
                    Assignment.role == role,
                    Assignment.slot == slot,
                    Assignment.person_id == expected_person_id,
                ).execution_options(synchronize_session=False)
            )
            if getattr(result, "rowcount", None) != 1:
                session.rollback()
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    raise AssignmentTemporarilyUnavailableError()
                time.sleep(min(0.05, remaining))
                continue
            session.add(audit)
            session.flush()
            session.expire(current.game, ["assignments"])
            session.expire(person, ["assignments"])
            session.expunge(current)
            return person
        except OperationalError as exc:
            session.rollback()
            _wait_for_assignment_retry(deadline, exc)
        except IntegrityError:
            session.rollback()
            raise


def set_role_assignments(session: Session, game: Game, role: str,
                          person_ids: list[int], actor: Person | None = None,
                          actor_tier: str = "system") -> None:
    """Legacy seeding helper: fill empty positions, refusing changed commitments."""
    _validate_slot(role, 0)
    if len(person_ids) > ROLE_SLOT_COUNT[role]:
        raise ValueError("Zu viele Personen für diese Aufgabe.")
    try:
        # Offered commitments cannot be replaced by a whole-role edit.
        # Only explicit pre-cutoff retirement of obsolete duties can remove rows.
        for slot in range(ROLE_SLOT_COUNT[role]):
            current = _slot_assignment(game, role, slot)
            desired_id = person_ids[slot] if slot < len(person_ids) else None
            if current is not None and current.person_id != desired_id:
                release_slot(
                    session, game, role, slot, current.person_id, actor, actor_tier
                )
        for slot in range(ROLE_SLOT_COUNT[role]):
            current = _slot_assignment(game, role, slot)
            desired_id = person_ids[slot] if slot < len(person_ids) else None
            if desired_id is None or (
                current is not None and current.person_id == desired_id
            ):
                continue
            person_id = desired_id
            person = session.get(Person, person_id)
            if person is not None:
                claim_slot(session, game, role, slot, None, person, actor, actor_tier)
        session.commit()
    except BaseException:
        # A refused replacement must not leave earlier releases or claims
        # pending for a caller to accidentally commit, including their audits.
        session.rollback()
        raise


def assign_person(
    session: Session,
    game: Game,
    person: Person,
    role: str,
    actor: Person | None = None,
    actor_tier: str = "system",
) -> Assignment:
    """Assign a person to a game with the given role (idempotent).

    A person may only ever hold one task per game – assigning them to
    another role of the same game raises ValueError.
    """
    existing = next(
        (a for a in game.assignments if a.person_id == person.id and a.role == role), None
    )
    if existing is not None:
        return existing
    other = next(
        (a for a in game.assignments if a.person_id == person.id and a.role != role), None
    )
    if other is not None:
        raise ValueError(
            f"{person.name} ist für dieses Spiel bereits als '{role_label(other.role)}' eingeteilt."
        )
    slot = next(
        (
            index
            for index in range(ROLE_SLOT_COUNT.get(role, 0))
            if _slot_assignment(game, role, index) is None
        ),
        None,
    )
    if slot is None:
        raise ValueError("Für diese Aufgabe ist kein Platz mehr frei.")
    return claim_slot(session, game, role, slot, None, person, actor, actor_tier)


def unassign_person(
    session: Session,
    game: Game,
    person: Person,
    role: str,
    actor: Person | None = None,
    actor_tier: str = "system",
) -> bool:
    """Remove an assignment; returns True if something was removed."""
    assignment = next(
        (a for a in game.assignments if a.person_id == person.id and a.role == role), None
    )
    if assignment is None:
        return False
    release_slot(
        session, game, role, assignment.slot, person.id, actor, actor_tier
    )
    return True
