# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Append-only assignment audit snapshots for game and block duties.
# ---------------------------------------------------------------

from __future__ import annotations

from typing import TYPE_CHECKING

from sqlalchemy.orm import Session

from nuligahelper.db.task_labels import role_label

from nuligahelper.db.roles import BLOCK_PHASE_LABELS
from nuligahelper.db.models import AssignmentAudit
from nuligahelper.db.games import game_display_name
from nuligahelper.db.day_blocks import block_slot_label

if TYPE_CHECKING:
    from nuligahelper.db.models import Assignment, BlockAssignment, DayBlock, Game, Person


def _game_snapshot(game: Game) -> str:
    return (
        f"{game.game_nr} | {game.date or ''} {game.time or ''} | "
        f"{game_display_name(game)}"
    )


def _audit_assignment(
    session: Session,
    assignment: Assignment,
    action: str,
    actor: Person | None,
    actor_tier: str,
) -> None:
    session.add(
        AssignmentAudit(
            actor_person_id=actor.id if actor else None,
            actor_tier=actor_tier,
            action=action,
            affected_person_id=assignment.person_id,
            game_id=assignment.game_id,
            role=role_label(assignment.role),
            slot=assignment.slot,
            actor_name=actor.name if actor else "System",
            affected_person_name=assignment.person.name,
            game_snapshot=_game_snapshot(assignment.game),
        )
    )


def _block_snapshot(block: DayBlock, slot: int | None = None) -> str:
    description = BLOCK_PHASE_LABELS.get(block.phase, block.phase)
    if block.phase in BLOCK_PHASE_LABELS and slot is not None:
        description = block_slot_label(block, slot)
    return f"Saison {block.season_year} | {block.date} | {description}"


def _audit_block_assignment(
    session: Session,
    assignment: BlockAssignment,
    action: str,
    actor: Person | None,
    actor_tier: str,
) -> None:
    session.add(
        AssignmentAudit(
            actor_person_id=actor.id if actor else None,
            actor_tier=actor_tier,
            action=action,
            affected_person_id=assignment.person_id,
            block_id=assignment.block_id,
            role=assignment.block.label,
            slot=assignment.slot,
            actor_name=actor.name if actor else "System",
            affected_person_name=assignment.person.name,
            block_snapshot=_block_snapshot(assignment.block, assignment.slot),
        )
    )
