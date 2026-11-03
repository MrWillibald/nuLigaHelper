"""Add automatic game-day task blocks and optional Unterstützung.

Revision ID: 0003_game_day_task_blocks
Revises: 0002_multi_team_membership
"""

from __future__ import annotations

from alembic import op
import sqlalchemy as sa


revision = "0003_game_day_task_blocks"
down_revision = "0002_multi_team_membership"
branch_labels = None
depends_on = None


def upgrade() -> None:
    bind = op.get_bind()
    collisions = bind.execute(
        sa.text(
            "SELECT old.game_id, old.slot FROM assignments AS old "
            "JOIN assignments AS new ON new.game_id=old.game_id AND new.slot=old.slot "
            "WHERE old.role='Reinigung' AND new.role='Unterstützung'"
        )
    ).fetchall()
    if collisions:
        raise RuntimeError(
            "Reinigung/Unterstützung role collisions require manual resolution: "
            f"{collisions!r}"
        )

    op.create_table(
        "day_blocks",
        sa.Column("id", sa.Integer(), primary_key=True),
        sa.Column("season_year", sa.Integer(), nullable=False),
        sa.Column("date", sa.String(length=20), nullable=False),
        sa.Column("phase", sa.String(length=20), nullable=False),
        sa.UniqueConstraint(
            "season_year", "date", "phase", name="uq_day_block_season_date_phase"
        ),
        sa.CheckConstraint(
            "phase IN ('preparation', 'cleanup')", name="ck_day_block_phase"
        ),
    )
    op.create_table(
        "block_assignments",
        sa.Column("id", sa.Integer(), primary_key=True),
        sa.Column(
            "block_id", sa.Integer(), sa.ForeignKey("day_blocks.id"), nullable=False
        ),
        sa.Column(
            "person_id", sa.Integer(), sa.ForeignKey("persons.id"), nullable=False
        ),
        sa.Column("slot", sa.Integer(), nullable=False),
        sa.UniqueConstraint("block_id", "person_id", name="uq_block_person"),
        sa.UniqueConstraint("block_id", "slot", name="uq_block_slot"),
        sa.CheckConstraint(
            "slot >= 0 AND slot < 3", name="ck_block_assignment_slot"
        ),
    )
    with op.batch_alter_table("assignment_audit", recreate="always") as batch:
        batch.add_column(sa.Column("block_id", sa.Integer(), nullable=True))
        batch.add_column(sa.Column("block_snapshot", sa.String(length=300), nullable=True))
        batch.alter_column(
            "game_snapshot", existing_type=sa.String(length=300), nullable=True
        )
        batch.create_foreign_key(
            "fk_assignment_audit_block_id_day_blocks",
            "day_blocks",
            ["block_id"],
            ["id"],
            ondelete="SET NULL",
        )
        batch.create_check_constraint(
            "ck_assignment_audit_one_target",
            "(game_snapshot IS NOT NULL AND block_snapshot IS NULL) OR "
            "(game_snapshot IS NULL AND block_snapshot IS NOT NULL)",
        )

    bind.execute(
        sa.text(
            "INSERT INTO day_blocks (season_year, date, phase) "
            "SELECT DISTINCT season_year, date, 'preparation' FROM games "
            "WHERE date IS NOT NULL AND trim(date) <> '' "
            "UNION ALL "
            "SELECT DISTINCT season_year, date, 'cleanup' FROM games "
            "WHERE date IS NOT NULL AND trim(date) <> ''"
        )
    )
    bind.execute(
        sa.text(
            "UPDATE assignments SET role='Unterstützung' WHERE role='Reinigung'"
        )
    )


def downgrade() -> None:
    raise RuntimeError(
        "The day-block migration is restored from its offline snapshot, not downgraded in place."
    )
