"""Add the unrestricted global schedule visibility setting."""

from alembic import op
import sqlalchemy as sa

revision = "0009_schedule_visibility"
down_revision = "0008_single_wischer"
branch_labels = None
depends_on = None


def upgrade():
    table = op.create_table(
        "schedule_visibility_setting",
        sa.Column("id", sa.Integer(), primary_key=True, nullable=False),
        sa.Column("hide_from", sa.Date(), nullable=True),
        sa.CheckConstraint("id = 1", name="ck_schedule_visibility_singleton"),
    )
    op.bulk_insert(table, [{"id": 1, "hide_from": None}])


def downgrade():
    raise RuntimeError("Restore the validated snapshot and matching application with writers stopped.")
