"""Persist replacement requests and the fixed duty cutoff without changing history.

Revision ID: 0007_binding_duty_assignments
Revises: 0006_game_duty_staffing
"""

from alembic import op
import sqlalchemy as sa

revision = "0007_binding_duty_assignments"
down_revision = "0006_game_duty_staffing"
branch_labels = None
depends_on = None


def upgrade() -> None:
    bind = op.get_bind()
    retained = {}
    inspector = sa.inspect(bind)
    for table in inspector.get_table_names():
        if table == "alembic_version":
            continue
        columns = [column["name"] for column in inspector.get_columns(table)]
        selected = ", ".join(f'"{column}"' for column in columns)
        rows = bind.exec_driver_sql(f'SELECT {selected} FROM "{table}"').fetchall()
        retained[table] = (selected, sorted((tuple(row) for row in rows), key=repr))
    for table in ("assignments", "block_assignments"):
        op.add_column(table, sa.Column("replacement_token", sa.String(64), nullable=True))
        op.add_column(table, sa.Column("locked_at", sa.DateTime(), nullable=True))
    op.create_table(
        "duty_cutoffs",
        sa.Column("id", sa.Integer(), primary_key=True),
        sa.Column("season_year", sa.Integer(), nullable=False),
        sa.Column("date", sa.String(20), nullable=False),
        sa.Column("locked_at", sa.DateTime(), nullable=False),
        sa.UniqueConstraint("season_year", "date", name="uq_duty_cutoff_season_date"),
    )
    for table, (selected, before) in retained.items():
        rows = bind.exec_driver_sql(f'SELECT {selected} FROM "{table}"').fetchall()
        if sorted((tuple(row) for row in rows), key=repr) != before:
            raise RuntimeError(f"Binding-duty migration changed retained data in {table}")
    # No historical consent, replacement request or reminder event is invented.
    for table in ("assignments", "block_assignments"):
        if bind.exec_driver_sql(
            f'SELECT COUNT(*) FROM "{table}" WHERE replacement_token IS NOT NULL OR locked_at IS NOT NULL'
        ).scalar_one():
            raise RuntimeError("Binding-duty migration fabricated historical lifecycle state")


def downgrade() -> None:
    raise RuntimeError("Restore the validated pre-migration snapshot with writers stopped.")
