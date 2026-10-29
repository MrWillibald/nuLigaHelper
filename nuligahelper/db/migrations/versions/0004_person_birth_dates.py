"""Add nullable, date-only birth dates without inferring legacy ages.

Revision ID: 0004_person_birth_dates
Revises: 0003_game_day_task_blocks
"""

from __future__ import annotations

from alembic import op
import sqlalchemy as sa


revision = "0004_person_birth_dates"
down_revision = "0003_game_day_task_blocks"
branch_labels = None
depends_on = None


def upgrade() -> None:
    bind = op.get_bind()
    # Adding a nullable column must change no existing value or relationship.
    # Keep snapshots in memory and never print personal data on failure.
    tables = sa.inspect(bind).get_table_names()
    retained = {}
    for table in tables:
        if table == "alembic_version":
            continue
        columns = [column["name"] for column in sa.inspect(bind).get_columns(table)]
        selected = ", ".join(f'"{column}"' for column in columns)
        rows = bind.exec_driver_sql(f'SELECT {selected} FROM "{table}"').fetchall()
        retained[table] = (selected, sorted(tuple(row) for row in rows))

    op.add_column("persons", sa.Column("birth_date", sa.Date(), nullable=True))

    for table, (selected, before) in retained.items():
        after = bind.exec_driver_sql(f'SELECT {selected} FROM "{table}"').fetchall()
        if sorted(tuple(row) for row in after) != before:
            raise RuntimeError(f"Birth-date migration changed retained data in {table}")
    if bind.exec_driver_sql(
        "SELECT COUNT(*) FROM persons WHERE birth_date IS NOT NULL"
    ).scalar_one() != 0:
        raise RuntimeError("Birth-date migration invented legacy birth dates")


def downgrade() -> None:
    raise RuntimeError(
        "The birth-date migration is restored from its offline snapshot, not downgraded in place."
    )
