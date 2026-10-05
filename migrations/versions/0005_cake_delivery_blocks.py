"""Add configurable cake delivery to the existing dated task blocks.

Revision ID: 0005_cake_delivery_blocks
Revises: 0004_person_birth_dates
"""

from __future__ import annotations

from datetime import datetime

from alembic import op
import sqlalchemy as sa


revision = "0005_cake_delivery_blocks"
down_revision = "0004_person_birth_dates"
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

    # SQLite cannot widen CHECK constraints in place. Offline migration turns
    # foreign keys off around these copies and checks every reference afterward.
    with op.batch_alter_table("day_blocks", recreate="always") as batch:
        batch.add_column(sa.Column("cake_quantity", sa.Integer(), nullable=True))
        batch.add_column(sa.Column("delivery_time", sa.String(length=5), nullable=True))
        batch.drop_constraint("ck_day_block_phase", type_="check")
        batch.create_check_constraint(
            "ck_day_block_phase", "phase IN ('preparation', 'cake_delivery', 'cleanup')"
        )
        batch.create_check_constraint(
            "ck_day_block_cake_quantity",
            "cake_quantity IS NULL OR (typeof(cake_quantity) = 'integer' AND cake_quantity >= 0)",
        )
        batch.create_check_constraint(
            "ck_day_block_cake_metadata",
            "phase = 'cake_delivery' OR (cake_quantity IS NULL AND delivery_time IS NULL)",
        )
    with op.batch_alter_table("block_assignments", recreate="always") as batch:
        batch.drop_constraint("ck_block_assignment_slot", type_="check")
        batch.create_check_constraint("ck_block_assignment_slot", "slot >= 0")

    dated_groups = []
    for season, game_date in bind.exec_driver_sql(
        "SELECT DISTINCT season_year, date FROM games"
    ).fetchall():
        try:
            datetime.strptime(game_date or "", "%d.%m.%Y")
        except (TypeError, ValueError):
            continue
        dated_groups.append({"season": season, "date": game_date})
    if dated_groups:
        bind.execute(sa.text(
            "INSERT INTO day_blocks (season_year, date, phase) "
            "VALUES (:season, :date, 'cake_delivery')"
        ), dated_groups)

    for table, (selected, before) in retained.items():
        clause = " WHERE phase <> 'cake_delivery'" if table == "day_blocks" else ""
        after = bind.exec_driver_sql(f'SELECT {selected} FROM "{table}"{clause}').fetchall()
        if sorted((tuple(row) for row in after), key=repr) != before:
            raise RuntimeError(f"Cake-block migration changed retained data in {table}")
    actual = bind.exec_driver_sql(
        "SELECT season_year, date, cake_quantity, delivery_time FROM day_blocks "
        "WHERE phase='cake_delivery'"
    ).fetchall()
    expected = [(row["season"], row["date"], None, None) for row in dated_groups]
    if sorted((tuple(row) for row in actual), key=repr) != sorted(expected, key=repr):
        raise RuntimeError("Cake-block migration did not seed exactly the unconfigured dated groups")


def downgrade() -> None:
    raise RuntimeError(
        "The cake-block migration is restored from its offline snapshot, not downgraded in place."
    )
