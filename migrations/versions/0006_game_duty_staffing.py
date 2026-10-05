"""Rename current Unterstützung to Kasse without rewriting historical audits.

Revision ID: 0006_game_duty_staffing
Revises: 0005_cake_delivery_blocks
"""

from __future__ import annotations

from alembic import op
import sqlalchemy as sa


revision = "0006_game_duty_staffing"
down_revision = "0005_cake_delivery_blocks"
branch_labels = None
depends_on = None


def upgrade() -> None:
    bind = op.get_bind()
    collisions = bind.execute(sa.text(
        "SELECT old.game_id, old.slot FROM assignments AS old "
        "JOIN assignments AS new ON new.game_id=old.game_id AND new.slot=old.slot "
        "WHERE old.role='Unterstützung' AND new.role='Kasse'"
    )).fetchall()
    if collisions:
        raise RuntimeError(
            "Unterstützung/Kasse role collisions require manual resolution: "
            f"{collisions!r}"
        )
    # Neither the new semantic role nor expanded positions are valid source
    # data. Refuse them rather than guessing whether they are a partial upgrade.
    unexpected = bind.execute(sa.text(
        "SELECT id, game_id, role, slot FROM assignments "
        "WHERE role IN ('Kasse', 'Reinigung') "
        "OR (role IN ('Unterstützung', 'Ordnungsdienst') AND slot <> 0)"
    )).fetchall()
    if unexpected:
        raise RuntimeError(
            "Unexpected game-duty target roles or positions require manual resolution: "
            f"{unexpected!r}"
        )

    retained = {}
    inspector = sa.inspect(bind)
    for table in inspector.get_table_names():
        if table == "alembic_version":
            continue
        columns = [column["name"] for column in inspector.get_columns(table)]
        selected = ", ".join(f'"{column}"' for column in columns)
        rows = [tuple(row) for row in bind.exec_driver_sql(
            f'SELECT {selected} FROM "{table}"'
        ).fetchall()]
        if table == "assignments":
            role_index = columns.index("role")
            rows = [tuple("Kasse" if index == role_index and value == "Unterstützung"
                          else value for index, value in enumerate(row)) for row in rows]
        retained[table] = (selected, sorted(rows, key=repr))

    bind.execute(sa.text(
        "UPDATE assignments SET role='Kasse' WHERE role='Unterstützung'"
    ))
    for table, (selected, before) in retained.items():
        after = bind.exec_driver_sql(f'SELECT {selected} FROM "{table}"').fetchall()
        if sorted((tuple(row) for row in after), key=repr) != before:
            raise RuntimeError(f"Game-duty migration changed retained data in {table}")


def downgrade() -> None:
    raise RuntimeError(
        "The game-duty migration is restored from its offline snapshot, not downgraded in place."
    )
