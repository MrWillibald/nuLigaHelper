"""Rename first cleaning duties and audit the explicitly requested second-duty removal."""

from datetime import datetime

from alembic import op
import sqlalchemy as sa

revision = "0008_single_wischer"
down_revision = "0007_binding_duty_assignments"
branch_labels = None
depends_on = None

SOURCE_REVISIONS = {"0006_game_duty_staffing", down_revision}


def removal_plan(bind, source_revision):
    """Read-only preflight; never interpret historical pre-game-duty Reinigung here."""
    unexpected = bind.execute(sa.text(
        "SELECT id FROM assignments WHERE role='Wischer'"
        + (" OR (role='Reinigung' AND slot NOT IN (0, 1))"
           if source_revision in SOURCE_REVISIONS else "")
    )).fetchall()
    if unexpected:
        raise RuntimeError("Unexpected Wischer target roles or cleaning positions require manual resolution: "
                           f"assignment ids={[row[0] for row in unexpected]!r}")
    if source_revision not in SOURCE_REVISIONS:
        return []
    rows = bind.execute(sa.text(
        "SELECT a.id, a.game_id, a.person_id, p.name, g.season_year, g.game_nr, "
        "g.date, g.time, g.home, g.guest FROM assignments a "
        "JOIN persons p ON p.id=a.person_id JOIN games g ON g.id=a.game_id "
        "WHERE a.role='Reinigung' AND a.slot=1 ORDER BY a.id"
    )).fetchall()
    next_id = bind.exec_driver_sql("SELECT COALESCE(MAX(id), 0)+1 FROM assignment_audit").scalar_one()
    changed_at = datetime.now().isoformat(sep=" ", timespec="microseconds")
    return [dict(
        id=next_id + index, changed_at=changed_at, actor_person_id=None,
        actor_tier="system", action="migration_remove", affected_person_id=row.person_id,
        game_id=row.game_id, block_id=None, role="Reinigung", slot=1,
        actor_name="System (Migration 0008)", affected_person_name=row.name,
        game_snapshot=(f"{row.season_year} | {row.game_nr} | {row.date or ''} {row.time or ''} | "
                       f"{row.home or ''} – {row.guest or ''} | Einteilung {row.id}"),
        block_snapshot=None,
    ) for index, row in enumerate(rows)]


def upgrade():
    bind = op.get_bind()
    planned = bind.info.get("wischer_removals")
    actual = removal_plan(bind, down_revision)
    if planned is None:
        planned = actual
    # Guard the source again inside the upgrade transaction. Only timestamps
    # may differ from the guarded command's earlier read-only preflight.
    if [{k: v for k, v in row.items() if k != "changed_at"} for row in planned] != [
        {k: v for k, v in row.items() if k != "changed_at"} for row in actual
    ]:
        raise RuntimeError("Wischer migration source changed after preflight")
    retained = {}
    inspector = sa.inspect(bind)
    for table in inspector.get_table_names():
        if table == "alembic_version":
            continue
        columns = [column["name"] for column in inspector.get_columns(table)]
        rows = [dict(row) for row in bind.exec_driver_sql(f'SELECT * FROM "{table}"').mappings()]
        if table == "assignments":
            rows = [row for row in rows if not (row["role"] == "Reinigung" and row["slot"] == 1)]
            for row in rows:
                if row["role"] == "Reinigung":
                    row["role"] = "Wischer"
        elif table == "assignment_audit":
            rows.extend(planned)
        retained[table] = (columns, sorted([tuple(row[c] for c in columns) for row in rows], key=repr))

    # One savepoint covers evidence and deletion even for callers with SQLite's
    # legacy transaction handling. No runtime release helper or provider is used.
    with bind.begin_nested():
        for row in planned:
            columns = ", ".join(row)
            values = ", ".join(f":{key}" for key in row)
            bind.execute(sa.text(f"INSERT INTO assignment_audit ({columns}) VALUES ({values})"), row)
        bind.exec_driver_sql("DELETE FROM assignments WHERE role='Reinigung' AND slot=1")
        bind.exec_driver_sql("UPDATE assignments SET role='Wischer' WHERE role='Reinigung' AND slot=0")
        for table, (columns, expected) in retained.items():
            selected = ", ".join(f'"{column}"' for column in columns)
            rows = bind.exec_driver_sql(f'SELECT {selected} FROM "{table}"').fetchall()
            if sorted(map(tuple, rows), key=repr) != expected:
                raise RuntimeError(f"Wischer migration changed retained data in {table}")


def downgrade():
    raise RuntimeError("Restore the validated pre-migration snapshot with writers stopped.")
