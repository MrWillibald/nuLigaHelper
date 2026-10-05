"""Offline cake-block upgrades, source fingerprints and snapshot recovery."""

from pathlib import Path
import shutil
import sqlite3
import tempfile

from alembic import command
from alembic.migration import MigrationContext
from alembic.operations import Operations

import helpers as h
import db
import schema_migrations
from schema_fixtures import create_versioned_database


def _older_database(path):
    engine = create_versioned_database(path, "0004_person_birth_dates")
    with sqlite3.connect(path) as connection:
        connection.executemany(
            "INSERT INTO persons (id, name, is_admin, account_status, birth_date, email) "
            "VALUES (?, ?, 0, 'active', ?, ?)",
            [(11, "Legacy adult", "1980-02-29", "adult@example.test"),
             (12, "Legacy unknown", None, None)]
        )
        connection.executemany("INSERT INTO person_teams VALUES (?, 1)", [(11,), (12,)])
        connection.execute("UPDATE teams SET mv_person_id=11 WHERE id=1")
        connection.executemany(
            "INSERT INTO games (id, season_year, game_nr, date, ak) VALUES (?, ?, ?, ?, 'BL M')",
            [(100, 2026, "9001", "01.11.2026"), (101, 2026, "9002", "01.11.2026"),
             (102, 2026, "9003", "08.11.2026"), (103, 2025, "9001", "01.11.2026"),
             (104, 2026, "9004", None), (105, 2026, "9005", ""),
             (106, 2026, "9006", "31.02.2026")]
        )
        connection.executemany(
            "INSERT INTO day_blocks (id, season_year, date, phase) VALUES (?, 2026, ?, ?)",
            [(201, "01.11.2026", "preparation"), (202, "01.11.2026", "cleanup"),
             (203, "08.11.2026", "preparation"), (204, "08.11.2026", "cleanup")]
        )
        connection.executemany("INSERT INTO block_assignments VALUES (?, ?, ?, ?)",
                               [(301, 201, 11, 2), (302, 202, 12, 0)])
        connection.execute("INSERT INTO assignments VALUES (401, 100, 11, 'Verkauf', 0)")
        connection.execute(
            "INSERT INTO assignment_audit "
            "(id, changed_at, actor_tier, action, affected_person_id, role, slot, "
            "actor_name, affected_person_name, block_id, block_snapshot) "
            "VALUES (501, '2026-01-01', 'system', 'claim', 11, 'Vorbereitung', 2, "
            "'System', 'Legacy adult', 201, '01.11.2026 | Vorbereitung')"
        )
        connection.execute(
            "INSERT INTO assignment_audit "
            "(id, changed_at, actor_tier, action, affected_person_id, role, slot, "
            "actor_name, affected_person_name, game_id, game_snapshot) "
            "VALUES (502, '2026-01-01', 'system', 'claim', 11, 'Verkauf', 0, "
            "'System', 'Legacy adult', 100, 'historical game snapshot')"
        )
        connection.execute(
            "INSERT INTO auth_tokens (id, nonce, code, purpose, person_id, issued_at, expires_at) "
            "VALUES (601, 'synthetic-token', '123456', 'login', 11, '2026-01-01', '2026-01-01 00:15')"
        )
    return engine


def _drop_check(engine, table, name):
    with engine.connect() as connection:
        connection.commit()
        connection.exec_driver_sql("PRAGMA foreign_keys=OFF")
        operations = Operations(MigrationContext.configure(connection))
        with operations.batch_alter_table(table, recreate="always") as batch:
            batch.drop_constraint(name, type_="check")
        connection.commit()
        connection.exec_driver_sql("PRAGMA foreign_keys=ON")


def test_cake_revision_seeds_unique_unconfigured_valid_dates_and_preserves_every_old_value():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "old-cake-source.db"
        engine = _older_database(path)
        columns, before = schema_migrations._retained_data(path)
        result = schema_migrations.migrate_to_head(path, engine)
        assert result.previous_state.revision == "0004_person_birth_dates"
        assert result.revision == schema_migrations.HEAD_REVISION
        assert result.backup_path is not None
        assert schema_migrations.inspect_schema(result.backup_path).revision == "0004_person_birth_dates"
        _, after = schema_migrations._retained_data(path, columns)
        assert after == before, "migration must preserve all previous identities, contacts, dates and history"
        with sqlite3.connect(path) as connection:
            assert connection.execute(
                "SELECT season_year, date, cake_quantity, delivery_time FROM day_blocks "
                "WHERE phase='cake_delivery' ORDER BY season_year, date"
            ).fetchall() == [(2025, "01.11.2026", None, None),
                            (2026, "01.11.2026", None, None),
                            (2026, "08.11.2026", None, None)]
            assert connection.execute("SELECT COUNT(*) FROM block_assignments").fetchone() == (2,)
            assert connection.execute("PRAGMA foreign_key_check").fetchall() == []
        assert schema_migrations.metadata_drift(engine, db.Base.metadata) == []
        db.verify_db(engine)
        engine.dispose()


def test_migrated_constraints_allow_four_cakes_but_refuse_negative_or_duplicate_positions():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "capacity.db"
        engine = _older_database(path)
        schema_migrations.migrate_to_head(path, engine)
        with sqlite3.connect(path) as connection:
            cake_id = connection.execute(
                "SELECT id FROM day_blocks WHERE season_year=2026 AND date='01.11.2026' "
                "AND phase='cake_delivery'"
            ).fetchone()[0]
            connection.execute("UPDATE day_blocks SET cake_quantity=4, delivery_time='10:00' WHERE id=?",
                               (cake_id,))
            connection.execute("INSERT INTO block_assignments (block_id, person_id, slot) VALUES (?, 11, 3)",
                               (cake_id,))
            for sql, parameters in (
                ("INSERT INTO block_assignments (block_id, person_id, slot) VALUES (?, 12, -1)", (cake_id,)),
                ("INSERT INTO block_assignments (block_id, person_id, slot) VALUES (?, 12, 3)", (cake_id,)),
                ("INSERT INTO block_assignments (block_id, person_id, slot) VALUES (?, 11, 0)", (cake_id,)),
                ("UPDATE day_blocks SET cake_quantity=-1 WHERE id=?", (cake_id,)),
                ("UPDATE day_blocks SET cake_quantity=1.5 WHERE id=?", (cake_id,)),
                ("UPDATE day_blocks SET cake_quantity=4 WHERE id=201", ()),
            ):
                try:
                    connection.execute(sql, parameters)
                except sqlite3.IntegrityError:
                    pass
                else:
                    raise AssertionError("migration must retain reviewed CHECK and assignment uniqueness guarantees")
        engine.dispose()


def test_cake_source_near_misses_fail_before_backup_or_any_schema_write():
    for modification in ("column", "phase-check", "slot-check"):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / f"near-miss-{modification}.db"
            engine = _older_database(path)
            if modification == "column":
                with sqlite3.connect(path) as connection:
                    connection.execute("ALTER TABLE day_blocks ADD COLUMN unreviewed TEXT")
            elif modification == "phase-check":
                _drop_check(engine, "day_blocks", "ck_day_block_phase")
            else:
                _drop_check(engine, "block_assignments", "ck_block_assignment_slot")
            before = path.read_bytes()
            try:
                schema_migrations.migrate_to_head(path, engine)
            except schema_migrations.SchemaMigrationError as exc:
                assert "cake-block source schema does not match" in str(exc)
            else:
                raise AssertionError("a recognized version label cannot authorize a near-miss source")
            assert path.read_bytes() == before
            assert not list(path.parent.glob("*.pre-schema-*.db"))
            engine.dispose()


def test_current_head_with_removed_cake_check_is_refused_without_repair():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "forged-check.db"
        engine = db.make_engine(str(path))
        db.initialize_db(engine)
        _drop_check(engine, "day_blocks", "ck_day_block_cake_quantity")
        before = path.read_bytes()
        for operation, error_type in (
            (lambda: schema_migrations.migrate_to_head(path, engine), schema_migrations.SchemaMigrationError),
            (lambda: db.verify_db(engine), db.SQLiteInitializationError),
        ):
            try:
                operation()
            except error_type as exc:
                assert "check-constraint drift" in str(exc)
            else:
                raise AssertionError("Alembic's missing CHECK comparison must not permit forged head storage")
        assert path.read_bytes() == before
        assert not list(path.parent.glob("*.pre-schema-*.db"))
        engine.dispose()


def test_cake_retained_data_failure_has_private_safe_snapshot_and_offline_recovery():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "recoverable.db"
        engine = _older_database(path)
        columns, before = schema_migrations._retained_data(path)
        original = schema_migrations._run_alembic

        def corrupt_after_upgrade(connection, operation, revision):
            original(connection, operation, revision)
            if operation is command.upgrade:
                connection.exec_driver_sql("UPDATE day_blocks SET date='02.11.2026' WHERE id=201")

        schema_migrations._run_alembic = corrupt_after_upgrade
        try:
            try:
                schema_migrations.migrate_to_head(path, engine)
            except schema_migrations.SchemaMigrationError as exc:
                message = str(exc)
                assert "Retained-data postflight failed for tables: day_blocks" in message
                assert "Retained backup:" in message
                assert "adult@example.test" not in message and "1980-02-29" not in message
            else:
                raise AssertionError("changing an old block must fail postflight even when new cakes are seeded")
        finally:
            schema_migrations._run_alembic = original
            engine.dispose()
        snapshots = list(path.parent.glob("*.pre-schema-*.db"))
        assert len(snapshots) == 1
        assert schema_migrations.inspect_schema(snapshots[0]).revision == "0004_person_birth_dates"
        assert schema_migrations._retained_data(snapshots[0], columns)[1] == before
        shutil.copyfile(snapshots[0], path)
        assert schema_migrations.inspect_schema(path).revision == "0004_person_birth_dates"
        assert schema_migrations._retained_data(path, columns)[1] == before
        engine = db.make_engine(str(path))
        schema_migrations.migrate_to_head(path, engine)
        db.verify_db(engine)
        engine.dispose()


if __name__ == "__main__":
    h.run_all(dict(globals()))
