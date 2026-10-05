"""Offline duty rename preserves identities, private values and recorded audits."""

from pathlib import Path
import shutil
import sqlite3
import tempfile

from alembic import command
from alembic.migration import MigrationContext
from alembic.operations import Operations
from alembic.script import ScriptDirectory

import helpers as h
import db
import schema_migrations
from schema_fixtures import (
    PRE_GAME_DUTY_REVISION, create_game_duty_source, create_legacy_database,
    create_versioned_database,
)


def _rows(path, table):
    with sqlite3.connect(path) as connection:
        return sorted(connection.execute(f'SELECT * FROM "{table}"').fetchall(), key=repr)


def test_game_duty_revision_joins_existing_single_head_without_model_drift():
    assert schema_migrations.head_revisions() == ("0006_game_duty_staffing",)
    scripts = ScriptDirectory.from_config(schema_migrations.alembic_config())
    assert scripts.get_revision("0006_game_duty_staffing").down_revision == PRE_GAME_DUTY_REVISION
    engine = h.make_engine()
    assert schema_migrations.metadata_drift(engine, db.Base.metadata) == []
    schema_migrations.verify_head_schema(engine)
    engine.dispose()


def test_populated_duty_upgrade_preserves_identities_cakes_people_and_all_recorded_audits():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "populated.db"
        engine = create_game_duty_source(path)
        columns, before = schema_migrations._retained_data(path)
        audits = _rows(path, "assignment_audit")
        result = schema_migrations.migrate_to_head(path, engine)
        assert result.previous_state.revision == PRE_GAME_DUTY_REVISION
        assert result.backup_path is not None
        assert schema_migrations.inspect_schema(result.backup_path).revision == PRE_GAME_DUTY_REVISION
        assert schema_migrations._retained_data(path, columns)[1] == before
        assert _rows(path, "assignments") == [
            (401, 100, 11, "Kasse", 0), (402, 100, 12, "Ordnungsdienst", 0),
            (403, 100, 13, "Verkauf", 1),
        ], "the rename must preserve assignment id, game, person and position"
        assert _rows(path, "assignment_audit") == audits, "audit labels and snapshots are immutable"
        with sqlite3.connect(path) as connection:
            assert connection.execute("SELECT COUNT(*) FROM assignments WHERE role='Reinigung'").fetchone() == (0,)
            assert connection.execute("PRAGMA foreign_key_check").fetchall() == []
        assert schema_migrations.metadata_drift(engine, db.Base.metadata) == []
        db.verify_db(engine)
        engine.dispose()


def test_original_cleaning_assignments_traverse_both_renames_without_filling_new_cleaning():
    for versioned in (False, True):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "original-cleaning.db"
            if versioned:
                engine = create_versioned_database(path, "0002_multi_team_membership")
            else:
                create_legacy_database(path)
                engine = db.make_engine(str(path))
            with sqlite3.connect(path) as connection:
                connection.execute("INSERT INTO persons (id, name, is_admin, account_status) "
                                   "VALUES (11, 'Original cleaner', 0, 'active')")
                connection.execute("INSERT INTO games (id, season_year, game_nr, date) "
                                   "VALUES (100, 2026, '9001', '01.11.2026')")
                connection.execute("INSERT INTO assignments VALUES (401, 100, 11, 'Reinigung', 0)")
                connection.execute(
                    "INSERT INTO assignment_audit (id, changed_at, actor_tier, action, "
                    "affected_person_id, game_id, role, slot, actor_name, affected_person_name, game_snapshot) "
                    "VALUES (501, '2026-01-01', 'system', 'claim', 11, 100, 'Reinigung', 0, "
                    "'System', 'Original cleaner', 'original unchanged snapshot')"
                )
            schema_migrations.migrate_to_head(path, engine)
            assert _rows(path, "assignments") == [(401, 100, 11, "Kasse", 0)]
            with sqlite3.connect(path) as connection:
                assert connection.execute("SELECT role, game_snapshot FROM assignment_audit").fetchall() == [
                    ("Reinigung", "original unchanged snapshot")
                ]
            db.verify_db(engine)
            engine.dispose()


def test_duty_collisions_and_unexpected_new_roles_refuse_without_changing_source_and_keep_snapshot():
    cases = [
        ("collision", 100, "Kasse", 0, "Unterstützung/Kasse role collisions"),
        ("target-only", 101, "Kasse", 0, "Unexpected game-duty target roles"),
        ("cleaning-first", 101, "Reinigung", 0, "Unexpected game-duty target roles"),
        ("cleaning-second", 101, "Reinigung", 1, "Unexpected game-duty target roles"),
        ("invalid-support", 101, "Unterstützung", 1, "Unexpected game-duty target roles"),
        ("invalid-security", 101, "Ordnungsdienst", 1, "Unexpected game-duty target roles"),
    ]
    for name, game_id, role, slot, reason in cases:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / f"{name}.db"
            engine = create_game_duty_source(path)
            with sqlite3.connect(path) as connection:
                connection.execute("INSERT INTO assignments VALUES (404, ?, 15, ?, ?)",
                                   (game_id, role, slot))
            before = path.read_bytes()
            try:
                schema_migrations.migrate_to_head(path, engine)
            except schema_migrations.SchemaMigrationError as exc:
                message = str(exc)
                assert reason in message
                assert "Retained backup:" in message
                assert "cashier@example.test" not in message and "1980-02-29" not in message
            else:
                raise AssertionError(f"{name}: unexpected source data must never be merged or discarded")
            assert path.read_bytes() == before, "preflight refusal must precede every duty revision write"
            snapshots = list(path.parent.glob("*.pre-schema-*.db"))
            assert len(snapshots) == 1
            assert schema_migrations.inspect_schema(snapshots[0]).revision == PRE_GAME_DUTY_REVISION
            assert _rows(snapshots[0], "assignments") == _rows(path, "assignments")
            engine.dispose()


def test_predecessor_near_miss_storage_refuses_before_backup_or_schema_writes():
    for modification in ("column", "cake-check"):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / f"near-miss-{modification}.db"
            engine = create_game_duty_source(path)
            if modification == "column":
                with sqlite3.connect(path) as connection:
                    connection.execute("ALTER TABLE assignments ADD COLUMN unexpected TEXT")
            else:
                with engine.connect() as connection:
                    connection.commit()
                    connection.exec_driver_sql("PRAGMA foreign_keys=OFF")
                    operations = Operations(MigrationContext.configure(connection))
                    with operations.batch_alter_table("day_blocks", recreate="always") as batch:
                        batch.drop_constraint("ck_day_block_cake_quantity", type_="check")
                    connection.commit()
                    connection.exec_driver_sql("PRAGMA foreign_keys=ON")
            before = path.read_bytes()
            try:
                schema_migrations.migrate_to_head(path, engine)
            except schema_migrations.SchemaMigrationError as exc:
                assert "game-duty source schema does not match its reviewed revision" in str(exc)
            else:
                raise AssertionError("a recognized predecessor label cannot authorize schema drift")
            assert path.read_bytes() == before
            assert not list(path.parent.glob("*.pre-schema-*.db"))
            engine.dispose()


def test_every_reviewed_predecessor_requires_explicit_upgrade_at_runtime_without_writes():
    for revision in sorted(schema_migrations.PRE_GAME_DUTY_REVISIONS):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "previous.db"
            engine = create_versioned_database(path, revision)
            before = path.read_bytes()
            try:
                db.verify_db(engine)
            except db.SQLiteInitializationError as exc:
                assert "migrate-schema" in str(exc)
            else:
                raise AssertionError(f"startup must require explicit upgrade from {revision}")
            assert path.read_bytes() == before
            assert not list(path.parent.glob("*.pre-schema-*.db"))
            engine.dispose()


def test_head_fingerprints_preserve_new_cleaning_as_a_distinct_current_role():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "new-cleaning.db"
        engine = create_game_duty_source(path)
        schema_migrations.migrate_to_head(path, engine)
        with sqlite3.connect(path) as connection:
            connection.execute("INSERT INTO assignments VALUES (404, 101, 15, 'Reinigung', 1)")
        columns, before = schema_migrations._retained_data(path)
        with sqlite3.connect(path) as connection:
            connection.execute("UPDATE assignments SET role='Kasse' WHERE id=404")
        _, after = schema_migrations._retained_data(path, columns)
        assert before["assignments"] != after["assignments"], "new Reinigung must never share a compatibility fingerprint with Kasse"
        engine.dispose()


def test_failed_identity_or_audit_postflight_retains_private_safe_snapshot_for_offline_recovery():
    for table, modification in (
        ("assignments", "UPDATE assignments SET role='Reinigung' WHERE id=401"),
        ("assignment_audit", "UPDATE assignment_audit SET role='Kasse' WHERE id=501"),
    ):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "recoverable.db"
            engine = create_game_duty_source(path)
            columns, before = schema_migrations._retained_data(path)
            original = schema_migrations._run_alembic

            def corrupt_after_upgrade(connection, operation, revision):
                original(connection, operation, revision)
                if operation is command.upgrade:
                    connection.exec_driver_sql(modification)

            schema_migrations._run_alembic = corrupt_after_upgrade
            try:
                try:
                    schema_migrations.migrate_to_head(path, engine)
                except schema_migrations.SchemaMigrationError as exc:
                    message = str(exc)
                    assert f"Retained-data postflight failed for tables: {table}" in message
                    assert "Retained backup:" in message
                    assert "cashier@example.test" not in message and "1980-02-29" not in message
                else:
                    raise AssertionError("unexpected identity or audit changes must fail postflight")
            finally:
                schema_migrations._run_alembic = original
                engine.dispose()
            snapshots = list(path.parent.glob("*.pre-schema-*.db"))
            assert len(snapshots) == 1
            assert schema_migrations.inspect_schema(snapshots[0]).revision == PRE_GAME_DUTY_REVISION
            assert schema_migrations._retained_data(snapshots[0], columns)[1] == before
            shutil.copyfile(snapshots[0], path)
            assert schema_migrations._retained_data(path, columns)[1] == before
            engine = db.make_engine(str(path))
            schema_migrations.migrate_to_head(path, engine)
            db.verify_db(engine)
            engine.dispose()


if __name__ == "__main__":
    h.run_all(dict(globals()))
