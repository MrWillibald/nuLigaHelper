import sqlite3
import tempfile
import contextlib
import io
from pathlib import Path

import helpers as h
import backup
import db
import manage_db
import schema_migrations
from schema_fixtures import create_legacy_database


def test_schema_inspection_is_read_only_and_recognizes_baseline():
    with tempfile.TemporaryDirectory() as raw_directory:
        path = Path(raw_directory) / "baseline.db"
        create_legacy_database(path)
        before = path.read_bytes()

        state = schema_migrations.inspect_schema(path)

        assert state.kind == "baseline"
        assert state.revision == schema_migrations.BASELINE_REVISION
        assert path.read_bytes() == before
        with sqlite3.connect(path) as connection:
            assert "alembic_version" not in {
                row[0] for row in connection.execute(
                    "SELECT name FROM sqlite_master WHERE type='table'"
                )
            }


def test_schema_inspection_distinguishes_empty_near_miss_and_corrupt():
    with tempfile.TemporaryDirectory() as raw_directory:
        directory = Path(raw_directory)
        missing = directory / "missing.db"
        assert schema_migrations.inspect_schema(missing).kind == "empty"
        assert not missing.exists()

        near_miss = directory / "near-miss.db"
        create_legacy_database(near_miss)
        with sqlite3.connect(near_miss) as connection:
            connection.execute("ALTER TABLE persons ADD COLUMN surprise TEXT")
        state = schema_migrations.inspect_schema(near_miss)
        assert state.kind == "unknown_unversioned"
        assert any("persons columns" in detail for detail in state.details)

        corrupt = directory / "corrupt.db"
        corrupt.write_bytes(b"not a sqlite database")
        assert schema_migrations.inspect_schema(corrupt).kind == "corrupt"


def test_schema_inspection_recognizes_legacy_game_identity():
    with tempfile.TemporaryDirectory() as raw_directory:
        path = Path(raw_directory) / "legacy-game.db"
        with sqlite3.connect(path) as connection:
            connection.execute(
                "CREATE TABLE games (id INTEGER PRIMARY KEY, source_key TEXT NOT NULL)"
            )
        assert schema_migrations.inspect_schema(path).kind == "legacy_game_identity"


def test_schema_inspection_recognizes_known_and_unknown_revisions():
    with tempfile.TemporaryDirectory() as raw_directory:
        path = Path(raw_directory) / "versioned.db"
        with sqlite3.connect(path) as connection:
            connection.execute("CREATE TABLE alembic_version (version_num VARCHAR(64))")
            connection.execute(
                "INSERT INTO alembic_version VALUES (?)",
                (schema_migrations.BASELINE_REVISION,),
            )
        state = schema_migrations.inspect_schema(path)
        assert state.kind == "versioned" and state.revision == schema_migrations.BASELINE_REVISION

        with sqlite3.connect(path) as connection:
            connection.execute("UPDATE alembic_version SET version_num='future_revision'")
        state = schema_migrations.inspect_schema(path)
        assert state.kind == "unknown_revision"
        assert state.revision == "future_revision"


def test_alembic_configuration_has_one_head_and_no_database_url():
    config = schema_migrations.alembic_config()
    assert not config.get_main_option("sqlalchemy.url")
    assert schema_migrations.head_revisions() == ("0003_game_day_task_blocks",)


def test_explicit_initialization_creates_stamps_seeds_and_then_refuses_reuse():
    with tempfile.TemporaryDirectory() as raw_directory:
        path = Path(raw_directory) / "fresh.db"
        engine = db.make_engine(str(path))
        db.initialize_db(engine)
        assert schema_migrations.inspect_schema(path).kind == "head"
        with db.Session(engine) as session:
            support = db.get_support_team(session)
            assert support is not None and support.name == db.SUPPORT_TEAM_NAME
        before = path.read_bytes()
        try:
            db.initialize_db(engine)
        except db.SQLiteInitializationError as exc:
            assert "absent or empty" in str(exc)
        else:
            raise AssertionError("initialization must refuse an existing database")
        assert path.read_bytes() == before


def test_runtime_verification_rejects_missing_and_unversioned_without_writing():
    with tempfile.TemporaryDirectory() as raw_directory:
        directory = Path(raw_directory)
        missing = directory / "missing.db"
        try:
            db.verify_db(db.make_engine(str(missing)))
        except db.SQLiteInitializationError as exc:
            assert "migrate-schema" in str(exc) and "empty" in str(exc)
        else:
            raise AssertionError("missing runtime database must fail closed")
        assert not missing.exists()

        baseline = directory / "baseline.db"
        create_legacy_database(baseline)
        before = baseline.read_bytes()
        try:
            db.verify_db(db.make_engine(str(baseline)))
        except db.SQLiteInitializationError as exc:
            assert "migrate-schema" in str(exc) and "baseline" in str(exc)
        else:
            raise AssertionError("unversioned baseline must require migration")
        assert baseline.read_bytes() == before


def test_runtime_verification_accepts_head_and_rejects_unknown_revision_unchanged():
    with tempfile.TemporaryDirectory() as raw_directory:
        directory = Path(raw_directory)
        head = directory / "head.db"
        engine = db.make_engine(str(head))
        db.initialize_db(engine)
        db.verify_db(engine)

        with sqlite3.connect(head) as connection:
            connection.execute("UPDATE alembic_version SET version_num='future_revision'")
        before = head.read_bytes()
        try:
            db.verify_db(db.make_engine(str(head)))
        except db.SQLiteInitializationError as exc:
            assert "unknown_revision" in str(exc) and "migrate-schema" in str(exc)
        else:
            raise AssertionError("unknown revision must fail closed")
        assert head.read_bytes() == before


def _run_cli(path: Path, *arguments: str) -> str:
    args = manage_db.build_parser().parse_args(
        ["--db", str(path), *arguments]
    )
    output = io.StringIO()
    with contextlib.redirect_stdout(output):
        args.func(args)
    return output.getvalue()


def _system_exit(callback) -> SystemExit:
    try:
        callback()
    except SystemExit as exc:
        return exc
    raise AssertionError("expected SystemExit")


def test_migrate_schema_requires_confirmation_then_adopts_exact_baseline():
    with tempfile.TemporaryDirectory() as raw_directory:
        path = Path(raw_directory) / "baseline.db"
        create_legacy_database(path)
        before = path.read_bytes()
        error = _system_exit(lambda: _run_cli(path, "migrate-schema"))
        assert "--confirm-stopped" in str(error)
        assert path.read_bytes() == before

        output = _run_cli(path, "migrate-schema", "--confirm-stopped")
        assert "Migrated database to schema head" in output
        backup_line = next(line for line in output.splitlines() if line.startswith("Backup: "))
        backup_path = Path(backup_line.removeprefix("Backup: "))
        assert backup_path.exists()
        assert schema_migrations.inspect_schema(backup_path).kind == "baseline"
        assert schema_migrations.inspect_schema(path).kind == "head"


def test_migrate_schema_head_is_noop_and_legacy_or_unknown_are_refused():
    with tempfile.TemporaryDirectory() as raw_directory:
        directory = Path(raw_directory)
        head = directory / "head.db"
        db.initialize_db(db.make_engine(str(head)))
        before = head.read_bytes()
        output = _run_cli(head, "migrate-schema", "--confirm-stopped")
        assert "already at schema head" in output
        assert "Backup:" not in output
        assert head.read_bytes() == before

        legacy = directory / "legacy-game.db"
        with sqlite3.connect(legacy) as connection:
            connection.execute(
                "CREATE TABLE games (id INTEGER PRIMARY KEY, source_key TEXT NOT NULL)"
            )
        error = _system_exit(
            lambda: _run_cli(legacy, "migrate-schema", "--confirm-stopped")
        )
        assert "migrate-game-identity" in str(error)

        unknown = directory / "unknown.db"
        with sqlite3.connect(unknown) as connection:
            connection.execute("CREATE TABLE surprise (id INTEGER PRIMARY KEY)")
        before = unknown.read_bytes()
        error = _system_exit(
            lambda: _run_cli(unknown, "migrate-schema", "--confirm-stopped")
        )
        assert "cannot be migrated automatically" in str(error)
        assert unknown.read_bytes() == before


def test_migrate_schema_reports_backup_and_upgrade_failures():
    with tempfile.TemporaryDirectory() as raw_directory:
        directory = Path(raw_directory)
        backup_failure = directory / "backup-failure.db"
        create_legacy_database(backup_failure)
        before = backup_failure.read_bytes()
        original_snapshot = backup.snapshot_database
        backup.snapshot_database = lambda _path: (_ for _ in ()).throw(
            OSError("synthetic backup failure")
        )
        try:
            error = _system_exit(
                lambda: _run_cli(
                    backup_failure, "migrate-schema", "--confirm-stopped"
                )
            )
        finally:
            backup.snapshot_database = original_snapshot
        assert "Schema backup failed" in str(error)
        assert backup_failure.read_bytes() == before

        upgrade_failure = directory / "upgrade-failure.db"
        create_legacy_database(upgrade_failure)
        original_run = schema_migrations._run_alembic

        def failing_upgrade(connection, operation, revision):
            if operation.__name__ == "upgrade":
                raise RuntimeError("synthetic upgrade failure")
            return original_run(connection, operation, revision)

        schema_migrations._run_alembic = failing_upgrade
        try:
            error = _system_exit(
                lambda: _run_cli(
                    upgrade_failure, "migrate-schema", "--confirm-stopped"
                )
            )
        finally:
            schema_migrations._run_alembic = original_run
        message = str(error)
        assert "synthetic upgrade failure" in message
        assert "Retained backup:" in message
        retained = Path(message.split("Retained backup: ", 1)[1].split(". ", 1)[0])
        assert retained.exists()


def test_membership_revision_preserves_union_identities_and_relationships():
    with tempfile.TemporaryDirectory() as raw_directory:
        path = Path(raw_directory) / "populated-baseline.db"
        create_legacy_database(path)
        with sqlite3.connect(path) as connection:
            connection.executemany(
                "INSERT INTO teams (id, name, is_support) VALUES (?, ?, 0)",
                [(2, "Alpha"), (3, "Beta")],
            )
            connection.executemany(
                "INSERT INTO persons "
                "(id, name, email, phone, team_id, desired_team_id, is_admin, account_status) "
                "VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                [
                    (1, "Active Admin", "admin@example.test", None, 2, None, 1, "active"),
                    (2, "Pending Contactless", None, None, None, 3, 0, "verified"),
                    (3, "Both References", None, "+491700000003", 2, 3, 0, "active"),
                    (4, "No Team", None, None, None, None, 0, "inactive"),
                    (5, "Appointed MV", None, "+491700000005", 3, None, 0, "active"),
                ],
            )
            connection.execute("UPDATE teams SET mv_person_id=5 WHERE id=3")
            connection.execute(
                "INSERT INTO games (id, season_year, game_nr) VALUES (10, 2026, '1001')"
            )
            connection.execute(
                "INSERT INTO assignments (id, game_id, person_id, role, slot) "
                "VALUES (20, 10, 1, 'Zeitnehmer', 0)"
            )
            connection.execute(
                "INSERT INTO assignment_audit "
                "(id, changed_at, actor_person_id, actor_tier, action, affected_person_id, "
                "game_id, role, slot, actor_name, affected_person_name, game_snapshot) "
                "VALUES (30, '2026-01-01', 3, 'admin', 'assign', 1, 10, "
                "'Zeitnehmer', 0, 'Both References', 'Active Admin', 'snapshot')"
            )
            connection.execute(
                "INSERT INTO auth_tokens "
                "(id, nonce, purpose, person_id, issued_at, expires_at) "
                "VALUES (40, 'nonce', 'register', 2, '2026-01-01', '2026-01-02')"
            )
            connection.commit()

        _run_cli(path, "migrate-schema", "--confirm-stopped")
        with sqlite3.connect(path) as connection:
            memberships = connection.execute(
                "SELECT person_id, team_id FROM person_teams ORDER BY person_id, team_id"
            ).fetchall()
            assert memberships == [(1, 2), (2, 3), (3, 2), (3, 3), (5, 3)]
            person_columns = {
                row[1] for row in connection.execute("PRAGMA table_info(persons)")
            }
            assert "team_id" not in person_columns
            assert "desired_team_id" not in person_columns
            assert connection.execute(
                "SELECT id, is_admin, email FROM persons WHERE id=1"
            ).fetchone() == (1, 1, "admin@example.test")
            assert connection.execute(
                "SELECT account_status, email, phone FROM persons WHERE id=2"
            ).fetchone() == ("verified", None, None)
            assert connection.execute(
                "SELECT game_id, person_id FROM assignments"
            ).fetchone() == (10, 1)
            assert connection.execute(
                "SELECT actor_person_id, affected_person_id, game_id FROM assignment_audit"
            ).fetchone() == (3, 1, 10)
            assert connection.execute(
                "SELECT person_id FROM auth_tokens"
            ).fetchone() == (2,)
            assert connection.execute(
                "SELECT mv_person_id FROM teams WHERE id=3"
            ).fetchone() == (5,)
            assert connection.execute("PRAGMA foreign_key_check").fetchall() == []


def test_day_block_revision_seeds_dates_renames_current_role_and_preserves_audit_text():
    with tempfile.TemporaryDirectory() as raw_directory:
        path = Path(raw_directory) / "day-block-baseline.db"
        create_legacy_database(path)
        with sqlite3.connect(path) as connection:
            connection.execute(
                "INSERT INTO persons (id, name, is_admin, account_status) "
                "VALUES (1, 'Helper', 0, 'active')"
            )
            connection.execute(
                "INSERT INTO games (id, season_year, game_nr, date, time) "
                "VALUES (10, 2026, '1001', '05.09.2026', '10:00')"
            )
            connection.execute(
                "INSERT INTO assignments (id, game_id, person_id, role, slot) "
                "VALUES (20, 10, 1, 'Reinigung', 0)"
            )
            connection.execute(
                "INSERT INTO assignment_audit "
                "(id, changed_at, actor_tier, action, game_id, role, slot, "
                "actor_name, affected_person_name, game_snapshot) VALUES "
                "(30, '2026-01-01', 'system', 'claim', 10, 'Reinigung', 0, "
                "'System', 'Helper', 'historical snapshot')"
            )
            connection.commit()

        _run_cli(path, "migrate-schema", "--confirm-stopped")
        with sqlite3.connect(path) as connection:
            assert connection.execute(
                "SELECT phase FROM day_blocks ORDER BY phase"
            ).fetchall() == [("cleanup",), ("preparation",)]
            assert connection.execute(
                "SELECT role FROM assignments WHERE id=20"
            ).fetchone() == ("Unterstützung",)
            assert connection.execute(
                "SELECT role, game_snapshot, block_snapshot FROM assignment_audit "
                "WHERE id=30"
            ).fetchone() == ("Reinigung", "historical snapshot", None)
            audit_columns = {
                row[1] for row in connection.execute("PRAGMA table_info(assignment_audit)")
            }
            assert {"block_id", "block_snapshot"} <= audit_columns


def test_day_block_revision_fails_closed_on_role_collision_and_retains_snapshot():
    with tempfile.TemporaryDirectory() as raw_directory:
        path = Path(raw_directory) / "collision-baseline.db"
        create_legacy_database(path)
        with sqlite3.connect(path) as connection:
            connection.executemany(
                "INSERT INTO persons (id, name, is_admin, account_status) "
                "VALUES (?, ?, 0, 'active')",
                [(1, "Old"), (2, "New")],
            )
            connection.execute(
                "INSERT INTO games (id, season_year, game_nr, date) "
                "VALUES (10, 2026, '1001', '05.09.2026')"
            )
            connection.executemany(
                "INSERT INTO assignments (id, game_id, person_id, role, slot) "
                "VALUES (?, 10, ?, ?, 0)",
                [(20, 1, "Reinigung"), (21, 2, "Unterstützung")],
            )
            connection.commit()

        error = _system_exit(
            lambda: _run_cli(path, "migrate-schema", "--confirm-stopped")
        )
        message = str(error)
        assert "role collisions" in message
        assert "Retained backup:" in message
        retained = Path(message.split("Retained backup: ", 1)[1].split(". ", 1)[0])
        assert retained.exists()
        assert schema_migrations.inspect_schema(retained).kind == "baseline"


def test_fresh_head_has_no_metadata_drift_and_detects_synthetic_mismatch():
    with tempfile.TemporaryDirectory() as raw_directory:
        path = Path(raw_directory) / "drift.db"
        engine = db.make_engine(str(path))
        db.initialize_db(engine)
        assert schema_migrations.metadata_drift(engine, db.Base.metadata) == []
        with engine.begin() as connection:
            connection.exec_driver_sql(
                "CREATE TABLE synthetic_drift (id INTEGER PRIMARY KEY)"
            )
        drift = schema_migrations.metadata_drift(engine, db.Base.metadata)
        assert any(item[0] == "remove_table" for item in drift)


if __name__ == "__main__":
    h.run_all(dict(globals()))
