"""Controlled migration tests for canonical game-number identity."""

import contextlib
import io
import os
import sqlite3
import tempfile
from pathlib import Path

import helpers as h

import db
import game_identity_migration as migration
import manage_db


LEGACY_SCHEMA = """
CREATE TABLE teams (
    id INTEGER PRIMARY KEY, name VARCHAR(120) UNIQUE NOT NULL,
    is_support BOOLEAN NOT NULL DEFAULT 0, mv_person_id INTEGER
);
CREATE TABLE persons (
    id INTEGER PRIMARY KEY, name VARCHAR(120) NOT NULL
);
CREATE TABLE games (
    id INTEGER PRIMARY KEY, season_year INTEGER NOT NULL,
    source_key VARCHAR(200) NOT NULL, game_nr INTEGER NOT NULL,
    day VARCHAR(20), date VARCHAR(20), time VARCHAR(30), hall INTEGER,
    ak VARCHAR(10), home VARCHAR(120), guest VARCHAR(120), score VARCHAR(120),
    jteam VARCHAR(120), team_id INTEGER,
    CONSTRAINT uq_season_source_key UNIQUE (season_year, source_key),
    FOREIGN KEY(team_id) REFERENCES teams(id)
);
CREATE TABLE assignments (
    id INTEGER PRIMARY KEY, game_id INTEGER NOT NULL, person_id INTEGER NOT NULL,
    role VARCHAR(40) NOT NULL, slot INTEGER NOT NULL,
    CONSTRAINT uq_game_person UNIQUE (game_id, person_id),
    CONSTRAINT uq_game_role_slot UNIQUE (game_id, role, slot),
    FOREIGN KEY(game_id) REFERENCES games(id),
    FOREIGN KEY(person_id) REFERENCES persons(id)
);
CREATE TABLE assignment_audit (
    id INTEGER PRIMARY KEY, changed_at DATETIME NOT NULL,
    actor_person_id INTEGER, actor_tier VARCHAR(20) NOT NULL,
    action VARCHAR(20) NOT NULL, affected_person_id INTEGER, game_id INTEGER,
    role VARCHAR(40) NOT NULL, slot INTEGER NOT NULL,
    actor_name VARCHAR(120) NOT NULL, affected_person_name VARCHAR(120) NOT NULL,
    game_snapshot VARCHAR(300) NOT NULL,
    FOREIGN KEY(game_id) REFERENCES games(id)
);
"""


def _legacy_database(path, *, duplicate_ordinary=False, conflicting_spf=False):
    connection = sqlite3.connect(path)
    connection.executescript(LEGACY_SCHEMA)
    connection.execute("INSERT INTO teams VALUES (1, 'Supporter', 1, NULL)")
    connection.execute("INSERT INTO teams VALUES (2, 'Responsible', 0, NULL)")
    connection.execute("INSERT INTO persons VALUES (1, 'Helper')")
    connection.execute(
        "INSERT INTO games VALUES (1, 2026, 'meeting:ordinary', 1001, "
        "'Sa', '05.09.2026', '15:00', 280340, 'BL mD', 'Home', 'Guest', '', NULL, NULL)"
    )
    next_id = 2
    for date_text in ("28.11.2026", "12.12.2026"):
        for number, time_text in zip(
            (1, 3, 4, 6, 7, 9, 10),
            ("09:00", "09:40", "10:00", "10:40", "11:00", "11:40", "12:00"),
        ):
            team_id = 2 if next_id == 3 else (1 if conflicting_spf and next_id == 4 else None)
            connection.execute(
                "INSERT INTO games VALUES (?, 2026, ?, ?, 'Sa', ?, ?, 280345, "
                "'SPF Mini', 'Home', 'Guest', '', NULL, ?)",
                (next_id, f"fallback:{next_id}", number, date_text, time_text, team_id),
            )
            next_id += 1
    connection.execute(
        "INSERT INTO assignments VALUES (1, 3, 1, 'Zeitnehmer', 0)"
    )
    connection.execute(
        "INSERT INTO assignment_audit VALUES "
        "(1, '2026-01-01', NULL, 'cli', 'claim', 1, 3, 'Zeitnehmer', 0, "
        "'System', 'Helper', 'legacy snapshot')"
    )
    if duplicate_ordinary:
        connection.execute(
            "INSERT INTO games VALUES (99, 2026, 'meeting:duplicate', 1001, "
            "'Sa', '06.09.2026', '16:00', 280340, 'BL mD', 'Home', 'Other', '', NULL, NULL)"
        )
    connection.commit()
    connection.close()


def test_preflight_reports_conflicts_without_writing():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "legacy.db"
        _legacy_database(path, duplicate_ordinary=True)
        before = path.read_bytes()
        try:
            migration.preflight(path)
        except migration.MigrationError as exc:
            assert "Spielnummern" in str(exc) and "1001" in str(exc)
        else:
            raise AssertionError("ordinary duplicates require manual reconciliation")
        assert path.read_bytes() == before


def test_preflight_rejects_conflicting_spf_responsibility_without_writing():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "legacy.db"
        _legacy_database(path, conflicting_spf=True)
        before = path.read_bytes()
        try:
            migration.preflight(path)
        except migration.MigrationError as exc:
            assert "Spielfest-Verantwortung" in str(exc)
        else:
            raise AssertionError("conflicting SPF teams must stop migration")
        assert path.read_bytes() == before


def test_migration_collapses_spf_and_preserves_ids_relationships_and_audit():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "legacy.db"
        _legacy_database(path)
        result = migration.migrate(path)
        assert result.backup_path.exists()
        assert result.plan.legacy_game_count == 15
        assert result.plan.resulting_game_count == 3
        assert result.plan.redundant_spielfest_rows == 12
        assert migration.schema_state(path) == "current"

        with sqlite3.connect(path) as connection:
            columns = {
                row[1]: row[2] for row in connection.execute("PRAGMA table_info(games)")
            }
            assert "source_key" not in columns and columns["game_nr"] == "VARCHAR(100)"
            games = connection.execute(
                "SELECT id, game_nr, home, guest FROM games ORDER BY id"
            ).fetchall()
            assert games == [
                (1, "1001", "Home", "Guest"),
                (3, "SPF:2026-11-28:spf mini", "Spielfest", ""),
                (9, "SPF:2026-12-12:spf mini", "Spielfest", ""),
            ]
            assert connection.execute(
                "SELECT game_id FROM assignments"
            ).fetchone()[0] == 3
            assert connection.execute(
                "SELECT game_id, game_snapshot FROM assignment_audit"
            ).fetchone() == (3, "legacy snapshot")
            assert connection.execute("PRAGMA user_version").fetchone()[0] == 2
            assert connection.execute("PRAGMA foreign_key_check").fetchall() == []


def test_legacy_startup_fails_closed_with_migration_instruction():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "legacy.db"
        _legacy_database(path)
        engine = db.make_engine(str(path))
        try:
            db.verify_db(engine)
        except db.DatabaseMigrationRequiredError as exc:
            assert "migrate-game-identity" in str(exc)
        else:
            raise AssertionError("legacy startup must require explicit migration")


def test_cli_requires_stopped_confirmation_then_reports_backup():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "legacy.db"
        _legacy_database(path)
        args = manage_db.build_parser().parse_args([
            "--db", str(path), "migrate-game-identity"
        ])
        try:
            args.func(args)
        except SystemExit as exc:
            assert "--confirm-stopped" in str(exc)
        else:
            raise AssertionError("migration must require stopped-writer confirmation")

        args = manage_db.build_parser().parse_args([
            "--db", str(path), "migrate-game-identity", "--confirm-stopped"
        ])
        output = io.StringIO()
        with contextlib.redirect_stdout(output):
            args.func(args)
        text = output.getvalue()
        assert "Migrated database:" in text and "Backup:" in text
        assert "15 -> 3" in text and migration.schema_state(path) == "current"


def test_active_writer_blocks_migration_without_changing_legacy_schema():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "legacy.db"
        _legacy_database(path)
        writer = sqlite3.connect(path)
        try:
            writer.execute("BEGIN IMMEDIATE")
            writer.execute("UPDATE persons SET name='Locked' WHERE id=1")
            try:
                migration.migrate(path)
            except migration.MigrationError as exc:
                assert "gestoppt" in str(exc)
            else:
                raise AssertionError("an active writer must block migration")
        finally:
            writer.rollback()
            writer.close()
        assert migration.schema_state(path) == "legacy"


def test_interrupted_table_rebuild_rolls_back_and_can_be_retried():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "legacy.db"
        _legacy_database(path)
        original_build_plan = migration._build_plan
        calls = []

        def fail_during_rebuild(connection):
            plan, rows = original_build_plan(connection)
            calls.append(True)
            if len(calls) == 2:
                rows.append(dict(rows[0]))  # force a primary-key failure mid-transaction
            return plan, rows

        migration._build_plan = fail_during_rebuild
        try:
            try:
                migration.migrate(path)
            except sqlite3.IntegrityError:
                pass
            else:
                raise AssertionError("synthetic rebuild interruption must fail")
        finally:
            migration._build_plan = original_build_plan

        assert migration.schema_state(path) == "legacy"
        with sqlite3.connect(path) as connection:
            assert connection.execute("SELECT COUNT(*) FROM games").fetchone()[0] == 15
            assert connection.execute(
                "SELECT COUNT(*) FROM sqlite_master WHERE name='games_new'"
            ).fetchone()[0] == 0
        assert migration.migrate(path).plan.resulting_game_count == 3


if __name__ == "__main__":
    h.run_all(dict(globals()))
