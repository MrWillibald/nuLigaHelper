"""Management CLI uses stable person IDs and bootstraps administrators."""

import contextlib
import io
import os
import sqlite3
import subprocess
import tempfile
from pathlib import Path

import helpers as h  # pyright: ignore[reportImplicitRelativeImport]

import backup
import db
import main
import manage_db
import webapp


def _run(path, *arguments):
    args = manage_db.build_parser().parse_args(["--db", path, *map(str, arguments)])
    output = io.StringIO()
    with contextlib.redirect_stdout(output):
        args.func(args)
    return output.getvalue()


def _expect_system_exit(callback):
    try:
        callback()
    except SystemExit as exc:
        return exc
    raise AssertionError("expected SystemExit")


def _create_items_database(path, values):
    with contextlib.closing(sqlite3.connect(path)) as connection:
        connection.execute("CREATE TABLE items (value INTEGER NOT NULL)")
        connection.executemany(
            "INSERT INTO items (value) VALUES (?)",
            [(value,) for value in values],
        )
        connection.commit()


def _item_values(path):
    with contextlib.closing(sqlite3.connect(path)) as connection:
        return [
            row[0]
            for row in connection.execute("SELECT value FROM items ORDER BY value")
        ]


def test_duplicate_names_are_discovered_and_selected_by_id():
    path = os.path.join(h._TEST_DIR, f"cli-{next(tempfile._get_candidate_names())}.db")
    engine = db.make_engine(path)
    db.init_db(engine)
    with h.Session(engine) as session:
        team = db.get_support_team(session)
        first = db.Person(name="Same Name", team=team)
        second = db.Person(name="Same Name", team=team)
        first_game = db.Game(
            season_year=h.SEASON, source_key="meeting:101", game_nr=1234,
            date="30.12.2099", time="10:00", ak="GE", home="Home", guest="Team A",
        )
        second_game = db.Game(
            season_year=h.SEASON, source_key="meeting:102", game_nr=1234,
            date="31.12.2099", time="11:00", ak="GE", home="Home", guest="Team B",
        )
        session.add_all([first, second, first_game, second_game])
        session.commit()
        first_id, second_id = first.id, second.id
        first_game_id, second_game_id = first_game.id, second_game.id

    listing = _run(path, "search-person", "Same")
    assert f"ID {first_id}" in listing and f"ID {second_id}" in listing
    games = _run(path, "--season", h.SEASON, "list-games", "--number", 1234)
    assert f"ID {first_game_id}" in games and f"ID {second_game_id}" in games
    assert "Team A" in games and "Team B" in games
    _run(
        path, "--season", h.SEASON, "assign", second_game_id,
        db.ROLE_TIMEKEEPER, second_id,
    )
    with h.Session(engine) as session:
        assignment = session.query(db.Assignment).one()
        assert assignment.person_id == second_id and assignment.game_id == second_game_id
    _run(path, "--season", h.SEASON, "set-jteam", first_game_id, "Supporter")
    with h.Session(engine) as session:
        assert session.get(db.Game, first_game_id).team_id is not None
        assert session.get(db.Game, second_game_id).team_id is None
    _run(
        path, "--season", h.SEASON, "unassign", second_game_id,
        db.ROLE_TIMEKEEPER, second_id,
    )
    with h.Session(engine) as session:
        assert session.query(db.Assignment).count() == 0


def test_grant_admin_changes_the_derived_account_fact():
    path = os.path.join(h._TEST_DIR, f"cli-{next(tempfile._get_candidate_names())}.db")
    engine = db.make_engine(path)
    db.init_db(engine)
    with h.Session(engine) as session:
        person = db.Person(name="Bootstrap", email="admin@example.test")
        session.add(person)
        session.commit()
        person_id = person.id
    _run(path, "grant-admin", person_id)
    with h.Session(engine) as session:
        assert session.get(db.Person, person_id).is_admin is True
    previous = os.environ["NULIGAHELPER_DB"]
    os.environ["NULIGAHELPER_DB"] = path
    try:
        app = webapp.create_app()
    finally:
        os.environ["NULIGAHELPER_DB"] = previous
    client = app.test_client()
    h.sign_in(client, person_id)
    assert client.get("/audit").status_code == 200


def test_contact_preflight_reports_issues_without_writing():
    path = os.path.join(h._TEST_DIR, f"cli-{next(tempfile._get_candidate_names())}.db")
    engine = db.make_engine(path)
    db.init_db(engine)
    with h.Session(engine) as session:
        session.add_all([
            db.Person(name="Changed", email=" HELPER@Example.Test "),
            db.Person(name="Collision", email="helper@example.test"),
            db.Person(name="Invalid", phone="not-a-number"),
        ])
        session.commit()
        before = [
            (person.id, person.email, person.phone)
            for person in session.query(db.Person).order_by(db.Person.id)
        ]

    output = _run(path, "contact-preflight")
    assert "CHANGED" in output
    assert "COLLISION" in output
    assert "INVALID" in output

    with h.Session(engine) as session:
        after = [
            (person.id, person.email, person.phone)
            for person in session.query(db.Person).order_by(db.Person.id)
        ]
    assert after == before, "preflight must never canonicalize or mutate records"


def test_restore_snapshot_parser_refuses_without_stopped_confirmation():
    with tempfile.TemporaryDirectory() as raw_directory:
        directory = Path(raw_directory)
        target = directory / "active.db"
        target.write_bytes(b"active database marker")
        candidate = directory / "candidate.db"

        args = manage_db.build_parser().parse_args(
            ["--db", str(target), "restore-snapshot", str(candidate)]
        )
        assert args.snapshot == str(candidate)
        assert args.confirm_stopped is False

        original_install = backup.install_snapshot
        called = []

        def unexpected_install(*_args, **_kwargs):
            called.append(True)
            raise AssertionError("restore helper must not run without confirmation")

        backup.install_snapshot = unexpected_install
        try:
            error = _expect_system_exit(
                lambda: _run(str(target), "restore-snapshot", candidate)
            )
        finally:
            backup.install_snapshot = original_install

        message = str(error)
        assert "--confirm-stopped" in message
        assert "stop all web and daily database users" in message
        assert called == []
        assert target.read_bytes() == b"active database marker"


def test_restore_snapshot_installs_standalone_wal_data_and_preserves_sidecars():
    with tempfile.TemporaryDirectory() as raw_directory:
        directory = Path(raw_directory)
        source = directory / "wal-source.db"
        writer = sqlite3.connect(source)
        try:
            assert writer.execute("PRAGMA journal_mode=WAL").fetchone()[0] == "wal"
            writer.execute("PRAGMA wal_autocheckpoint=0")
            writer.execute("CREATE TABLE items (value INTEGER NOT NULL)")
            writer.executemany("INSERT INTO items VALUES (?)", [(10,), (20,)])
            writer.commit()
            assert Path(f"{source}-wal").exists(), "committed data must reside in WAL"

            candidate = directory / "standalone-snapshot.db"
            candidate.write_bytes(backup.snapshot_database(source))
        finally:
            writer.close()

        assert not Path(f"{candidate}-wal").exists()
        assert not Path(f"{candidate}-shm").exists()

        target = directory / "active.db"
        _create_items_database(target, (99,))
        target_wal = Path(f"{target}-wal")
        target_shm = Path(f"{target}-shm")
        target_wal.write_bytes(b"stale target wal")
        target_shm.write_bytes(b"stale target shm")
        old_files = {
            path.name: path.read_bytes()
            for path in (target, target_wal, target_shm)
        }

        original_open_session = manage_db.open_session
        original_init_db = db.init_db

        def unexpected_database_initialization(*_args, **_kwargs):
            raise AssertionError("restore must not initialize or open the target database")

        manage_db.open_session = unexpected_database_initialization
        db.init_db = unexpected_database_initialization
        try:
            output = _run(
                str(target),
                "restore-snapshot",
                candidate,
                "--confirm-stopped",
            )
        finally:
            manage_db.open_session = original_open_session
            db.init_db = original_init_db

        rollback_directories = list(directory.glob(".active.db.rollback-*"))
        assert len(rollback_directories) == 1
        rollback_directory = rollback_directories[0]
        assert f"Restored database: {target.resolve()}" in output
        assert f"Rollback directory: {rollback_directory}" in output
        assert "Preserved files:" in output
        for name, contents in old_files.items():
            preserved = rollback_directory / name
            assert str(preserved) in output
            assert preserved.read_bytes() == contents

        assert not target_wal.exists(), "stale target WAL must not survive installation"
        assert not target_shm.exists(), "stale target SHM must not survive installation"
        assert _item_values(target) == [10, 20]


def test_restore_snapshot_invalid_candidate_leaves_active_set_untouched():
    with tempfile.TemporaryDirectory() as raw_directory:
        directory = Path(raw_directory)
        target = directory / "active.db"
        _create_items_database(target, (77,))
        target_wal = Path(f"{target}-wal")
        target_shm = Path(f"{target}-shm")
        target_wal.write_bytes(b"active wal")
        target_shm.write_bytes(b"active shm")
        original = {
            path: path.read_bytes()
            for path in (target, target_wal, target_shm)
        }
        invalid = directory / "invalid.db"
        invalid.write_bytes(b"not a sqlite snapshot")

        error = _expect_system_exit(
            lambda: _run(
                str(target),
                "restore-snapshot",
                invalid,
                "--confirm-stopped",
            )
        )

        message = str(error)
        assert "restore_validation" in message
        assert "active database was not changed" in message.lower()
        assert {path: path.read_bytes() for path in original} == original
        assert list(directory.glob(".active.db.rollback-*")) == []
        assert list(directory.glob(".active.db.restore-*")) == []


def test_restore_snapshot_maps_typed_errors_without_exposing_causes():
    with tempfile.TemporaryDirectory() as raw_directory:
        directory = Path(raw_directory)
        target = directory / "active.db"
        candidate = directory / "candidate.db"
        calls = []
        original_install = backup.install_snapshot

        def fail_install(snapshot, target_path, *, confirmed_stopped):
            calls.append((snapshot, target_path, confirmed_stopped))
            raise backup.RestoreError(
                backup.FailureStage.RESTORE_INSTALL,
                "installation failed with credential top-secret-value",
                cause=RuntimeError("token=another-secret-value"),
            )

        backup.install_snapshot = fail_install
        try:
            error = _expect_system_exit(
                lambda: _run(
                    str(target),
                    "restore-snapshot",
                    candidate,
                    "--confirm-stopped",
                )
            )
        finally:
            backup.install_snapshot = original_install

        assert calls == [(str(candidate), str(target.resolve()), True)]
        message = str(error)
        assert "restore_install" in message
        assert "original database set was recovered" in message.lower()
        assert "top-secret-value" not in message
        assert "another-secret-value" not in message


def test_launchers_validate_the_mandatory_secret():
    for script in ("run.sh", "run_webapp.sh"):
        path = os.path.join(h.PROJECT_DIR, script)
        assert subprocess.run(["bash", "-n", path], check=False).returncode == 0
        with open(path, encoding="utf-8") as source:
            assert "NULIGAHELPER_SECRET:?" in source.read()
    secret = os.environ.pop("NULIGAHELPER_SECRET")
    try:
        try:
            main.main()
        except RuntimeError as exc:
            assert "NULIGAHELPER_SECRET" in str(exc)
        else:
            raise AssertionError("the daily job must refuse an absent secret")
    finally:
        os.environ["NULIGAHELPER_SECRET"] = secret


if __name__ == "__main__":
    h.run_all(dict(globals()))
