"""Management CLI uses stable person IDs and bootstraps administrators."""

import contextlib
import io
import os
import sqlite3
import subprocess
import tempfile
from pathlib import Path

import helpers as h  # pyright: ignore[reportImplicitRelativeImport]

from nuligahelper.ops import backup
from nuligahelper import db
from nuligahelper.jobs import daily as main
from nuligahelper import cli as manage_db
from nuligahelper import web as webapp


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
        first = db.Person(birth_date=h.ADULT_BIRTH_DATE, name="Same Name", teams=[team])
        second = db.Person(birth_date=h.ADULT_BIRTH_DATE, name="Same Name", teams=[team])
        first_game = db.Game(
            season_year=h.SEASON, game_nr="1234",
            date="30.12.2099", time="10:00", ak="BL M", home="Home", guest="Team A",
        )
        second_game = db.Game(
            season_year=h.SEASON, game_nr="1235",
            date="31.12.2099", time="11:00", ak="BL M", home="Home", guest="Team B",
        )
        session.add_all([first, second, first_game, second_game])
        session.commit()
        first_id, second_id = first.id, second.id
        first_game_id, second_game_id = first_game.id, second_game.id

    listing = _run(path, "search-person", "Same")
    assert f"ID {first_id}" in listing and f"ID {second_id}" in listing
    games = _run(path, "--season", h.SEASON, "list-games", "--number", "1234")
    assert f"ID {first_game_id}" in games and f"ID {second_game_id}" not in games
    assert "Team A" in games and "Team B" not in games
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
    error = _expect_system_exit(lambda: _run(
        path, "--season", h.SEASON, "unassign", second_game_id,
        db.ROLE_TIMEKEEPER, second_id,
    ))
    assert "verbindlich" in str(error), "CLI identity selection cannot bypass binding duties"
    with h.Session(engine) as session:
        assignment = session.query(db.Assignment).one()
        assert assignment.person_id == second_id and assignment.game_id == second_game_id
        assert session.query(db.AssignmentAudit).filter_by(action="release").count() == 0


def test_grant_admin_changes_the_derived_account_fact():
    path = os.path.join(h._TEST_DIR, f"cli-{next(tempfile._get_candidate_names())}.db")
    engine = db.make_engine(path)
    db.init_db(engine)
    with h.Session(engine) as session:
        person = db.Person(birth_date=h.ADULT_BIRTH_DATE, name="Bootstrap", email="admin@example.test")
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
            db.Person(birth_date=h.ADULT_BIRTH_DATE, name="Changed", email=" HELPER@Example.Test "),
            db.Person(birth_date=h.ADULT_BIRTH_DATE, name="Collision", email="helper@example.test"),
            db.Person(birth_date=h.ADULT_BIRTH_DATE, name="Invalid", phone="not-a-number"),
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


def test_membership_commands_use_person_ids_replace_sets_and_clear_mv():
    path = os.path.join(h._TEST_DIR, f"cli-{next(tempfile._get_candidate_names())}.db")
    engine = db.make_engine(path)
    db.initialize_db(engine)
    with h.Session(engine) as session:
        first_team = db.Team(name="First Team")
        second_team = db.Team(name="Second Team")
        target = db.Person(birth_date=h.ADULT_BIRTH_DATE, name="Duplicate", teams=[first_team])
        untouched = db.Person(birth_date=h.ADULT_BIRTH_DATE, name="Duplicate", teams=[second_team])
        session.add_all([first_team, second_team, target, untouched])
        session.flush()
        first_team.mv_person_id = target.id
        session.commit()
        target_id, untouched_id = target.id, untouched.id

    output = _run(path, "set-memberships", target_id, "Second Team", "First Team", "Second Team")
    assert "First Team, Second Team" in output
    shown = _run(path, "show-memberships", target_id)
    assert f"ID {target_id}" in shown and "First Team, Second Team" in shown
    _run(path, "set-memberships", target_id, "Second Team")
    with h.Session(engine) as session:
        target = session.get(db.Person, target_id)
        untouched = session.get(db.Person, untouched_id)
        assert db.membership_team_names(target) == ("Second Team",)
        assert db.membership_team_names(untouched) == ("Second Team",)
        assert session.query(db.Team).filter_by(name="First Team").one().mv_person_id is None

    error = _expect_system_exit(
        lambda: _run(path, "set-memberships", target_id, "Missing Team")
    )
    assert "Unknown team" in str(error)
    with h.Session(engine) as session:
        assert db.membership_team_names(session.get(db.Person, target_id)) == ("Second Team",)


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


def test_cli_ordner_and_legacy_alias_target_one_stable_role_and_preserve_binding():
    path = os.path.join(h._TEST_DIR, f"cli-ordner-{next(tempfile._get_candidate_names())}.db")
    engine = db.make_engine(path)
    db.initialize_db(engine)
    with h.Session(engine) as session:
        helper = db.Person(name="Steward", birth_date=h.ADULT_BIRTH_DATE)
        games = [db.Game(season_year=h.SEASON, game_nr=str(9300 + index),
                         date="31.12.2099", ak="BL M") for index in range(2)]
        session.add_all([helper, *games])
        session.commit()
        person_id, game_ids = helper.id, [game.id for game in games]
    for game_id, spelling in zip(game_ids, ("Ordner", "Ordnungsdienst")):
        assert _run(path, "assign", game_id, spelling, person_id).endswith("as Ordner\n")
        refusal = _expect_system_exit(lambda: _run(path, "unassign", game_id, spelling, person_id))
        assert "verbindlich" in str(refusal), "aliases cannot authorize ordinary release"
    listing = _run(path, "--season", h.SEASON, "list-games")
    assert listing.count("Ordner: Steward") == 2 and "Ordnungsdienst" not in listing
    with h.Session(engine) as session:
        assert [a.role for a in session.query(db.Assignment)] == ["Ordnungsdienst"] * 2
        assert [a.role for a in session.query(db.AssignmentAudit)] == ["Ordner"] * 2
    parser = manage_db.build_parser()
    with contextlib.redirect_stderr(io.StringIO()) as errors:
        assert _expect_system_exit(lambda: parser.parse_args(["assign", "1", "invalid", "1"])).code == 2
    assert "Ordner" in errors.getvalue() and "Ordnungsdienst" not in errors.getvalue()
    engine.dispose()


def test_cli_security_and_adult_timing_follow_revised_minimums_without_override():
    from datetime import date
    path = os.path.join(h._TEST_DIR, f"cli-ages-{next(tempfile._get_candidate_names())}.db")
    engine = db.make_engine(path)
    db.initialize_db(engine)
    with h.Session(engine) as session:
        game = db.Game(season_year=h.SEASON, game_nr="9201", date="01.11.2026", ak="BL M")
        people = [db.Person(name="Timing sixteen", birth_date=date(2010, 11, 1)),
                  db.Person(name="Security seventeen", birth_date=date(2008, 11, 2)),
                  db.Person(name="Security unknown"),
                  db.Person(name="Security eighteen", birth_date=date(2008, 11, 1))]
        session.add_all([game, *people])
        session.commit()
        game_id, ids = game.id, [person.id for person in people]
    for person_id in ids[1:3]:
        error = _expect_system_exit(lambda: _run(path, "assign", game_id, db.ROLE_SECURITY, person_id))
        assert "Ordner" in str(error)
        with h.Session(engine) as session:
            assert session.query(db.Assignment).count() == 0
            assert session.query(db.AssignmentAudit).count() == 0
    _run(path, "assign", game_id, db.ROLE_TIMEKEEPER, ids[0])
    _run(path, "assign", game_id, db.ROLE_SECURITY, ids[3])
    with h.Session(engine) as session:
        assert session.query(db.Assignment).count() == 2
        assert session.query(db.AssignmentAudit).filter_by(action="claim").count() == 2


if __name__ == "__main__":
    h.run_all(dict(globals()))
