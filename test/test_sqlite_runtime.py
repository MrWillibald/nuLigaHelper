"""Focused checks for the production SQLite connection and startup profile."""

import os
import sqlite3
import tempfile

import helpers as h
import db
from sqlalchemy.exc import IntegrityError


_TEST_DIR = tempfile.mkdtemp(prefix="nuligahelper-sqlite-runtime-")


def _database_path(name: str) -> str:
    return os.path.join(_TEST_DIR, name)


def test_every_factory_connection_reports_the_runtime_profile():
    path = _database_path("runtime.db")
    first_engine = db.make_engine(path)
    db.init_db(first_engine)
    second_engine = db.make_engine(path)
    try:
        with first_engine.connect() as first, second_engine.connect() as second:
            for settings in (
                db.inspect_sqlite_runtime(first),
                db.inspect_sqlite_runtime(second),
            ):
                assert settings == {
                    "foreign_keys": 1,
                    "busy_timeout": 5000,
                    "synchronous": 2,
                    "journal_mode": "wal",
                }, settings
    finally:
        first_engine.dispose()
        second_engine.dispose()


def test_startup_rejects_a_database_that_cannot_enter_wal():
    engine = db.make_engine(":memory:")
    try:
        try:
            db.init_db(engine)
        except db.SQLiteInitializationError as exc:
            message = str(exc)
            assert "WAL journal mode" in message and "memory" in message
            assert "local filesystem" in message
        else:
            raise AssertionError("startup must reject a non-WAL journal mode")
    finally:
        engine.dispose()


def test_factory_connection_enforces_foreign_keys():
    engine = db.make_engine(_database_path("foreign-keys.db"))
    db.init_db(engine)
    try:
        with engine.connect() as connection:
            transaction = connection.begin()
            try:
                connection.exec_driver_sql(
                    "INSERT INTO assignments "
                    "(game_id, person_id, role, slot) VALUES (999, 999, 'Verkauf', 0)"
                )
            except IntegrityError:
                transaction.rollback()
            else:
                transaction.rollback()
                raise AssertionError("configured connections must reject orphan rows")
    finally:
        engine.dispose()


def test_startup_reports_preexisting_foreign_key_violations_actionably():
    path = _database_path("orphan.db")
    engine = db.make_engine(path)
    db.init_db(engine)
    engine.dispose()

    raw = sqlite3.connect(path)
    try:
        raw.execute("PRAGMA foreign_keys=OFF")
        raw.execute(
            "INSERT INTO assignments "
            "(game_id, person_id, role, slot) VALUES (999, 999, 'Verkauf', 0)"
        )
        raw.commit()
    finally:
        raw.close()

    checked_engine = db.make_engine(path)
    try:
        try:
            db.init_db(checked_engine)
        except db.SQLiteInitializationError as exc:
            message = str(exc)
            assert "foreign_key_check" in message
            assert "assignments" in message and "Repair" in message
        else:
            raise AssertionError("startup must reject pre-existing orphan rows")
    finally:
        checked_engine.dispose()


if __name__ == "__main__":
    h.run_all(dict(globals()))
