# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# SQLite engine, connection PRAGMAs, explicit initialization, startup
# verification and the bounded write-retry policy.
# ---------------------------------------------------------------

from __future__ import annotations

import logging
import os
import time

from sqlalchemy import create_engine, event
from sqlalchemy.exc import OperationalError
from sqlalchemy.orm import Session, sessionmaker

from db.errors import (
    AssignmentTemporarilyUnavailableError,
    DatabaseMigrationRequiredError,
    SQLiteInitializationError,
)
from db.models import SUPPORT_TEAM_NAME, Base, ScheduleVisibilitySetting, Team
from db.people import get_support_team


# Repository root; relative database paths resolve here, as before the split.
PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

DEFAULT_DB_PATH = "nuliga_helper.db"
SQLITE_TIMEOUT_SECONDS = 5.0
SQLITE_BUSY_TIMEOUT_MS = 5000
SQLITE_SYNCHRONOUS_FULL = 2
GAME_IDENTITY_SCHEMA_VERSION = 2


def resolve_db_path(db_path: str) -> str:
    """Resolve a relative DB path against the project directory."""
    if os.path.isabs(db_path):
        return db_path
    return os.path.join(PROJECT_DIR, db_path)


def _sqlite_pragma_value(dbapi_connection, pragma: str):
    cursor = dbapi_connection.cursor()
    try:
        cursor.execute(f"PRAGMA {pragma}")
        row = cursor.fetchone()
        return row[0] if row else None
    finally:
        cursor.close()


def _configure_sqlite_connection(dbapi_connection, _connection_record=None) -> None:
    cursor = dbapi_connection.cursor()
    try:
        cursor.execute("PRAGMA foreign_keys=ON")
        cursor.execute(f"PRAGMA busy_timeout={SQLITE_BUSY_TIMEOUT_MS}")
        cursor.execute("PRAGMA synchronous=FULL")
    finally:
        cursor.close()

    actual = {
        "foreign_keys": _sqlite_pragma_value(dbapi_connection, "foreign_keys"),
        "busy_timeout": _sqlite_pragma_value(dbapi_connection, "busy_timeout"),
        "synchronous": _sqlite_pragma_value(dbapi_connection, "synchronous"),
    }
    expected = {
        "foreign_keys": 1,
        "busy_timeout": SQLITE_BUSY_TIMEOUT_MS,
        "synchronous": SQLITE_SYNCHRONOUS_FULL,
    }
    mismatches = [
        f"{name}={actual[name]!r} (expected {expected[name]!r})"
        for name in expected
        if actual[name] != expected[name]
    ]
    if mismatches:
        raise SQLiteInitializationError(
            "SQLite connection safety settings could not be verified: "
            + ", ".join(mismatches)
        )


def inspect_sqlite_runtime(connection) -> dict[str, int | str | None]:
    """Return the SQLite settings used by startup checks and focused tests."""
    return {
        "foreign_keys": connection.exec_driver_sql(
            "PRAGMA foreign_keys"
        ).scalar_one_or_none(),
        "busy_timeout": connection.exec_driver_sql(
            "PRAGMA busy_timeout"
        ).scalar_one_or_none(),
        "synchronous": connection.exec_driver_sql(
            "PRAGMA synchronous"
        ).scalar_one_or_none(),
        "journal_mode": connection.exec_driver_sql(
            "PRAGMA journal_mode"
        ).scalar_one_or_none(),
    }


def make_engine(db_path: str = DEFAULT_DB_PATH):
    """Create an engine with the supported SQLite runtime profile."""
    database_url = (
        "sqlite:///:memory:"
        if db_path == ":memory:"
        else f"sqlite:///{resolve_db_path(db_path)}"
    )
    engine = create_engine(
        database_url,
        connect_args={"timeout": SQLITE_TIMEOUT_SECONDS},
    )
    event.listen(engine, "connect", _configure_sqlite_connection)
    return engine


def _establish_sqlite_runtime(engine) -> None:
    """Establish and verify the writable SQLite runtime profile."""
    try:
        with engine.connect() as connection:
            _configure_sqlite_connection(connection.connection.driver_connection)
            mode = connection.exec_driver_sql("PRAGMA journal_mode=WAL").scalar_one()
            if str(mode).lower() != "wal":
                raise SQLiteInitializationError(
                    "SQLite startup could not establish WAL journal mode "
                    f"(database reported {mode!r}). Ensure the database is writable "
                    "and stored on a local filesystem."
                )
            settings = inspect_sqlite_runtime(connection)
            expected = {
                "foreign_keys": 1,
                "busy_timeout": SQLITE_BUSY_TIMEOUT_MS,
                "synchronous": SQLITE_SYNCHRONOUS_FULL,
            }
            mismatches = [
                f"{name}={settings[name]!r} (expected {value!r})"
                for name, value in expected.items()
                if settings[name] != value
            ]
            if mismatches:
                raise SQLiteInitializationError(
                    "SQLite startup safety checks failed: " + ", ".join(mismatches)
                )
    except SQLiteInitializationError:
        raise
    except OperationalError as exc:
        raise SQLiteInitializationError(
            "SQLite startup failed while establishing and verifying WAL mode. "
            "Ensure the database is writable, not held by another startup, and "
            "stored on a local filesystem."
        ) from exc


def _validate_database(engine) -> None:
    """Run post-initialization/runtime relational integrity checks."""

    try:
        with engine.connect() as connection:
            violations = connection.exec_driver_sql(
                "PRAGMA foreign_key_check"
            ).fetchall()
    except OperationalError as exc:
        raise SQLiteInitializationError(
            "SQLite startup could not run foreign_key_check. Ensure the database "
            "is readable and not locked, then retry."
        ) from exc
    if violations:
        details = "; ".join(
            f"table={table!r}, rowid={rowid!r}, parent={parent!r}, fk_index={fk_id!r}"
            for table, rowid, parent, fk_id in violations
        )
        raise SQLiteInitializationError(
            "SQLite foreign_key_check failed: " + details
            + ". Repair the listed relationships or restore a valid backup."
        )


def _engine_database_path(engine) -> str | None:
    database = engine.url.database
    if not database or database == ":memory:":
        return None
    return resolve_db_path(database)


def initialize_db(engine) -> None:
    """Create and stamp a new database; refuse every populated target."""
    import schema_migrations

    database_path = _engine_database_path(engine)
    if database_path is not None:
        state = schema_migrations.inspect_schema(database_path)
        if state.kind != "empty":
            raise SQLiteInitializationError(
                f"Database initialization requires an absent or empty target; "
                f"found schema state {state.kind!r}."
            )

    _establish_sqlite_runtime(engine)
    Base.metadata.create_all(engine)
    schema_migrations.stamp_head(engine)
    with engine.begin() as connection:
        connection.exec_driver_sql(
            f"PRAGMA user_version={GAME_IDENTITY_SCHEMA_VERSION}"
        )
    _validate_database(engine)

    session = Session(engine)
    try:
        session.add(ScheduleVisibilitySetting(id=1, hide_from=None))
        if get_support_team(session) is None:
            session.add(Team(name=SUPPORT_TEAM_NAME, is_support=True))
            logging.info(f"Support team '{SUPPORT_TEAM_NAME}' created")
        session.commit()
    finally:
        session.close()

    if database_path is not None:
        state = schema_migrations.inspect_schema(database_path)
        if state.kind != "head":
            raise SQLiteInitializationError(
                f"Fresh database validation expected Alembic head, found {state.kind!r}."
            )


def verify_db(engine) -> None:
    """Verify an existing head database without creating or upgrading schema."""
    import schema_migrations

    database_path = _engine_database_path(engine)
    if database_path is None:
        raise SQLiteInitializationError(
            "Runtime verification requires a file-backed SQLite database."
        )
    state = schema_migrations.inspect_schema(database_path)
    if state.kind == "legacy_game_identity":
        raise DatabaseMigrationRequiredError(
            "Die Datenbank verwendet die alte Spielidentität. Diese Version "
            "enthält deren Umstellung nicht mehr; führe sie zuerst mit einem "
            "früheren Release durch."
        )
    if state.kind == "invalid":
        details = "; ".join(state.details)
        raise SQLiteInitializationError(
            "SQLite foreign_key_check failed: " + details
            + ". Repair the listed relationships or restore a valid backup."
        )
    if state.kind != "head":
        details = "; ".join(state.details)
        suffix = f" Details: {details}" if details else ""
        raise SQLiteInitializationError(
            f"Database schema is {state.kind!r}, not the required Alembic head. "
            "Stop database users and run 'manage_db.py migrate-schema "
            f"--confirm-stopped'.{suffix}"
        )
    try:
        schema_migrations.verify_head_schema(engine)
    except schema_migrations.SchemaMigrationError as exc:
        raise SQLiteInitializationError(
            f"Database schema verification failed: {exc}. "
            "Stop database users and restore a recognized snapshot or review the schema."
        ) from exc
    _establish_sqlite_runtime(engine)
    _validate_database(engine)


# Backward-compatible name for explicit test/bootstrap callers. Runtime entry
# points deliberately use verify_db() instead.
init_db = initialize_db


def make_session_factory(engine):
    return sessionmaker(bind=engine, expire_on_commit=False)


def _is_sqlite_contention(exc: OperationalError) -> bool:
    original = exc.orig
    error_code = getattr(original, "sqlite_errorcode", None)
    if isinstance(error_code, int) and error_code & 0xFF in (5, 6):
        return True
    message = str(original).lower()
    return "database is locked" in message or "database is busy" in message


def _wait_for_assignment_retry(deadline: float, exc: OperationalError) -> None:
    if not _is_sqlite_contention(exc):
        raise exc
    remaining = deadline - time.monotonic()
    if remaining <= 0:
        raise AssignmentTemporarilyUnavailableError() from exc
    time.sleep(min(0.05, remaining))


def database_ready(database_path: str) -> bool:
    """Read the existing SQLite schema without creating a missing database."""
    import sqlite3
    from pathlib import Path
    try:
        uri = Path(database_path).resolve().as_uri() + '?mode=ro'
        with sqlite3.connect(uri, uri=True, timeout=1) as connection:
            connection.execute('PRAGMA query_only=ON')
            # Resolve essential schema even for a freshly initialized empty roster.
            connection.execute('SELECT id FROM games LIMIT 0')
            connection.execute('SELECT id FROM persons LIMIT 0')
            connection.execute('SELECT id FROM auth_tokens LIMIT 0')
        return True
    except (sqlite3.Error, OSError, ValueError):
        return False
