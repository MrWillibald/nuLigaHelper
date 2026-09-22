"""Read-only schema inspection and guarded Alembic orchestration."""

from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from alembic import command
from alembic.autogenerate import compare_metadata
from alembic.config import Config
from alembic.migration import MigrationContext
from alembic.script import ScriptDirectory

import backup


BASELINE_REVISION = "0001_current_schema_baseline"

BASELINE_COLUMNS = {
    "assignment_audit": (
        "id", "changed_at", "actor_person_id", "actor_tier", "action",
        "affected_person_id", "game_id", "role", "slot", "actor_name",
        "affected_person_name", "game_snapshot",
    ),
    "assignments": ("id", "game_id", "person_id", "role", "slot"),
    "auth_abuse_counters": (
        "id", "action", "dimension", "subject_digest", "channel",
        "window_started_at", "count", "expires_at",
    ),
    "auth_tokens": (
        "id", "nonce", "code", "purpose", "person_id", "issued_at",
        "expires_at", "used_at",
    ),
    "games": (
        "id", "season_year", "game_nr", "day", "date", "time", "hall",
        "ak", "home", "guest", "score", "jteam", "team_id",
    ),
    "persons": (
        "id", "name", "email", "phone", "team_id", "desired_team_id",
        "is_admin", "account_status", "registered_at", "rejected_at",
        "verified_at", "approved_at",
    ),
    "teams": ("id", "name", "is_support", "mv_person_id"),
}

BASELINE_SQL_MARKERS = {
    "assignments": ("uq_game_person", "uq_game_role_slot"),
    "auth_abuse_counters": ("uq_auth_abuse_bucket",),
    "games": ("uq_season_game_nr",),
    "persons": ("unique (email)", "unique (phone)"),
    "teams": ("unique (name)",),
}


class SchemaMigrationError(RuntimeError):
    pass


@dataclass(frozen=True)
class SchemaState:
    kind: str
    revision: str | None = None
    details: tuple[str, ...] = ()


@dataclass(frozen=True)
class MigrationResult:
    database_path: Path
    backup_path: Path | None
    previous_state: SchemaState
    revision: str


def alembic_config() -> Config:
    root = Path(__file__).resolve().parent
    config = Config(str(root / "alembic.ini"))
    config.set_main_option("script_location", str(root / "migrations"))
    if config.get_main_option("sqlalchemy.url"):
        raise SchemaMigrationError("Alembic must not define an independent database URL")
    return config


def known_revisions() -> set[str]:
    scripts = ScriptDirectory.from_config(alembic_config())
    return {revision.revision for revision in scripts.walk_revisions()}


def head_revisions() -> tuple[str, ...]:
    return tuple(ScriptDirectory.from_config(alembic_config()).get_heads())


def _readonly_uri(path: Path) -> str:
    return f"{path.resolve().as_uri()}?mode=ro"


def _table_names(connection: sqlite3.Connection) -> set[str]:
    return {
        row[0]
        for row in connection.execute(
            "SELECT name FROM sqlite_master "
            "WHERE type='table' AND name NOT LIKE 'sqlite_%'"
        )
    }


def _baseline_differences(connection: sqlite3.Connection) -> list[str]:
    names = _table_names(connection)
    expected = set(BASELINE_COLUMNS)
    differences = []
    if names != expected:
        differences.append(
            f"tables expected={sorted(expected)!r} actual={sorted(names)!r}"
        )
        return differences

    for table, expected_columns in BASELINE_COLUMNS.items():
        actual_columns = tuple(
            row[1] for row in connection.execute(f'PRAGMA table_info("{table}")')
        )
        if actual_columns != expected_columns:
            differences.append(
                f"{table} columns expected={expected_columns!r} actual={actual_columns!r}"
            )
        sql_row = connection.execute(
            "SELECT sql FROM sqlite_master WHERE type='table' AND name=?", (table,)
        ).fetchone()
        normalized = " ".join((sql_row[0] if sql_row else "").lower().split())
        for marker in BASELINE_SQL_MARKERS.get(table, ()):
            if marker.lower() not in normalized:
                differences.append(f"{table} missing schema marker {marker!r}")

    indexes = {
        row[0]
        for row in connection.execute(
            "SELECT name FROM sqlite_master WHERE type='index' AND sql IS NOT NULL"
        )
    }
    if indexes != {"ix_auth_abuse_expires_at"}:
        differences.append(
            "explicit indexes expected=['ix_auth_abuse_expires_at'] "
            f"actual={sorted(indexes)!r}"
        )
    return differences


def inspect_schema(database_path: str | Path) -> SchemaState:
    """Classify a file without creating or changing it."""
    path = Path(database_path)
    if not path.exists() or path.stat().st_size == 0:
        return SchemaState("empty")
    try:
        with sqlite3.connect(_readonly_uri(path), uri=True) as connection:
            quick = connection.execute("PRAGMA quick_check").fetchall()
            if quick != [("ok",)]:
                return SchemaState("corrupt", details=(f"quick_check={quick!r}",))
            violations = connection.execute("PRAGMA foreign_key_check").fetchall()
            if violations:
                return SchemaState(
                    "invalid", details=(f"foreign_key_check={violations[:10]!r}",)
                )
            names = _table_names(connection)
            if not names:
                return SchemaState("empty")
            if "games" in names:
                game_columns = {
                    row[1] for row in connection.execute("PRAGMA table_info(games)")
                }
                if "source_key" in game_columns:
                    return SchemaState("legacy_game_identity")
            if "alembic_version" in names:
                rows = connection.execute(
                    "SELECT version_num FROM alembic_version ORDER BY version_num"
                ).fetchall()
                revisions = tuple(row[0] for row in rows)
                if len(revisions) != 1:
                    return SchemaState(
                        "divergent", details=(f"revisions={revisions!r}",)
                    )
                revision = revisions[0]
                if revision not in known_revisions():
                    return SchemaState("unknown_revision", revision=revision)
                kind = "head" if revision in head_revisions() else "versioned"
                return SchemaState(kind, revision=revision)

            differences = _baseline_differences(connection)
            if differences:
                return SchemaState("unknown_unversioned", details=tuple(differences))
            return SchemaState("baseline", revision=BASELINE_REVISION)
    except sqlite3.DatabaseError as exc:
        return SchemaState("corrupt", details=(str(exc),))


def _run_alembic(connection, operation, revision: str) -> None:
    config = alembic_config()
    config.attributes["connection"] = connection
    operation(config, revision)


def stamp_head(engine) -> None:
    """Stamp a freshly created ORM schema at the sole Alembic head."""
    heads = head_revisions()
    if len(heads) != 1:
        raise SchemaMigrationError(f"Expected one Alembic head, found {heads!r}")
    with engine.begin() as connection:
        _run_alembic(connection, command.stamp, heads[0])


def _backup_path(path: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%dT%H%M%S%f")
    return path.with_name(f"{path.name}.pre-schema-{stamp}.db")


def _create_backup(path: Path) -> Path:
    destination = _backup_path(path)
    try:
        payload = backup.snapshot_database(path)
        with destination.open("xb") as handle:
            handle.write(payload)
    except BaseException as exc:
        try:
            destination.unlink(missing_ok=True)
        except OSError:
            pass
        raise SchemaMigrationError(f"Schema backup failed: {exc}") from exc
    return destination


def migrate_to_head(database_path: str | Path, engine) -> MigrationResult:
    """Adopt an exact baseline or upgrade a known revision after a safe snapshot."""
    path = Path(database_path).resolve()
    state = inspect_schema(path)
    heads = head_revisions()
    if len(heads) != 1:
        raise SchemaMigrationError(f"Expected one Alembic head, found {heads!r}")
    head = heads[0]
    if state.kind == "head":
        return MigrationResult(path, None, state, head)
    if state.kind == "legacy_game_identity":
        raise SchemaMigrationError(
            "Legacy game identity detected. Run 'manage_db.py migrate-game-identity "
            "--confirm-stopped' before migrate-schema."
        )
    if state.kind not in {"baseline", "versioned"}:
        details = "; ".join(state.details)
        suffix = f" Details: {details}" if details else ""
        raise SchemaMigrationError(
            f"Schema state {state.kind!r} cannot be migrated automatically.{suffix}"
        )

    backup_path = _create_backup(path)
    try:
        with engine.connect() as connection:
            connection.commit()
            connection.exec_driver_sql("PRAGMA foreign_keys=OFF")
            if connection.exec_driver_sql("PRAGMA foreign_keys").scalar_one() != 0:
                raise SchemaMigrationError("Could not disable SQLite foreign keys for migration")
            try:
                if state.kind == "baseline":
                    _run_alembic(connection, command.stamp, BASELINE_REVISION)
                    connection.commit()
                _run_alembic(connection, command.upgrade, head)
                connection.commit()
            except BaseException:
                connection.rollback()
                connection.exec_driver_sql("PRAGMA foreign_keys=ON")
                raise
            connection.exec_driver_sql("PRAGMA foreign_keys=ON")
            if connection.exec_driver_sql("PRAGMA foreign_keys").scalar_one() != 1:
                raise SchemaMigrationError("Could not restore SQLite foreign keys after migration")
            violations = connection.exec_driver_sql("PRAGMA foreign_key_check").fetchall()
            if violations:
                raise SchemaMigrationError(
                    f"Post-migration foreign-key check failed: {violations[:10]!r}"
                )
    except BaseException as exc:
        if isinstance(exc, SchemaMigrationError):
            message = str(exc)
        else:
            message = f"{type(exc).__name__}: {exc}"
        raise SchemaMigrationError(
            f"Schema upgrade failed. Retained backup: {backup_path}. {message}"
        ) from exc

    final_state = inspect_schema(path)
    if final_state.kind != "head" or final_state.revision != head:
        raise SchemaMigrationError(
            f"Schema postflight failed with state {final_state.kind!r}. "
            f"Retained backup: {backup_path}"
        )
    return MigrationResult(path, backup_path, state, head)


def metadata_drift(engine, metadata) -> list:
    """Return Alembic autogenerate differences for an initialized head schema."""
    with engine.connect() as connection:
        context = MigrationContext.configure(
            connection,
            opts={"compare_type": True, "compare_server_default": True},
        )
        return compare_metadata(context, metadata)
