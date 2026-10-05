"""Read-only schema inspection and guarded Alembic orchestration."""

from __future__ import annotations

import sqlite3
import hashlib
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

from alembic import command
from alembic.autogenerate import compare_metadata
from alembic.config import Config
from alembic.migration import MigrationContext
from alembic.script import ScriptDirectory
from sqlalchemy import CheckConstraint, MetaData, UniqueConstraint, inspect

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

# Previously recognized manual baselines can contain these exact defaults.
# Preserve them when present; adding or changing one is still unreviewed drift.
HISTORICAL_DEFAULTS = {
    ("persons", "is_admin"): "0",
    ("persons", "account_status"): "'active'",
    ("teams", "is_support"): "0",
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


def _retained_data(path: Path, columns: dict | None = None) -> tuple[dict, dict]:
    """Fingerprint preserved values without exposing contacts or birth dates."""
    with sqlite3.connect(_readonly_uri(path), uri=True) as connection:
        tables = _table_names(connection) - {"alembic_version"}
        if columns is None:
            columns = {
                table: tuple(row[1] for row in connection.execute(
                    f'PRAGMA table_info("{table}")'
                ) if not (table == "persons" and row[1] in {"team_id", "desired_team_id"}))
                for table in tables
            }
        fingerprints = {}
        for table, selected_columns in columns.items():
            if table == "person_teams":
                continue
            if table not in tables:
                raise SchemaMigrationError(f"Retained table is missing: {table}")
            selected = ", ".join(
                "CASE WHEN role='Reinigung' THEN 'Unterstützung' ELSE role END"
                if table == "assignments" and column == "role" else f'"{column}"'
                for column in selected_columns
            )
            # This reviewed revision adds rows to an existing table. Compare
            # every original block, while checking the newly seeded cake rows
            # separately inside its revision. Never ignore old block values.
            clause = (" WHERE phase <> 'cake_delivery'"
                      if table == "day_blocks" and "cake_quantity" not in selected_columns
                      else "")
            rows = [tuple(row) for row in connection.execute(
                f'SELECT {selected} FROM "{table}"{clause}'
            )]
            fingerprints[table] = hashlib.sha256(
                repr(sorted(rows, key=repr)).encode("utf-8")
            ).hexdigest()
        if "person_teams" in tables:
            memberships = connection.execute(
                "SELECT person_id, team_id FROM person_teams"
            ).fetchall()
        else:
            person_columns = {row[1] for row in connection.execute("PRAGMA table_info(persons)")}
            legacy = [column for column in ("team_id", "desired_team_id")
                      if column in person_columns]
            memberships = connection.execute(" UNION ".join(
                f"SELECT id, {column} FROM persons WHERE {column} IS NOT NULL"
                for column in legacy
            )).fetchall() if legacy else []
        fingerprints["person_teams"] = hashlib.sha256(
            repr(sorted(tuple(row) for row in memberships)).encode("utf-8")
        ).hexdigest()
        return columns, fingerprints


def _historical_defaults(engine) -> dict:
    found = {}
    with engine.connect() as connection:
        for (table, column), allowed in HISTORICAL_DEFAULTS.items():
            rows = connection.exec_driver_sql(f'PRAGMA table_info("{table}")')
            actual = next((row[4] for row in rows if row[1] == column), None)
            if actual == allowed:
                found[(table, column)] = actual
    return found


def _unreviewed_metadata_drift(engine, metadata, inherited_defaults: dict) -> list:
    """Use SQLite's actual constraints to exclude reflection false positives."""
    differences = metadata_drift(engine, metadata)
    actual_defaults = _historical_defaults(engine)
    remaining = []
    with engine.connect() as connection:
        for difference in differences:
            if isinstance(difference, list):
                unreviewed = []
                for item in difference:
                    key = (item[2], item[3])
                    if not (item[0] == "modify_default" and key in inherited_defaults
                            and actual_defaults.get(key) == inherited_defaults[key]):
                        unreviewed.append(item)
                if unreviewed:
                    remaining.append(unreviewed)
                continue
            if difference[0] == "add_constraint" and isinstance(difference[1], UniqueConstraint):
                constraint = difference[1]
                columns = tuple(column.name for column in constraint.columns)
                table = constraint.table.name
                indexes = connection.exec_driver_sql(f'PRAGMA index_list("{table}")').fetchall()
                existing_unique = []
                for index in indexes:
                    if index[2] == 1 and index[3] == "u" and index[4] == 0:
                        existing_unique.append(tuple(row[2] for row in connection.exec_driver_sql(
                            f'PRAGMA index_info("{index[1]}")'
                        )))
                # SQLite reflection misses multiline and inline unique clauses.
                # Only an actual whole-table UNIQUE constraint can satisfy it.
                if columns in existing_unique:
                    continue
            remaining.append(difference)
    return remaining


def verify_head_schema(engine, *, inherited_defaults: dict | None = None) -> None:
    """Verify actual current schema, including an already-head database, read-only."""
    try:
        with engine.connect() as connection:
            birth_column = next((row for row in connection.exec_driver_sql(
                "PRAGMA table_info(persons)"
            ) if row[1] == "birth_date"), None)
            if (birth_column is None or str(birth_column[2]).upper() != "DATE"
                    or birth_column[3] != 0 or birth_column[4] is not None):
                raise SchemaMigrationError("Schema verification found invalid birth-date storage")
        actual_defaults = _historical_defaults(engine)
        if inherited_defaults is None:
            inherited_defaults = actual_defaults
        elif actual_defaults != inherited_defaults:
            raise SchemaMigrationError("Schema verification changed recognized historical defaults")
        import db
        if _unreviewed_metadata_drift(engine, db.Base.metadata, inherited_defaults):
            raise SchemaMigrationError("Schema verification found unreviewed model metadata drift")
        if _check_constraint_drift(engine, db.Base.metadata):
            raise SchemaMigrationError("Schema verification found unreviewed check-constraint drift")
    except SchemaMigrationError:
        raise
    except Exception as exc:
        # Keep raw data out of routine startup diagnostics even if reflection
        # cannot interpret a malformed schema.
        raise SchemaMigrationError(
            f"Schema verification could not inspect the database ({type(exc).__name__})"
        ) from exc


def _check_constraint_drift(engine, metadata) -> bool:
    """Alembic's metadata comparison omits CHECKs, so verify them explicitly."""
    def normalized(expression):
        return "".join(str(expression).lower().split())

    inspector = inspect(engine)
    for table in metadata.tables.values():
        expected = {constraint.name: normalized(constraint.sqltext)
                    for constraint in table.constraints
                    if isinstance(constraint, CheckConstraint)}
        actual = {constraint["name"]: normalized(constraint["sqltext"])
                  for constraint in inspector.get_check_constraints(table.name)}
        if actual != expected:
            return True
    return False


def _verify_pre_cake_schema(engine, revision: str, inherited_defaults: dict) -> None:
    """Verify the two reviewed predecessors against their actual storage shape."""
    import db
    metadata = MetaData()
    for table in db.Base.metadata.tables.values():
        table.to_metadata(metadata)
    blocks = metadata.tables["day_blocks"]
    for name in ("cake_quantity", "delivery_time"):
        blocks._columns.remove(blocks.c[name])
    for constraint in list(blocks.constraints):
        if isinstance(constraint, CheckConstraint):
            blocks.constraints.remove(constraint)
    blocks.append_constraint(CheckConstraint(
        "phase IN ('preparation', 'cleanup')", name="ck_day_block_phase"
    ))
    assignments = metadata.tables["block_assignments"]
    for constraint in list(assignments.constraints):
        if isinstance(constraint, CheckConstraint):
            assignments.constraints.remove(constraint)
    assignments.append_constraint(CheckConstraint(
        "slot >= 0 AND slot < 3", name="ck_block_assignment_slot"
    ))
    if revision == "0003_game_day_task_blocks":
        people = metadata.tables["persons"]
        people._columns.remove(people.c.birth_date)
    label = "birth-date" if revision == "0003_game_day_task_blocks" else "cake-block"
    if (_unreviewed_metadata_drift(engine, metadata, inherited_defaults)
            or _check_constraint_drift(engine, metadata)):
        raise SchemaMigrationError(f"The {label} source schema does not match its reviewed revision")


def migrate_to_head(database_path: str | Path, engine) -> MigrationResult:
    """Adopt an exact baseline or upgrade a known revision after a safe snapshot."""
    path = Path(database_path).resolve()
    state = inspect_schema(path)
    heads = head_revisions()
    if len(heads) != 1:
        raise SchemaMigrationError(f"Expected one Alembic head, found {heads!r}")
    head = heads[0]
    if state.kind == "head":
        verify_head_schema(engine)
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

    inherited_defaults = _historical_defaults(engine)
    if state.revision in {"0003_game_day_task_blocks", "0004_person_birth_dates"}:
        _verify_pre_cake_schema(engine, state.revision, inherited_defaults)

    backup_path = _create_backup(path)
    try:
        retained_columns, retained_fingerprints = _retained_data(backup_path)
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
        final_state = inspect_schema(path)
        if final_state.kind != "head" or final_state.revision != head:
            raise SchemaMigrationError(
                f"Schema postflight failed with state {final_state.kind!r}"
            )
        _, final_fingerprints = _retained_data(path, retained_columns)
        changed = [table for table in retained_fingerprints
                   if final_fingerprints.get(table) != retained_fingerprints[table]]
        if changed:
            raise SchemaMigrationError(
                f"Retained-data postflight failed for tables: {', '.join(sorted(changed))}"
            )
        verify_head_schema(engine, inherited_defaults=inherited_defaults)
    except BaseException as exc:
        if isinstance(exc, SchemaMigrationError):
            message = str(exc)
        else:
            message = f"{type(exc).__name__}: {exc}"
        raise SchemaMigrationError(
            f"Schema upgrade failed. Retained backup: {backup_path}. {message}"
        ) from exc

    return MigrationResult(path, backup_path, state, head)


def metadata_drift(engine, metadata) -> list:
    """Return Alembic autogenerate differences for an initialized head schema."""
    with engine.connect() as connection:
        context = MigrationContext.configure(
            connection,
            opts={"compare_type": True, "compare_server_default": True},
        )
        return compare_metadata(context, metadata)
