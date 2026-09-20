"""One-time SQLite migration from source-key to canonical game-number identity."""

from __future__ import annotations

import sqlite3
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path

import scraper


SCHEMA_VERSION = 2


class MigrationError(RuntimeError):
    """Raised when the legacy database cannot be migrated without guessing."""


@dataclass(frozen=True)
class MigrationPlan:
    legacy_game_count: int
    resulting_game_count: int
    spielfest_groups: int
    redundant_spielfest_rows: int
    assignment_count: int
    audit_count: int


@dataclass(frozen=True)
class MigrationResult:
    database_path: Path
    backup_path: Path
    plan: MigrationPlan


def _connect(path: str | Path, *, readonly: bool = False) -> sqlite3.Connection:
    resolved = Path(path).resolve()
    target = f"file:{resolved}?mode=ro" if readonly else str(resolved)
    connection = sqlite3.connect(target, uri=readonly, timeout=0)
    connection.row_factory = sqlite3.Row
    return connection


def _columns(connection: sqlite3.Connection, table: str) -> set[str]:
    return {row["name"] for row in connection.execute(f"PRAGMA table_info({table})")}


def schema_state(path: str | Path) -> str:
    """Return ``missing``, ``legacy`` or ``current`` without changing the database."""
    resolved = Path(path).resolve()
    if not resolved.exists():
        return "missing"
    with _connect(resolved, readonly=True) as connection:
        columns = _columns(connection, "games")
        if not columns:
            return "missing"
        return "legacy" if "source_key" in columns else "current"


def _game_dict(row: sqlite3.Row) -> dict:
    return {
        "day": row["day"] or "",
        "date": row["date"] or "",
        "time": row["time"] or "",
        "hall": row["hall"],
        "game_nr": str(row["game_nr"]),
        "ak": row["ak"] or "",
        "home": row["home"] or "",
        "guest": row["guest"] or "",
        "score": row["score"] or "",
    }


def _build_plan(connection: sqlite3.Connection) -> tuple[MigrationPlan, list[dict]]:
    if "source_key" not in _columns(connection, "games"):
        raise MigrationError("Die Datenbank verwendet nicht das alte Spielschema.")

    rows = list(connection.execute("SELECT * FROM games ORDER BY id"))
    ordinary: list[sqlite3.Row] = []
    spf_groups: dict[tuple[int, str, str], list[sqlite3.Row]] = defaultdict(list)
    for row in rows:
        if scraper.is_spielfest_age_group(row["ak"]):
            key = (
                row["season_year"],
                row["date"] or "",
                scraper.normalize_age_group(row["ak"] or ""),
            )
            spf_groups[key].append(row)
        else:
            ordinary.append(row)

    duplicates: dict[tuple[int, str], list[int]] = defaultdict(list)
    for row in ordinary:
        duplicates[(row["season_year"], str(row["game_nr"]))].append(row["id"])
    ambiguous = {key: ids for key, ids in duplicates.items() if len(ids) > 1}
    if ambiguous:
        details = "; ".join(
            f"Saison {season}, Nr. {number}: IDs {ids}"
            for (season, number), ids in sorted(ambiguous.items())
        )
        raise MigrationError(
            "Mehrdeutige normale Spielnummern müssen manuell bereinigt werden: "
            + details
        )

    migrated: list[dict] = []
    for row in ordinary:
        item = dict(row)
        item["game_nr"] = str(row["game_nr"])
        migrated.append(item)

    redundant_count = 0
    for (_season, _date, _normalized_ak), group in sorted(spf_groups.items()):
        aggregate = scraper.collapse_spielfeste([_game_dict(row) for row in group])
        if len(aggregate) != 1:
            raise MigrationError(f"Spielfest-Gruppe konnte nicht verdichtet werden: {group!r}")

        ids = [row["id"] for row in group]
        assignments = list(connection.execute(
            f"SELECT * FROM assignments WHERE game_id IN ({','.join('?' for _ in ids)})",
            ids,
        ))
        slot_keys = [(row["role"], row["slot"]) for row in assignments]
        person_ids = [row["person_id"] for row in assignments]
        if len(slot_keys) != len(set(slot_keys)) or len(person_ids) != len(set(person_ids)):
            raise MigrationError(
                f"Widersprüchliche Spielfest-Zuordnungen bei Spiel-IDs {ids}."
            )
        team_ids = {row["team_id"] for row in group if row["team_id"] is not None}
        jteams = {row["jteam"] for row in group if row["jteam"]}
        if len(team_ids) > 1 or len(jteams) > 1:
            raise MigrationError(
                f"Widersprüchliche Spielfest-Verantwortung bei Spiel-IDs {ids}."
            )

        state_ids = {
            row["game_id"] for row in assignments
        } | {
            row["id"] for row in group
            if row["team_id"] is not None or row["jteam"]
        }
        survivor_id = min(state_ids) if state_ids else min(ids)
        survivor = next(row for row in group if row["id"] == survivor_id)
        event = aggregate[0]
        item = dict(survivor)
        item.update(event)
        item["id"] = survivor_id
        if team_ids:
            item["team_id"] = next(iter(team_ids))
        if jteams:
            item["jteam"] = next(iter(jteams))
        item["_merged_ids"] = ids
        migrated.append(item)
        redundant_count += len(group) - 1

    assignment_count = connection.execute(
        "SELECT COUNT(*) FROM assignments"
    ).fetchone()[0]
    audit_count = connection.execute(
        "SELECT COUNT(*) FROM assignment_audit"
    ).fetchone()[0]
    return MigrationPlan(
        legacy_game_count=len(rows),
        resulting_game_count=len(migrated),
        spielfest_groups=len(spf_groups),
        redundant_spielfest_rows=redundant_count,
        assignment_count=assignment_count,
        audit_count=audit_count,
    ), migrated


def preflight(path: str | Path) -> MigrationPlan:
    """Validate a legacy database and return a non-mutating migration summary."""
    with _connect(path, readonly=True) as connection:
        plan, _rows = _build_plan(connection)
        return plan


def _backup_path(path: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%dT%H%M%S%f")
    return path.with_name(f"{path.name}.pre-game-identity-{stamp}.db")


def _create_backup(source_path: Path) -> Path:
    destination_path = _backup_path(source_path)
    with _connect(source_path, readonly=True) as source:
        with sqlite3.connect(destination_path) as destination:
            source.backup(destination)
            result = destination.execute("PRAGMA integrity_check").fetchone()[0]
            if result != "ok":
                raise MigrationError(f"Backup-Integritätsprüfung fehlgeschlagen: {result}")
    return destination_path


def migrate(path: str | Path) -> MigrationResult:
    """Back up and transactionally migrate one stopped legacy SQLite database."""
    resolved = Path(path).resolve()
    plan = preflight(resolved)
    backup_path = _create_backup(resolved)

    connection = _connect(resolved)
    try:
        connection.execute("PRAGMA foreign_keys=OFF")
        connection.execute("BEGIN EXCLUSIVE")
        current_plan, rows = _build_plan(connection)
        if current_plan != plan:
            raise MigrationError("Die Datenbank wurde nach dem Preflight verändert.")

        for row in rows:
            merged_ids = row.get("_merged_ids", [row["id"]])
            if len(merged_ids) == 1:
                continue
            survivor_id = row["id"]
            placeholders = ",".join("?" for _ in merged_ids)
            connection.execute(
                f"UPDATE assignments SET game_id=? WHERE game_id IN ({placeholders})",
                [survivor_id, *merged_ids],
            )
            connection.execute(
                f"UPDATE assignment_audit SET game_id=? WHERE game_id IN ({placeholders})",
                [survivor_id, *merged_ids],
            )

        connection.execute("""
            CREATE TABLE games_new (
                id INTEGER NOT NULL PRIMARY KEY,
                season_year INTEGER NOT NULL,
                game_nr VARCHAR(100) NOT NULL,
                day VARCHAR(20), date VARCHAR(20), time VARCHAR(30),
                hall INTEGER, ak VARCHAR(10), home VARCHAR(120), guest VARCHAR(120),
                score VARCHAR(120), jteam VARCHAR(120), team_id INTEGER,
                CONSTRAINT uq_season_game_nr UNIQUE (season_year, game_nr),
                FOREIGN KEY(team_id) REFERENCES teams (id)
            )
        """)
        insert_sql = """
            INSERT INTO games_new
                (id, season_year, game_nr, day, date, time, hall, ak, home, guest,
                 score, jteam, team_id)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """
        for row in rows:
            connection.execute(insert_sql, (
                row["id"], row["season_year"], str(row["game_nr"]), row["day"],
                row["date"], row["time"], row["hall"], row["ak"], row["home"],
                row["guest"], row["score"], row["jteam"], row["team_id"],
            ))
        connection.execute("DROP TABLE games")
        connection.execute("ALTER TABLE games_new RENAME TO games")
        connection.execute(f"PRAGMA user_version={SCHEMA_VERSION}")

        if connection.execute("SELECT COUNT(*) FROM games").fetchone()[0] != plan.resulting_game_count:
            raise MigrationError("Die Anzahl migrierter Spiele stimmt nicht.")
        if connection.execute("SELECT COUNT(*) FROM assignments").fetchone()[0] != plan.assignment_count:
            raise MigrationError("Zuordnungen gingen während der Migration verloren.")
        if connection.execute("SELECT COUNT(*) FROM assignment_audit").fetchone()[0] != plan.audit_count:
            raise MigrationError("Audit-Einträge gingen während der Migration verloren.")
        violations = connection.execute("PRAGMA foreign_key_check").fetchall()
        if violations:
            raise MigrationError(f"Fremdschlüsselprüfung fehlgeschlagen: {violations!r}")
        connection.commit()
    except sqlite3.OperationalError as exc:
        connection.rollback()
        raise MigrationError(
            "SQLite konnte nicht exklusiv migriert werden. Prüfe, dass alle "
            "Web- und Tagesjob-Schreiber gestoppt sind."
        ) from exc
    except Exception:
        connection.rollback()
        raise
    finally:
        connection.close()

    return MigrationResult(resolved, backup_path, plan)
