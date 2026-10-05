from __future__ import annotations

import sqlite3
from pathlib import Path


PRE_GAME_DUTY_REVISION = "0005_cake_delivery_blocks"


LEGACY_SCHEMA = """
CREATE TABLE persons (
    id INTEGER NOT NULL PRIMARY KEY,
    name VARCHAR(120) NOT NULL,
    email VARCHAR(200),
    phone VARCHAR(60),
    team_id INTEGER REFERENCES teams(id),
    desired_team_id INTEGER REFERENCES teams(id),
    is_admin BOOLEAN NOT NULL DEFAULT 0,
    account_status VARCHAR(20) NOT NULL DEFAULT 'active',
    registered_at DATETIME,
    rejected_at DATETIME,
    verified_at DATETIME,
    approved_at DATETIME,
    UNIQUE (email),
    UNIQUE (phone)
);
CREATE TABLE teams (
    id INTEGER NOT NULL PRIMARY KEY,
    name VARCHAR(120) NOT NULL,
    is_support BOOLEAN NOT NULL DEFAULT 0,
    mv_person_id INTEGER REFERENCES persons(id),
    UNIQUE (name)
);
CREATE TABLE games (
    id INTEGER NOT NULL PRIMARY KEY,
    season_year INTEGER NOT NULL,
    game_nr VARCHAR(100) NOT NULL,
    day VARCHAR(20), date VARCHAR(20), time VARCHAR(30), hall INTEGER,
    ak VARCHAR(10), home VARCHAR(120), guest VARCHAR(120), score VARCHAR(120),
    jteam VARCHAR(120), team_id INTEGER REFERENCES teams(id),
    CONSTRAINT uq_season_game_nr UNIQUE (season_year, game_nr)
);
CREATE TABLE assignments (
    id INTEGER NOT NULL PRIMARY KEY,
    game_id INTEGER NOT NULL REFERENCES games(id),
    person_id INTEGER NOT NULL REFERENCES persons(id),
    role VARCHAR(40) NOT NULL,
    slot INTEGER NOT NULL,
    CONSTRAINT uq_game_person UNIQUE (game_id, person_id),
    CONSTRAINT uq_game_role_slot UNIQUE (game_id, role, slot)
);
CREATE TABLE auth_tokens (
    id INTEGER NOT NULL PRIMARY KEY,
    nonce VARCHAR(120) NOT NULL UNIQUE,
    code VARCHAR(12), purpose VARCHAR(30) NOT NULL,
    person_id INTEGER NOT NULL REFERENCES persons(id),
    issued_at DATETIME NOT NULL, expires_at DATETIME NOT NULL, used_at DATETIME
);
CREATE TABLE assignment_audit (
    id INTEGER NOT NULL PRIMARY KEY,
    changed_at DATETIME NOT NULL,
    actor_person_id INTEGER REFERENCES persons(id), actor_tier VARCHAR(20) NOT NULL,
    action VARCHAR(20) NOT NULL,
    affected_person_id INTEGER REFERENCES persons(id),
    game_id INTEGER REFERENCES games(id), role VARCHAR(40) NOT NULL,
    slot INTEGER NOT NULL, actor_name VARCHAR(120) NOT NULL,
    affected_person_name VARCHAR(120) NOT NULL,
    game_snapshot VARCHAR(300) NOT NULL
);
CREATE TABLE auth_abuse_counters (
    id INTEGER NOT NULL PRIMARY KEY,
    action VARCHAR(40) NOT NULL, dimension VARCHAR(20) NOT NULL,
    subject_digest VARCHAR(64) NOT NULL, channel VARCHAR(10) NOT NULL,
    window_started_at DATETIME NOT NULL, count INTEGER NOT NULL,
    expires_at DATETIME NOT NULL,
    CONSTRAINT uq_auth_abuse_bucket UNIQUE (
        action, dimension, subject_digest, channel, window_started_at
    )
);
CREATE INDEX ix_auth_abuse_expires_at ON auth_abuse_counters (expires_at);
"""


def create_legacy_database(path: str | Path) -> None:
    with sqlite3.connect(path) as connection:
        connection.executescript(LEGACY_SCHEMA)
        connection.execute(
            "INSERT INTO teams (id, name, is_support) VALUES (1, 'Supporter', 1)"
        )
        connection.execute("PRAGMA user_version=2")
        connection.commit()


def create_versioned_database(path: str | Path, revision: str):
    """Build an actual historical revision, independent of the current ORM."""
    from alembic import command
    import db
    import schema_migrations

    # Match the historical ORM rather than the separately accepted manual
    # baseline (whose server defaults are intentionally retained).
    schema = LEGACY_SCHEMA.replace(" DEFAULT 0", "").replace(" DEFAULT 'active'", "")
    schema = schema.replace("nonce VARCHAR(120) NOT NULL UNIQUE", "nonce VARCHAR(120) NOT NULL")
    schema = schema.replace("expires_at DATETIME NOT NULL, used_at DATETIME\n);",
                            "expires_at DATETIME NOT NULL, used_at DATETIME,\n    UNIQUE (nonce)\n);")
    schema = schema.replace("UNIQUE (\n        action, dimension, subject_digest, channel, window_started_at\n    )",
                            "UNIQUE (action, dimension, subject_digest, channel, window_started_at)")
    with sqlite3.connect(path) as connection:
        connection.executescript(schema)
        connection.execute("INSERT INTO teams (id, name, is_support) VALUES (1, 'Supporter', 1)")
    engine = db.make_engine(str(path))
    with engine.connect() as connection:
        connection.commit()
        connection.exec_driver_sql("PRAGMA foreign_keys=OFF")
        schema_migrations._run_alembic(connection, command.stamp, schema_migrations.BASELINE_REVISION)
        connection.commit()
        schema_migrations._run_alembic(connection, command.upgrade, revision)
        connection.commit()
        connection.exec_driver_sql("PRAGMA foreign_keys=ON")
    return engine


def create_game_duty_source(path: str | Path):
    """Populated reviewed predecessor with literal historical role names."""
    engine = create_versioned_database(path, PRE_GAME_DUTY_REVISION)
    with sqlite3.connect(path) as connection:
        connection.executemany(
            "INSERT INTO persons (id, name, is_admin, account_status, birth_date, email) "
            "VALUES (?, ?, 0, 'active', ?, ?)",
            [(11, "Legacy cashier", "1980-02-29", "cashier@example.test"),
             (12, "Legacy security", None, None),
             (13, "Legacy seller", "1990-01-01", None),
             (14, "Legacy cake helper", None, None),
             (15, "Unassigned helper", None, None)],
        )
        connection.executemany("INSERT INTO person_teams VALUES (?, 1)",
                               [(11,), (12,), (13,), (14,), (15,)])
        connection.execute("UPDATE teams SET mv_person_id=11 WHERE id=1")
        connection.executemany(
            "INSERT INTO games (id, season_year, game_nr, date, ak) VALUES (?, 2026, ?, ?, ?)",
            [(100, "9001", "01.11.2026", "BL M"),
             (101, "9002", "08.11.2026", "BL mD")],
        )
        connection.executemany("INSERT INTO assignments VALUES (?, 100, ?, ?, ?)",
                               [(401, 11, "Unterstützung", 0),
                                (402, 12, "Ordnungsdienst", 0),
                                (403, 13, "Verkauf", 1)])
        connection.execute(
            "INSERT INTO day_blocks (id, season_year, date, phase, cake_quantity, delivery_time) "
            "VALUES (201, 2026, '01.11.2026', 'cake_delivery', 4, '09:30')"
        )
        connection.execute("INSERT INTO block_assignments VALUES (301, 201, 14, 3)")
        connection.executemany(
            "INSERT INTO assignment_audit "
            "(id, changed_at, actor_tier, action, affected_person_id, game_id, role, slot, "
            "actor_name, affected_person_name, game_snapshot) "
            "VALUES (?, '2026-01-01', 'system', 'claim', 11, 100, ?, 0, "
            "'System', 'Legacy cashier', ?)",
            [(501, "Unterstützung", "recorded support snapshot"),
             (502, "Reinigung", "recorded original cleaning snapshot")],
        )
        connection.execute(
            "INSERT INTO assignment_audit "
            "(id, changed_at, actor_tier, action, affected_person_id, role, slot, "
            "actor_name, affected_person_name, block_id, block_snapshot) "
            "VALUES (503, '2026-01-01', 'system', 'claim', 14, 'Kuchenlieferung', 3, "
            "'System', 'Legacy cake helper', 201, 'recorded cake snapshot')"
        )
        connection.execute(
            "INSERT INTO auth_tokens (id, nonce, code, purpose, person_id, issued_at, expires_at) "
            "VALUES (601, 'synthetic-token', '123456', 'login', 11, '2026-01-01', '2026-01-01 00:15')"
        )
    return engine
