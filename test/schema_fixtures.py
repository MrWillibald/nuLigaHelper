from __future__ import annotations

import sqlite3
from pathlib import Path


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
