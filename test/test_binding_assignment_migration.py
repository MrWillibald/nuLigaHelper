"""Explicit single-head upgrade retains every old value without inventing consent."""

from pathlib import Path
import sqlite3
import tempfile

import helpers as h
from nuligahelper import db
from nuligahelper.db import schema_upgrade
from alembic.script import ScriptDirectory
from schema_fixtures import create_versioned_database

PREVIOUS = "0006_game_duty_staffing"


def test_binding_revision_is_single_head_and_fresh_schema_has_no_model_drift():
    assert schema_upgrade.head_revisions() == (schema_upgrade.HEAD_REVISION,)
    scripts = ScriptDirectory.from_config(schema_upgrade.alembic_config())
    assert scripts.get_revision("0007_binding_duty_assignments").down_revision == PREVIOUS
    engine = h.make_engine()
    schema_upgrade.verify_head_schema(engine)
    assert schema_upgrade.metadata_drift(engine, db.Base.metadata) == []
    engine.dispose()


def test_populated_upgrade_preserves_assignments_teams_and_audits_without_fake_lifecycle():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "previous.db"
        engine = create_versioned_database(path, PREVIOUS)
        with sqlite3.connect(path) as con:
            con.execute("INSERT INTO persons (id,name,is_admin,account_status,birth_date) VALUES (11,'Synthetic',0,'active','1990-01-01')")
            con.execute("INSERT INTO teams (id,name,is_support,mv_person_id) VALUES (12,'Synthetic team',0,11)")
            con.execute("INSERT INTO person_teams VALUES (11,12)")
            con.execute("INSERT INTO games (id,season_year,game_nr,date,ak) VALUES (100,2026,'9001','01.11.2026','M')")
            con.execute("INSERT INTO assignments VALUES (200,100,11,'Verkauf',1)")
            con.execute("INSERT INTO day_blocks (id,season_year,date,phase,cake_quantity,delivery_time) VALUES (300,2026,'01.11.2026','cake_delivery',4,'10:30')")
            con.execute("INSERT INTO block_assignments VALUES (400,300,11,3)")
            con.execute("INSERT INTO assignment_audit (id,changed_at,actor_tier,action,affected_person_id,game_id,role,slot,actor_name,affected_person_name,game_snapshot) VALUES (500,'2026-01-01','system','claim',11,100,'Verkauf',1,'System','Synthetic','untouched snapshot')")
        columns, before = schema_upgrade._retained_data(path)
        result = schema_upgrade.migrate_to_head(path, engine)
        assert result.previous_state.revision == PREVIOUS and result.backup_path
        assert schema_upgrade._retained_data(path, columns)[1] == before
        with sqlite3.connect(path) as con:
            assert con.execute("SELECT id,slot,replacement_token,locked_at FROM assignments").fetchall() == [(200,1,None,None)]
            assert con.execute("SELECT id,slot,replacement_token,locked_at FROM block_assignments").fetchall() == [(400,3,None,None)]
            assert con.execute("SELECT * FROM duty_cutoffs").fetchall() == []
            assert con.execute("SELECT game_snapshot FROM assignment_audit").fetchall() == [('untouched snapshot',)]
            assert con.execute("PRAGMA foreign_key_check").fetchall() == []
        db.verify_db(engine)
        engine.dispose()


def test_predecessor_drift_is_refused_before_snapshot_or_schema_write():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "drift.db"
        engine = create_versioned_database(path, PREVIOUS)
        with sqlite3.connect(path) as con:
            con.execute("ALTER TABLE assignments ADD COLUMN replacement_token TEXT")
        before = path.read_bytes()
        try:
            schema_upgrade.migrate_to_head(path, engine)
        except schema_upgrade.SchemaMigrationError as exc:
            assert "source schema does not match" in str(exc)
        else:
            raise AssertionError("a recognized revision cannot authorize a partially added lifecycle schema")
        assert path.read_bytes() == before
        assert not list(path.parent.glob('*.pre-schema-*.db'))
        engine.dispose()


if __name__ == '__main__':
    h.run_all(dict(globals()))
