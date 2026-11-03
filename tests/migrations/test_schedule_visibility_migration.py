"""Guarded visibility migration preserves every pre-existing domain value."""

from datetime import date
from pathlib import Path
import shutil
import sqlite3
import tempfile
from unittest.mock import patch

from alembic import command
from alembic.script import ScriptDirectory

from tests import helpers as h
from nuligahelper import db
from nuligahelper.db import schema_upgrade as sm
from tests.schema_fixtures import create_versioned_database

PREVIOUS = sm.WISCHER_REVISION


def _rows(path):
    with sqlite3.connect(path) as connection:
        tables = [row[0] for row in connection.execute(
            "SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%'"
        ) if row[0] not in ('alembic_version', 'schedule_visibility_setting')]
        return {table: sorted(connection.execute(f'SELECT * FROM "{table}"').fetchall(), key=repr)
                for table in tables}


def _source(path):
    engine = create_versioned_database(path, PREVIOUS)
    with h.Session(engine) as session:
        first = db.Person(name='Fixture admin', birth_date=h.ADULT_BIRTH_DATE, is_admin=True)
        second = db.Person(name='Fixture helper', birth_date=None)
        team = db.get_support_team(session)
        first.teams = [team]
        second.teams = [team]
        session.add_all([first, second])
        session.flush()
        team.mv_person = first
        db.sync_games(session, h.sample_games(), h.SEASON)
        game = session.query(db.Game).filter_by(game_nr='1005').one()
        block = db.get_day_blocks(session, h.SEASON, game.date)[0]
        session.add_all([
            db.Assignment(game=game, person=first, role=db.ROLE_SALE, slot=0,
                          replacement_token='preserved-game-generation'),
            db.Assignment(game=game, person=second, role=db.ROLE_SECRETARY, slot=0,
                          locked_at=db.datetime(2026, 10, 31, 9)),
            db.BlockAssignment(block=block, person=first, slot=2,
                               replacement_token='preserved-block-generation'),
            db.BlockAssignment(block=block, person=second, slot=0,
                               locked_at=db.datetime(2026, 10, 31, 9)),
            db.AssignmentAudit(actor_tier='system', action='claim', role='Reinigung', slot=1,
                               actor_name='System', affected_person_name='Historical name',
                               game_snapshot='literal historical snapshot', game_id=game.id),
        ])
        session.commit()
    return engine


def test_successor_and_fresh_initialization_match_and_retain_every_domain_row():
    scripts = ScriptDirectory.from_config(sm.alembic_config())
    assert scripts.get_heads() == [sm.HEAD_REVISION]
    assert scripts.get_revision(sm.HEAD_REVISION).down_revision == PREVIOUS
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / 'prior.db'
        engine = _source(path)
        before = _rows(path)
        result = sm.migrate_to_head(path, engine)
        assert result.backup_path and _rows(result.backup_path) == before
        assert _rows(path) == before, 'identities, locks, tokens, memberships and all old audits must survive'
        with h.Session(engine) as session:
            assert session.query(db.ScheduleVisibilitySetting).count() == 1
            assert db.schedule_visibility_boundary(session) is None
            session.get(db.ScheduleVisibilitySetting, 1).hide_from = date(2026, 12, 1)
            session.commit()
        assert sm.migrate_to_head(path, engine).backup_path is None
        with h.Session(engine) as session:
            assert db.schedule_visibility_boundary(session) == date(2026, 12, 1)
        assert _rows(path) == before
        fresh = h.make_engine()
        assert sm.metadata_drift(engine, db.Base.metadata) == sm.metadata_drift(fresh, db.Base.metadata) == []
        db.verify_db(engine)
        fresh.dispose()
        engine.dispose()
        shutil.copyfile(result.backup_path, path)
        assert sm.inspect_schema(path).revision == PREVIOUS and _rows(path) == before
        engine = db.make_engine(str(path))
        sm.migrate_to_head(path, engine)
        assert _rows(path) == before
        engine.dispose()


def test_old_startup_and_near_miss_schemas_fail_without_implicit_seed_or_upgrade():
    for mutation in (None, 'ALTER TABLE persons ADD COLUMN surprise TEXT',
                     'CREATE INDEX unreviewed_index ON games (date)'):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'old.db'
            engine = create_versioned_database(path, PREVIOUS)
            if mutation:
                with sqlite3.connect(path) as connection:
                    connection.execute(mutation)
            before = path.read_bytes()
            try:
                db.verify_db(engine)
            except db.SQLiteInitializationError as exc:
                assert 'migrate-schema' in str(exc)
            else:
                raise AssertionError('old runtime schema must require explicit migration')
            assert path.read_bytes() == before
            if mutation:
                try:
                    sm.migrate_to_head(path, engine)
                except sm.SchemaMigrationError as exc:
                    assert 'drift' in str(exc)
                else:
                    raise AssertionError('the previous revision label must not authorize drift')
                assert not list(Path(directory).glob('*.pre-schema-*.db'))
            engine.dispose()


def test_singleton_integrity_and_initial_unrestricted_postflight_fail_closed():
    for mutation in ('DELETE FROM schedule_visibility_setting',
                     "UPDATE schedule_visibility_setting SET hide_from='invalid'",
                     "UPDATE schedule_visibility_setting SET hide_from='2026-12-01'"):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'failure.db'
            engine = _source(path)
            before = _rows(path)
            original = sm._run_alembic

            def damage(connection, operation, revision):
                original(connection, operation, revision)
                if operation is command.upgrade:
                    connection.exec_driver_sql(mutation)

            with patch.object(sm, '_run_alembic', side_effect=damage):
                try:
                    sm.migrate_to_head(path, engine)
                except sm.SchemaMigrationError as exc:
                    assert 'Retained backup:' in str(exc)
                else:
                    raise AssertionError('postflight must verify the singleton and unrestricted initial value')
            assert _rows(path) == before
            if '2026-12-01' not in mutation:
                try:
                    db.verify_db(engine)
                except db.SQLiteInitializationError:
                    pass
                else:
                    raise AssertionError('invalid head state must refuse startup, never repair itself')
            engine.dispose()


def test_singleton_key_and_nullable_date_are_constrained():
    engine = h.make_engine()
    with engine.begin() as connection:
        try:
            connection.exec_driver_sql('INSERT INTO schedule_visibility_setting (id) VALUES (2)')
        except db.IntegrityError:
            pass
        else:
            raise AssertionError('the setting must have one fixed identity')
    with h.Session(engine) as session:
        session.delete(session.get(db.ScheduleVisibilitySetting, 1))
        session.commit()
        try:
            db.schedule_visibility_boundary(session)
        except db.SQLiteInitializationError:
            pass
        else:
            raise AssertionError('a GET lookup must never seed missing state')
        assert session.query(db.ScheduleVisibilitySetting).count() == 0
    engine.dispose()


if __name__ == '__main__':
    h.run_all(dict(globals()))
