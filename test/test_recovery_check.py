"""Offline recovery-point and read-only schema gate checks."""

import importlib.util
from pathlib import Path
import sqlite3
import tempfile
from unittest.mock import patch

import helpers as h
import backup
import db
import schema_migrations


SOURCE = Path(h.PROJECT_DIR) / 'release-assets/recovery_check.py'
SPEC = importlib.util.spec_from_file_location('nuligahelper_recovery_check', SOURCE)
recovery = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(recovery)


def test_snapshot_captures_committed_wal_state_and_is_self_contained():
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        database = root / 'live.db'
        recovery_dir = root / 'recovery'
        recovery_dir.mkdir()
        destination = recovery_dir / 'before.db'
        with sqlite3.connect(database) as connection:
            assert connection.execute('PRAGMA journal_mode=WAL').fetchone()[0] == 'wal'
            connection.execute('CREATE TABLE example (id INTEGER PRIMARY KEY, value TEXT)')
            connection.execute("INSERT INTO example(value) VALUES ('before')")
            connection.commit()
            connection.execute("INSERT INTO example(value) VALUES ('committed in WAL')")
            connection.commit()
            with patch.object(recovery, '_private_directory'):
                assert recovery.durable_snapshot(database, destination) == destination
            assert destination.stat().st_mode & 0o777 == 0o600
            assert not Path(str(destination) + '-wal').exists()
            assert not Path(str(destination) + '-shm').exists()
            backup.validate_snapshot(destination)
            with sqlite3.connect(f'{destination.as_uri()}?mode=ro&immutable=1', uri=True) as snapshot:
                values = [row[0] for row in snapshot.execute('SELECT value FROM example ORDER BY id')]
            assert values == ['before', 'committed in WAL']


def test_failed_snapshot_publication_removes_partial_destination():
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        database = root / 'live.db'
        with sqlite3.connect(database) as connection:
            connection.execute('CREATE TABLE example (id INTEGER PRIMARY KEY)')
        destination = root / 'before.db'
        with patch.object(recovery, '_private_directory'), \
                patch.object(recovery, '_fsync_directory', side_effect=OSError('synthetic')):
            try:
                recovery.durable_snapshot(database, destination)
            except OSError:
                pass
            else:
                raise AssertionError('failed durability check accepted snapshot')
        assert not destination.exists()


def test_schema_gate_accepts_head_pauses_old_and_refuses_unknown_or_missing():
    with tempfile.TemporaryDirectory() as directory:
        database = Path(directory) / 'live.db'
        assert recovery.schema_gate(database)['gate'] == 'refused'
        engine = db.make_engine(str(database))
        db.initialize_db(engine)
        engine.dispose()
        ready = recovery.schema_gate(database)
        assert ready['gate'] == 'ready'
        assert ready['revision'] == '0003_game_day_task_blocks'
        with sqlite3.connect(database) as connection:
            connection.execute("UPDATE alembic_version SET version_num='0002_multi_team_membership'")
        assert recovery.schema_gate(database)['gate'] == 'migration_required'
        with sqlite3.connect(database) as connection:
            connection.execute("UPDATE alembic_version SET version_num='unrecognized_revision'")
        assert recovery.schema_gate(database)['gate'] == 'refused'


def test_schema_gate_refuses_newer_divergent_and_corrupt_databases():
    with tempfile.TemporaryDirectory() as directory:
        database = Path(directory) / 'live.db'
        engine = db.make_engine(str(database))
        db.initialize_db(engine)
        engine.dispose()
        with patch.object(schema_migrations, 'head_revisions',
                          return_value=('0002_multi_team_membership',)):
            assert recovery.schema_gate(database)['gate'] == 'refused', \
                'a database ahead of the candidate must never be migrated backwards'
        with patch.object(schema_migrations, 'inspect_schema', return_value=
                          schema_migrations.SchemaState('divergent')):
            assert recovery.schema_gate(database)['gate'] == 'refused'
        corrupt = Path(directory) / 'corrupt.db'
        corrupt.write_bytes(b'not sqlite')
        assert recovery.schema_gate(corrupt)['gate'] == 'refused'


if __name__ == '__main__':
    h.run_all(globals())
