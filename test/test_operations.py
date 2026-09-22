"""Offline deployment validation, generic health and operation-state checks."""
import copy
from datetime import datetime, timedelta, timezone
import json
import os
from pathlib import Path
import tempfile
from unittest.mock import patch

import helpers as h
import db
import operations
import production as p
import webapp


def settings():
    return p.read_json(Path(h.PROJECT_DIR) / 'deploy/operations.json')


def legal():
    return dict(schema_version=2, version='test-v1', pages={
        name: dict(title=title, blocks=[dict(type='paragraph', text='<script>alert(1)</script>'),
              dict(type='list', items=['Synthetic list']),
              dict(type='link', text='Synthetic link', url='https://club.test/info')])
        for name, title in [('impressum', 'Impressum'), ('datenschutz', 'Datenschutzerklärung')]})


def test_operations_settings_reject_missing_placeholders_and_invalid_thresholds():
    good = settings()
    assert p.validate_operations(good) == good
    for key in good:
        bad = copy.deepcopy(good)
        del bad[key]
        try:
            p.validate_operations(bad)
        except p.ConfigurationError:
            pass
        else:
            raise AssertionError('missing setting accepted: ' + key)
    for key, value in [('hostname','example.invalid'), ('disk_critical', 25),
                       ('alert_to', 'not an email'), ('state_dir','relative'),
                       ('stale_seconds', float('nan'))]:
        bad = {**good, key:value}
        try:
            p.validate_operations(bad)
        except p.ConfigurationError as error:
            assert str(value) not in str(error)
        else:
            raise AssertionError('invalid setting accepted: ' + key)


def test_success_markers_are_atomic_restricted_and_reject_future_or_missing():
    now = datetime.now(timezone.utc)
    with tempfile.TemporaryDirectory() as directory:
        p.write_success('backup', directory, now - timedelta(days=1))
        path = Path(directory) / 'backup.success.json'
        assert path.stat().st_mode & 0o777 == 0o600
        assert p.success_age('backup', directory, now) == 86400
        before = path.read_bytes()
        with patch('production.os.replace', side_effect=OSError('synthetic')):
            try:
                p.write_success('backup', directory, now)
            except OSError:
                pass
        assert path.read_bytes() == before and len(list(Path(directory).iterdir())) == 1
        try:
            p.success_age('backup', directory, now - timedelta(days=2))
        except p.ConfigurationError:
            pass
        else:
            raise AssertionError('future marker accepted')


def test_health_is_fixed_even_with_authenticated_session_and_missing_database():
    with tempfile.TemporaryDirectory() as directory:
        path = str(Path(directory) / 'db.sqlite')
        db.initialize_db(db.make_engine(path))
        with patch.dict(os.environ, {'NULIGAHELPER_DB':path}):
            app = webapp.create_app()
        client = app.test_client()
        with client.session_transaction() as session:
            session['person_id'] = 999
        with patch.object(webapp.notifier, 'Notifier', side_effect=AssertionError('no dispatch')):
            response = client.get('/healthz')
            assert response.status_code == 200 and response.data == b'ok\n'
            assert 'Set-Cookie' not in response.headers
            os.rename(path, path + '.old')
            response = client.get('/healthz')
            assert response.status_code == 503 and response.data == b'unavailable\n'
            assert not Path(path).exists(), 'health must not recreate the database'


def test_legal_rendering_escapes_content_and_is_public_for_all_account_states():
    data = legal()
    with tempfile.TemporaryDirectory() as directory:
        file = Path(directory) / 'legal.json'
        file.write_text(json.dumps(data))
        with patch.dict(os.environ, {'NULIGAHELPER_LEGAL':str(file)}):
            app = webapp.create_app()
            for identity in (None, 999):
                client = app.test_client()
                if identity:
                    h.sign_in(client, identity)
                for url in ('/impressum','/datenschutz'):
                    response = client.get(url)
                    assert response.status_code == 200
                    html = response.get_data(as_text=True)
                    assert '&lt;script&gt;' in html and '<script>alert' not in html
                    assert 'href="/impressum"' in html and 'href="/datenschutz"' in html
            for url in ('/', '/login', '/registrieren'):
                html = app.test_client().get(url).get_data(as_text=True)
                assert 'href="/impressum"' in html and 'href="/datenschutz"' in html
            file.unlink()
            assert app.test_client().get('/impressum').status_code == 503


def test_legal_rejects_unsafe_links_and_placeholders():
    good = legal()
    assert p.validate_legal(good)
    for modify in ('url', 'placeholder', 'missing_page'):
        bad = copy.deepcopy(good)
        if modify == 'url':
            bad['pages']['impressum']['blocks'][-1]['url'] = 'javascript:alert(1)'
        if modify == 'placeholder':
            bad['pages']['impressum']['title'] = 'OFFEN'
        if modify == 'missing_page': bad['pages'].pop('impressum')
        try: p.validate_legal(bad)
        except p.ConfigurationError: pass
        else: raise AssertionError('invalid legal content accepted: ' + modify)


def test_monitor_detects_staleness_missing_timers_clock_disk_and_tls():
    cfg = settings()
    def healthy_command(*args):
        if args[0] == 'timedatectl': return 'yes'
        return 'LoadState=loaded\nActiveState=active\nResult=success\nUnitFileState=enabled'
    with tempfile.TemporaryDirectory() as directory:
        cfg['state_dir'] = directory
        for phase in ('application','backup','cleanup'): p.write_success(phase,directory)
        with patch.object(db, 'database_ready', return_value=True):
            result = operations.monitor(cfg,runner=healthy_command,cert=lambda host:30,
                                       statvfs=lambda path:os.statvfs(directory))
            assert not any(result.values()), result
            result = operations.monitor(cfg,runner=lambda *args:'no',cert=lambda host:1,
                                       statvfs=lambda path:os.statvfs(directory))
            assert result['clock'] == result['certificate'] == result['nuligahelper-daily.timer'] == 2
            p.write_success('backup',directory,datetime.now(timezone.utc)-timedelta(days=6))
            result = operations.monitor(cfg,runner=healthy_command,cert=lambda host:10,
                                       statvfs=lambda path:os.statvfs(directory))
            assert result['backup'] == 2 and result['certificate'] == 1
    assert operations.threshold(9,20,10)==2
    assert operations.threshold(15,20,10)==1
    assert operations.threshold(21,20,10)==0


if __name__ == '__main__':
    h.run_all(globals())
