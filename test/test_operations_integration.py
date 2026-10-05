"""Operation outcomes, redaction, launch binding and operator-only alert tests."""
import copy
from dataclasses import replace
from datetime import datetime, timezone
import io
import json
import logging
import os
from pathlib import Path
import tempfile
from unittest.mock import Mock, patch

import helpers as h
import backup
import main
import messages
import notifier
import operations
import production as p
from production_logging import JournalFormatter
from test_main import _fake_job
from test_operations import settings, legal
from test_privacy import policy


def test_journal_never_formats_raw_payloads_or_tracebacks():
    formatter = JournalFormatter('%(levelname)s %(message)s')
    for name in ('webapp','gunicorn.error','gunicorn.access','nuligahelper.security','root'):
        record = logging.LogRecord(name,logging.ERROR,'',0,
            'canary@example.test +49123456789 secret-canary 654321 /login/token/token-canary',(),None)
        try: raise ValueError('secret-canary')
        except ValueError:
            import sys
            record.exc_info = sys.exc_info()
        result = formatter.format(record)
        for canary in ('canary','654321','49123456789','/login/token'):
            assert canary not in result
    record = logging.LogRecord('nuligahelper.operations',logging.ERROR,'',0,
        'operation=backup outcome=failure reason=retention_delete',(),None)
    assert 'retention_delete' in formatter.format(record)


def test_daily_markers_separate_backup_and_application_failures():
    for failed in ('none','backup','application'):
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory,'test.db')
            kwargs = {}
            if failed == 'backup': kwargs['backup_error'] = backup.BackupError(backup.FailureStage.LATEST_UPLOAD,'synthetic')
            if failed == 'application': kwargs['notification_error'] = RuntimeError('synthetic')
            with patch.dict(os.environ,{'NULIGAHELPER_STATE_DIR':directory}), _fake_job(path,**kwargs):
                try: main.main()
                except main.DailyJobError:
                    assert failed != 'none'
            assert (Path(directory)/'backup.success.json').exists() == (failed != 'backup')
            assert (Path(directory)/'application.success.json').exists() == (failed != 'application')


def test_alert_has_only_operator_recipient_and_fixed_diagnostics():
    sent = []
    connections, authentications = [], []
    class SMTP:
        def __init__(self,*args,**kwargs): connections.append((args, kwargs))
        def __enter__(self): return self
        def __exit__(self,*args): pass
        def login(self,*args): authentications.append(args)
        def send_message(self,message): sent.append(message)
    source = messages.CATALOG['operations.alert']
    template = replace(source, subject='Synthetic operator source', email='SOURCE\n' + source.email)
    occurred_at = datetime(2026, 9, 17, 15, 41, 2, tzinfo=timezone.utc)
    with patch('common.load_config',return_value={'club':{'email':{'smtpserver':'smtp.test','mail_ID':'synthetic','mail_password':'secret-canary'}}}), \
            patch.dict(messages.CATALOG, {'operations.alert': template}), \
            patch('operations.datetime') as clock, \
            patch.object(notifier.Notifier, 'send_account_message', side_effect=AssertionError('no member dispatch')):
        clock.now.return_value = occurred_at
        operations.send_alert(settings(),['test', 'database', 'backup'],smtp=SMTP)
        clock.now.assert_called_once_with(timezone.utc)
    assert len(sent) == 1
    assert sent[0]['To'] == settings()['alert_to']
    assert sent[0]['From'] == settings()['alert_from']
    assert sent[0]['Subject'] == 'Synthetic operator source', 'operator copy must use the shared source'
    assert 'secret-canary' not in sent[0].as_string()
    body = sent[0].get_content()
    assert body.startswith('SOURCE\n')
    assert 'Komponenten: backup, database, test' in body
    assert 'Zeit: 2026-09-17T15:41:02+00:00' in body
    assert 'Anleitung: deploy/OPERATIONS.md (Störungen)' in body
    assert connections[0][0] == ('smtp.test',)
    assert connections[0][1]['timeout'] == 15
    assert connections[0][1]['context'].check_hostname
    assert authentications == [('synthetic', 'secret-canary')]


def test_operator_rendering_failure_never_connects_or_discloses_context():
    invalid = replace(messages.CATALOG['operations.alert'], subject='private-canary\nInvalid header')
    smtp = Mock(side_effect=AssertionError('no SMTP connection'))
    with patch('common.load_config',return_value={'club':{'email':{'smtpserver':'smtp.test','mail_ID':'synthetic','mail_password':'secret-canary'}}}), \
            patch.dict(messages.CATALOG, {'operations.alert': invalid}):
        try:
            operations.send_alert(settings(), ['database'], smtp=smtp)
        except messages.MessageRenderError as error:
            assert 'private-canary' not in str(error) and 'secret-canary' not in str(error)
        else:
            raise AssertionError('an invalid operator subject must be refused before delivery')
    smtp.assert_not_called()


def test_launch_gate_requires_officer_decision_and_core_checks():
    cfg, pol, leg = settings(), policy(), legal()
    data=dict(schema_version=2,decision='go',operator='Synthetic operator',
              reviewed_at='2026-01-01T00:00:00+00:00',
              checks={key:True for key in p.LAUNCH_CHECKS})
    assert p.validate_launch(data,cfg,pol,leg)
    for key in p.LAUNCH_CHECKS:
        bad=copy.deepcopy(data);bad['checks'][key]=False
        try: p.validate_launch(bad,cfg,pol,leg)
        except p.ConfigurationError: pass
        else: raise AssertionError('missing launch check accepted')
    bad=copy.deepcopy(data);bad['decision']='no-go'
    try: p.validate_launch(bad,cfg,pol,leg)
    except p.ConfigurationError: pass
    else: raise AssertionError('no-go decision accepted')


if __name__ == '__main__':
    h.run_all(globals())
