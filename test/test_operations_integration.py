"""Operation outcomes, redaction, launch binding and operator-only alert tests."""
import copy
import io
import json
import logging
import os
from pathlib import Path
import tempfile
from unittest.mock import patch

import helpers as h
import backup
import main
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
    class SMTP:
        def __init__(self,*args,**kwargs): pass
        def __enter__(self): return self
        def __exit__(self,*args): pass
        def login(self,*args): pass
        def send_message(self,message): sent.append(message)
    with patch('common.load_config',return_value={'club':{'email':{'smtpserver':'smtp.test','mail_ID':'synthetic','mail_password':'secret-canary'}}}):
        operations.send_alert(settings(),['test'],smtp=SMTP)
    assert len(sent) == 1
    assert sent[0]['To'] == settings()['alert_to']
    assert sent[0]['From'] == settings()['alert_from']
    assert 'secret-canary' not in sent[0].as_string()
    assert 'deploy/OPERATIONS.md' in sent[0].get_content()


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
