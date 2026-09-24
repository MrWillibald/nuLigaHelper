"""Offline boundary checks for host tools and service configuration."""
from datetime import datetime, timezone
import os
from pathlib import Path
import ssl
import subprocess
import tempfile
from types import SimpleNamespace
from unittest.mock import patch

import helpers as h
import operations
import production as p
import runtime_check
import monitor_watchdog
from test_operations import settings


def test_runtime_preflight_reports_only_missing_or_invalid_setting():
    cfg = settings()
    env=dict(NULIGAHELPER_SECRET='REPLACE_WITH_SECRET',NULIGAHELPER_CONFIG=cfg['config_file'],
        NULIGAHELPER_DB=cfg['database'],NULIGAHELPER_STATE_DIR=cfg['state_dir'],
        NULIGAHELPER_OPERATIONS='synthetic.json',NULIGAHELPER_TRUSTED_HOSTS=cfg['hostname'],TZ='UTC')
    with patch.dict(os.environ,env,clear=True):
        assert runtime_check.check() == 'NULIGAHELPER_SECRET'
        os.environ['NULIGAHELPER_SECRET']='synthetic-secret-for-test'
        with patch('production.read_json',return_value=cfg), patch('common.load_config',side_effect=ValueError('secret-canary')):
            assert runtime_check.check() == 'NULIGAHELPER_CONFIG'


def test_tls_check_uses_verified_context_and_public_hostname():
    class Connection:
        def __enter__(self): return self
        def __exit__(self,*args): pass
    class TLS(Connection):
        def getpeercert(self): return {'notAfter':'Sep 30 00:00:00 2026 GMT'}
    context = ssl.create_default_context()
    assert context.check_hostname and context.verify_mode == ssl.CERT_REQUIRED
    with patch('operations.socket.create_connection',return_value=Connection()), patch.object(context,'wrap_socket',return_value=TLS()) as wrap, patch('operations.ssl.create_default_context',return_value=context):
        assert operations.certificate_days('club.test',datetime(2026,9,16,tzinfo=timezone.utc)) == 14
        assert wrap.call_args.kwargs['server_hostname'] == 'club.test'
    for category in ('expired','hostname mismatch','untrusted'):
        with patch('operations.socket.create_connection',return_value=Connection()), patch.object(context,'wrap_socket',side_effect=ssl.SSLCertVerificationError(category)), patch('operations.ssl.create_default_context',return_value=context):
            try: operations.certificate_days('club.test')
            except ssl.SSLCertVerificationError: pass
            else: raise AssertionError('verification failure swallowed')


def test_watchdog_detects_stale_or_missing_monitor_and_uses_only_alert_unit():
    for age in (0,1801):
        with patch('production.success_age',return_value=age), patch('monitor_watchdog.subprocess.run',return_value=SimpleNamespace(returncode=0)) as run:
            assert monitor_watchdog.main() == 0
            assert run.called == (age > 1800)
            if run.called: assert run.call_args.args[0] == ['systemctl','start','nuligahelper-alert@monitor.service']
    with patch('production.success_age',side_effect=p.ConfigurationError('marker')), patch('monitor_watchdog.subprocess.run',return_value=SimpleNamespace(returncode=1)):
        assert monitor_watchdog.main() == 2


def test_unit_syntax_with_synthetic_executable_paths():
    # The project lives outside /opt here. Validate copied units with existing
    # harmless executables; this does not claim installed-host execution evidence.
    with tempfile.TemporaryDirectory() as directory:
        files=[]
        for file in (Path(h.PROJECT_DIR)/'release-assets/systemd').glob('nuligahelper-*'):
            if file.suffix not in {'.service','.timer'}: continue
            content=file.read_text().replace('/opt/nuligahelper/current/venv/bin/python','/usr/bin/python3')
            target=Path(directory)/file.name
            target.write_text(content)
            files.append(str(target))
        result=subprocess.run(['systemd-analyze','verify',*files],capture_output=True,text=True)
        assert result.returncode == 0, result.stderr
    timer=(Path(h.PROJECT_DIR)/'release-assets/systemd/nuligahelper-daily.timer').read_text()
    assert '09:00:00 Europe/Berlin' in timer and 'Persistent=true' in timer


if __name__ == '__main__':
    h.run_all(globals())
