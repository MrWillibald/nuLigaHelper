"""Offline checks for reviewed, non-secret release assets."""
import os
from pathlib import Path
import re
import shutil
import socket
import subprocess
import sys
import tempfile
import time
import urllib.request

from tests import helpers as h


PROJECT = Path(h.PROJECT_DIR)
ASSETS = PROJECT / 'release-assets'


def test_versioned_deployment_assets_are_complete_and_placeholder_only():
    requirements = [line.strip() for line in
                    (PROJECT / 'requirements-production.txt').read_text().splitlines()
                    if line.strip() and not line.startswith('#')]
    assert requirements and all(re.fullmatch(r'[A-Za-z0-9_.-]+==[A-Za-z0-9_.+-]+', line)
                                for line in requirements), 'production dependencies must be pinned'
    for package in ('gunicorn', 'sqlalchemy', 'alembic', 'flask', 'pandas'):
        assert any(line.lower().startswith(package + '==') for line in requirements), package

    gunicorn = (ASSETS / 'gunicorn.conf.py').read_text()
    for setting in ('bind = "127.0.0.1:8080"', 'workers = 1',
                    'worker_class = "sync"', 'accesslog = "-"',
                    'errorlog = "-"', 'limit_request_line'):
        assert setting in gunicorn

    assert (ASSETS / 'nuligahelper-deploy.py').is_file()
    assert (ASSETS / 'recovery_check.py').is_file()

    environment = (ASSETS / 'nuligahelper.env.example').read_text()
    for key in ('NULIGAHELPER_ENV=production', 'NULIGAHELPER_SECRET=REPLACE_',
                'NULIGAHELPER_DB=/var/lib/nuligahelper/',
                'NULIGAHELPER_TRUSTED_HOSTS='):
        assert key in environment
    assert 'example.invalid' in environment
    assert 'sk_' not in environment

    units = list((ASSETS / 'systemd').glob('nuligahelper-*.service'))
    assert len(units) == 9
    for path in units:
        text = path.read_text()
        assert 'WorkingDirectory=/opt/nuligahelper/current' in text, path.name
        assert '/opt/nuligahelper/venv/' not in text, path.name
        assert '/opt/nuligahelper/current/venv/' in text, path.name
        assert 'EnvironmentFile=/etc/nuligahelper/web.env' in text, path.name
        assert 'ReadWritePaths=/opt/nuligahelper' not in text, path.name
    web_unit = (ASSETS / 'systemd/nuligahelper-web.service').read_text()
    assert '/opt/nuligahelper/current/release-assets/gunicorn.conf.py' in web_unit
    assert 'ReadWritePaths=/var/lib/nuligahelper' in web_unit
    assert ('root cd /opt/nuligahelper/current && /opt/nuligahelper/current/venv/bin/python '
            '-m nuligahelper.ops.watchdog') in (ASSETS / 'nuligahelper-watchdog.cron').read_text()

    caddy = (ASSETS / 'Caddyfile.example').read_text()
    for directive in ('reverse_proxy 127.0.0.1:8080', 'max_size 1MB',
                      'header_up -Forwarded', 'header_up X-Forwarded-For {remote_host}',
                      'header_up X-Forwarded-Proto https',
                      'header_up X-Forwarded-Host {host}',
                      'Strict-Transport-Security', 'Content-Security-Policy'):
        assert directive in caddy
    for header in ('X-Forwarded-For', 'X-Forwarded-Proto', 'X-Forwarded-Host'):
        assert f'header_up -{header}' not in caddy

    guide = (ASSETS / 'README.md').read_text()
    for phrase in ('/opt/nuligahelper/current', '/etc/nuligahelper/web.env',
                   '/var/lib/nuligahelper', 'systemd-analyze verify',
                   'persistent', 'snapshot'):
        assert phrase.lower() in guide.lower(), phrase



def test_units_start_package_modules():
    for path in (ASSETS / 'systemd').glob('nuligahelper-*.service'):
        for line in path.read_text().splitlines():
            if line.startswith(('ExecStart=', 'ExecStartPre=')) and '/venv/bin/python' in line:
                assert ' -B -m nuligahelper.' in line, f'{path.name}: {line}'
    web_unit = (ASSETS / 'systemd/nuligahelper-web.service').read_text()
    assert web_unit.rstrip().count('nuligahelper.web:app') == 1 and ' webapp:app' not in web_unit
    for script, module in (('run.sh', 'exec python -m nuligahelper.jobs.daily'),
                           ('run_webapp.sh', 'python -m nuligahelper.web')):
        assert module in (PROJECT / script).read_text(), script


def _serve_health(release, target, environment):
    """Start Gunicorn like the web unit does; return the /healthz body and its log."""
    with socket.socket() as probe:
        probe.bind(('127.0.0.1', 0))
        port = probe.getsockname()[1]
    server = subprocess.Popen(
        [sys.executable, '-B', '-m', 'gunicorn', '--config',
         str(release / 'release-assets/gunicorn.conf.py'), '--bind', f'127.0.0.1:{port}', target],
        cwd=release, env=environment, text=True,
        stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
    try:
        deadline = time.monotonic() + 60
        while True:
            if server.poll() is not None:
                raise AssertionError(f'{target} exited before serving: {server.stderr.read()}')
            try:
                with urllib.request.urlopen(f'http://127.0.0.1:{port}/healthz', timeout=5) as response:
                    body = response.read()
                break
            except OSError:
                if time.monotonic() > deadline:
                    raise
                time.sleep(0.2)
    finally:
        server.terminate()
        log = server.communicate(timeout=30)[1]
    return body, log


def test_release_gunicorn_config_serves_the_package_web_target():
    # The installed web unit loads <release>/release-assets/gunicorn.conf.py; its logging
    # config resolves module paths at start-up, which --check-config does not exercise.
    with tempfile.TemporaryDirectory() as directory:
        release = Path(directory) / 'release'
        release.mkdir()
        shutil.copytree(PROJECT / 'nuligahelper', release / 'nuligahelper',
                        ignore=shutil.ignore_patterns('__pycache__'))
        shutil.copytree(ASSETS, release / 'release-assets',
                        ignore=shutil.ignore_patterns('__pycache__'))
        shutil.copy(PROJECT / 'alembic.ini', release / 'alembic.ini')
        database = Path(directory) / 'web.db'
        environment = {key: value for key, value in os.environ.items()
                       if not key.startswith(('NULIGAHELPER_', 'PYTHON'))}
        init = subprocess.run([sys.executable, '-B', '-m', 'nuligahelper.cli', '--db', str(database),
                               'init'], cwd=release, env=environment, capture_output=True,
                              text=True, timeout=60)
        assert init.returncode == 0, init.stderr
        environment.update(NULIGAHELPER_DB=str(database), NULIGAHELPER_SECRET='synthetic-web-secret')
        body, log = _serve_health(release, 'nuligahelper.web:app', environment)
        assert body == b'ok\n', log
        assert 'level=INFO operation=web event=diagnostic' in log and '127.0.0.1' not in log, (
            f'Gunicorn must log through the redacting journal formatter: {log}')


if __name__ == '__main__':
    h.run_all(globals())
