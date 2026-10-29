"""Offline checks for reviewed, non-secret release assets."""
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile

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


SHIMS = {
    'main.py': 'nuligahelper.jobs.daily', 'manage_db.py': 'nuligahelper.cli',
    'operations.py': 'nuligahelper.ops.monitoring', 'privacy.py': 'nuligahelper.ops.privacy',
    'runtime_check.py': 'nuligahelper.ops.runtime_check',
    'monitor_watchdog.py': 'nuligahelper.ops.watchdog',
}


def test_transition_shims_work_from_a_release_tree():
    # Installed units run "<release>/venv/bin/python -B <release>/<shim>.py" with the
    # release as working directory; the shims must reach the package from there.
    for name, module in SHIMS.items():
        source = (PROJECT / name).read_text()
        assert f'runpy.run_module("{module}", run_name="__main__", alter_sys=True)' in source, name
    with tempfile.TemporaryDirectory() as directory:
        release = Path(directory) / 'release'
        release.mkdir()
        shutil.copytree(PROJECT / 'nuligahelper', release / 'nuligahelper',
                        ignore=shutil.ignore_patterns('__pycache__'))
        for name in [*SHIMS, 'webapp.py', 'schema_migrations.py', 'alembic.ini']:
            shutil.copy(PROJECT / name, release / name)
        environment = {key: value for key, value in os.environ.items()
                       if not key.startswith(('NULIGAHELPER_', 'PYTHON'))}

        def run(*args, **extra):
            return subprocess.run([sys.executable, '-B', *args], cwd=release, text=True,
                                  capture_output=True, env={**environment, **extra}, timeout=60)

        result = run('runtime_check.py')
        assert result.returncode == 1 and result.stderr.strip() == 'NULIGAHELPER_SECRET', result
        result = run('manage_db.py', '--help')
        assert result.returncode == 0 and 'migrate-schema' in result.stdout, result.stderr
        result = run('operations.py', 'config-check', '--config',
                     str(PROJECT / 'tests/fixtures/operations.json'))
        assert result.stdout.strip() == 'operations=valid', result
        result = run('privacy.py', '--help')
        assert result.returncode == 0 and '--apply' in result.stdout, result.stderr
        result = run('-c', 'import schema_migrations; print(schema_migrations.head_revisions()[0])')
        assert result.returncode == 0 and re.fullmatch(r'\d{4}_\w+', result.stdout.strip()), result
        database = Path(directory) / 'web.db'
        init = run('-m', 'nuligahelper.cli', '--db', str(database), 'init')
        assert init.returncode == 0, init.stderr
        result = run('-c', 'import webapp; print(webapp.app.name)',
                     NULIGAHELPER_DB=str(database), NULIGAHELPER_SECRET='synthetic-shim-secret')
        assert result.returncode == 0 and result.stdout.strip() == 'nuligahelper.web', result.stderr
        for name in ('main.py', 'monitor_watchdog.py'):
            result = run('-c', f'import importlib.util; print(importlib.util.find_spec({SHIMS[name]!r}) is not None)')
            assert result.stdout.strip() == 'True', (name, result.stderr)


if __name__ == '__main__':
    h.run_all(globals())
