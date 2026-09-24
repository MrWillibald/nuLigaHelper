"""Offline checks for reviewed, non-secret release assets."""
from pathlib import Path
import re

import helpers as h


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
    assert '/opt/nuligahelper/current/monitor_watchdog.py' in (
        ASSETS / 'nuligahelper-watchdog.cron').read_text()

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


if __name__ == '__main__':
    h.run_all(globals())
