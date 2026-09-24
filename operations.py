"""Local production preflight and monitoring. No network work at import time."""
from __future__ import annotations

import argparse
from datetime import datetime, timezone
from email.message import EmailMessage
import grp
import json
import logging
import os
from pathlib import Path
import pwd
import smtplib
import socket
import ssl
import stat
import subprocess

import common
import db
import production as p


UNITS = ('nuligahelper-web.service', 'nuligahelper-daily.service',
         'nuligahelper-daily.timer', 'nuligahelper-cleanup.service',
         'nuligahelper-cleanup.timer', 'nuligahelper-monitor.timer')


def command(*args):
    result = subprocess.run(args, capture_output=True, text=True, timeout=10, check=False)
    if result.returncode:
        raise p.ConfigurationError('host_command')
    return result.stdout.strip()


def certificate_days(host, now=None):
    context = ssl.create_default_context()
    with socket.create_connection((host, 443), timeout=5) as raw:
        with context.wrap_socket(raw, server_hostname=host) as tls:
            # Default context verifies hostname, chain and current validity.
            expires = ssl.cert_time_to_seconds(tls.getpeercert()['notAfter'])
    return (expires - (now or datetime.now(timezone.utc)).timestamp()) / 86400


def threshold(value, warning, critical, low=True):
    if low:
        return 2 if value < critical else 1 if value < warning else 0
    return 2 if value > critical else 1 if value > warning else 0


def monitor(config, *, runner=command, cert=certificate_days, statvfs=os.statvfs, now=None):
    """Return fixed identifiers and severity (0 healthy, 1 warning, 2 critical)."""
    results = {}
    for unit in UNITS:
        try:
            properties = dict(line.split('=', 1) for line in runner(
                'systemctl', 'show', unit, '--property=LoadState,ActiveState,Result,UnitFileState'
            ).splitlines() if '=' in line)
            healthy = properties.get('LoadState') == 'loaded'
            healthy &= properties.get('Result', 'success') == 'success'
            if unit.endswith('.timer'):
                healthy &= properties.get('ActiveState') == 'active'
                healthy &= properties.get('UnitFileState') == 'enabled'
            elif unit.endswith('web.service'):
                healthy &= properties.get('ActiveState') == 'active'
            else:
                healthy &= properties.get('ActiveState') in {'active', 'inactive', 'activating'}
            results[unit] = 0 if healthy else 2
        except (OSError, ValueError, subprocess.SubprocessError):
            results[unit] = 2
    for component in ('application', 'backup', 'cleanup'):
        try:
            age = p.success_age(component, config['state_dir'], now)
            results[component] = 2 if age > config['stale_seconds'] else 0
        except (OSError, ValueError):
            results[component] = 2
    try:
        results['clock'] = 0 if runner('timedatectl', 'show', '--property=NTPSynchronized', '--value') == 'yes' else 2
    except (OSError, ValueError, subprocess.SubprocessError):
        results['clock'] = 2
    # Application, configuration, state and journals may use separate filesystems.
    for label, path in [('app', config['app_dir']), ('config', config['config_dir']),
                        ('state', config['state_dir']), ('journal', '/var/log')]:
        try:
            info = statvfs(path)
            for key, free, total in [('disk', info.f_bavail, info.f_blocks),
                                     ('inodes', info.f_favail, info.f_files)]:
                results[label + '_' + key] = threshold(100 * free / total,
                    config[key + '_warning'], config[key + '_critical']) if total else 2
        except (OSError, ValueError, ZeroDivisionError):
            results[label + '_disk'] = results[label + '_inodes'] = 2
    try:
        results['certificate'] = threshold(cert(config['hostname']),
            config['certificate_warning'], config['certificate_critical'])
    except (OSError, ValueError, KeyError):
        results['certificate'] = 2
    results['database'] = 0 if db.database_ready(config['database']) else 2
    return results


def permissions(config):
    """Inspect owners/modes and parent traversal; never print file contents."""
    user = pwd.getpwnam(config['service_user'])
    group = grp.getgrnam(config['service_group'])
    errors = []
    if user.pw_uid == 0 or user.pw_gid != group.gr_gid or user.pw_shell not in {'/usr/sbin/nologin', '/sbin/nologin', '/bin/false'}:
        errors.append('service_identity')
    if any(g.gr_gid != group.gr_gid and user.pw_name in g.gr_mem for g in grp.getgrall()):
        errors.append('service_groups')
    expected = [(config['app_dir'], 0, group.gr_gid, 0o750),
                (config['config_dir'], 0, group.gr_gid, 0o750),
                (config['config_file'], 0, group.gr_gid, 0o640),
                (config['environment_file'], 0, 0, 0o600),
                (config['state_dir'], user.pw_uid, group.gr_gid, 0o700),
                (config['database'], user.pw_uid, group.gr_gid, 0o600)]
    for variable in ('NULIGAHELPER_OPERATIONS', 'NULIGAHELPER_POLICY', 'NULIGAHELPER_LEGAL'):
        if os.environ.get(variable):
            expected.append((os.environ[variable], 0, group.gr_gid, 0o640))
    secret = Path(config['config_dir']) / '.nuligahelper_secret'
    if secret.exists():
        expected.append((str(secret), 0, group.gr_gid, 0o640))
    for folder, owner, gid in [(config['app_dir'], 0, group.gr_gid),
                               (config['state_dir'], user.pw_uid, group.gr_gid)]:
        errors.extend(tree_permission_errors(Path(folder), owner, gid,
                                             state=folder == config['state_dir']))
    for filename, uid, gid, mode in expected:
        path = Path(filename)
        try:
            info = path.lstat()
            if path.is_symlink() or (info.st_uid, info.st_gid, stat.S_IMODE(info.st_mode)) != (uid, gid, mode):
                errors.append('path_permissions:' + str(path))
            for parent in path.parents:
                st = parent.stat()
                if st.st_uid != 0 and parent != Path(config['state_dir']):
                    errors.append('parent_owner')
                if stat.S_IMODE(st.st_mode) & 0o022:
                    errors.append('parent_writable')
        except OSError:
            errors.append('path_unavailable:' + str(path))
    return sorted(set(errors))


def tree_permission_errors(folder: Path, owner: int, gid: int, *, state: bool = False):
    """Inspect descendants, including release links, without following links."""
    errors = []
    for root, dirs, files in os.walk(folder, followlinks=False):
        for name in dirs + files:
            path = Path(root) / name
            if path.is_symlink():
                try:
                    target = path.resolve(strict=True)
                except (OSError, RuntimeError):
                    errors.append('symlink')
                    continue
                # Venv interpreter links may resolve into the root-owned OS
                # installation; release links must remain inside app root.
                if state or not (target.is_relative_to('/usr') or
                                 target.is_relative_to(folder)):
                    errors.append('symlink')
                continue
            info = path.stat()
            if info.st_uid != owner or info.st_gid != gid:
                errors.append('tree_owner')
            mode = stat.S_IMODE(info.st_mode)
            if state and mode != (0o700 if path.is_dir() else 0o600):
                errors.append('state_mode')
            if not state and mode & 0o027:
                errors.append('app_mode')
    return errors


def send_alert(config, components, *, smtp=smtplib.SMTP_SSL):
    """Dedicated operator-only mail path; never uses member notification dispatch."""
    allowed = set(UNITS) | {'application', 'backup', 'cleanup', 'monitor', 'clock',
        'certificate', 'database', 'test'} | {a + '_' + b for a in ('app', 'config', 'state', 'journal') for b in ('disk', 'inodes')}
    p.require(bool(components) and set(components) <= allowed, 'alert.components')
    email = common.load_config(config['config_file'])['club']['email']
    msg = EmailMessage()
    msg['From'], msg['To'] = config['alert_from'], config['alert_to']
    msg['Subject'] = 'nuLigaHelper: Betriebsmeldung'
    msg.set_content('Komponenten: ' + ', '.join(sorted(components)) + '\nZeit: ' +
        datetime.now(timezone.utc).isoformat() + '\nAnleitung: deploy/OPERATIONS.md (Störungen)\n')
    with smtp(email['smtpserver'], timeout=15, context=ssl.create_default_context()) as server:
        server.login(email['mail_ID'], email['mail_password'])
        server.send_message(msg)


def validate_documents(config):
    policy = p.validate_policy(p.read_json(os.environ.get('NULIGAHELPER_POLICY')))
    legal = p.validate_legal(p.read_json(os.environ.get('NULIGAHELPER_LEGAL')))
    from privacy import validate_cleanup
    validate_cleanup(policy)
    return policy, legal


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('action', choices=['config-check', 'permissions', 'preflight', 'monitor', 'alert-test', 'alert', 'launch-check'])
    parser.add_argument('--config', default=os.environ.get('NULIGAHELPER_OPERATIONS', 'deploy/operations.json'))
    parser.add_argument('--component', choices=UNITS + ('application', 'backup', 'cleanup', 'monitor'))
    parser.add_argument('--notify', action='store_true')
    args = parser.parse_args(argv)
    try:
        config = p.validate_operations(p.read_json(args.config))
        if args.action == 'config-check':
            print('operations=valid')
            return 0
        if args.action == 'permissions':
            errors = permissions(config)
            print(json.dumps({'permissions': errors or ['ok']}))
            return 2 if errors else 0
        if args.action in {'alert-test', 'alert'}:
            p.require(args.action == 'alert-test' or args.component is not None, 'component')
            send_alert(config, ['test' if args.action == 'alert-test' else args.component])
            print('alert=sent')
            return 0
        if args.action in {'preflight', 'launch-check'}:
            from runtime_check import check
            error = check()
            p.require(error is None, error or 'runtime')
            for variable, key in [('NULIGAHELPER_DB', 'database'), ('NULIGAHELPER_CONFIG', 'config_file'),
                                  ('NULIGAHELPER_STATE_DIR', 'state_dir'), ('NULIGAHELPER_TRUSTED_HOSTS', 'hostname')]:
                p.require(os.environ.get(variable) == config[key], variable)
            p.require(common.DEBUG_FLAG is False and common.CHANGE_DAY is False, 'debug_flags')
            common.load_config(config['config_file'])
            policy, legal = validate_documents(config)
            if args.action == "launch-check":
                p.validate_launch(p.read_json(os.environ.get("NULIGAHELPER_LAUNCH")), config, policy, legal)
            p.require(not permissions(config), 'permissions')
        results = monitor(config)
        print(json.dumps(results, sort_keys=True))
        code = max(results.values(), default=0)
        if args.notify and code:
            send_alert(config, [key for key, value in results.items() if value])
        # A completed monitor run does not imply components were healthy.
        p.write_success('monitor', config['state_dir'])
        return code
    except Exception as error:
        # Never expose config payloads, host-command stderr or SMTP exceptions.
        reason = str(error) if isinstance(error, p.ConfigurationError) else type(error).__name__
        print('operation=failed field=' + reason)
        return 2


if __name__ == '__main__':
    raise SystemExit(main())
