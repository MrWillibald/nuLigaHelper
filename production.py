"""Validated deployment data and non-sensitive operation state.

No production credentials are read at import time. Errors identify fields only.
"""
from __future__ import annotations

import hashlib
import json
import logging
import math
import os
from pathlib import Path
import re
import tempfile
from datetime import datetime, timezone
from urllib.parse import urlsplit


class ConfigurationError(ValueError):
    pass


def require(condition, field):
    if not condition:
        raise ConfigurationError(field)


def text(value, field):
    require(isinstance(value, str) and bool(value.strip()), field)
    require(not re.search(r'(?i)(\bOFFEN\b|\bTODO\b|\bTBD\b|REPLACE_|example\.invalid|\bPLACEHOLDER\b)', value), field)
    return value


def number(value, field, minimum=0):
    require(type(value) in (int, float) and math.isfinite(value) and value >= minimum, field)
    return value


def read_json(path):
    try:
        return json.loads(Path(path).read_text(encoding='utf-8'))
    except (OSError, ValueError, TypeError):
        raise ConfigurationError('deployment_file') from None


def digest(value):
    return hashlib.sha256(json.dumps(value, sort_keys=True, ensure_ascii=False,
                                     separators=(',', ':')).encode()).hexdigest()


def approval(document, payload):
    a = document.get('approval', {})
    require(isinstance(a, dict), 'approval')
    for key in ('operator', 'reviewer', 'approved_at', 'evidence', 'version'):
        text(a.get(key), 'approval.' + key)
    require(a['version'] == document.get('version'), 'approval.version')
    require(a.get('sha256') == digest(payload), 'approval.sha256')
    try:
        stamp = datetime.fromisoformat(a['approved_at'])
        require(stamp.tzinfo is not None and stamp <= datetime.now(timezone.utc), 'approval.approved_at')
    except ValueError:
        raise ConfigurationError('approval.approved_at') from None


def validate_operations(data):
    require(isinstance(data, dict), 'operations')
    require(data.get('schema_version') == 1, 'schema_version')
    for key in ('app_dir', 'config_dir', 'state_dir', 'database', 'config_file', 'environment_file'):
        value = text(data.get(key), key)
        require(Path(value).is_absolute() and '..' not in Path(value).parts, key)
    require(Path(data['database']).parent == Path(data['state_dir']), 'database')
    require(Path(data['config_file']).parent == Path(data['config_dir']), 'config_file')
    for key in ('service_user', 'service_group'):
        require(bool(re.fullmatch('[a-z_][a-z0-9_-]*', text(data.get(key), key))), key)
        require(data[key] != 'root', key)
    host = text(data.get('hostname'), 'hostname')
    require(bool(re.fullmatch(r'[a-z0-9]+(?:[a-z0-9.-]*[a-z0-9])?', host)) and '.' in host, 'hostname')
    import contact_validation as contacts
    for key in ('alert_from', 'alert_to'):
        try:
            require(contacts.normalize_email(text(data.get(key), key)) == data[key], key)
        except ValueError:
            raise ConfigurationError(key) from None
    text(data.get('operator'), 'operator')
    for key in ('stale_seconds', 'monitor_seconds', 'observation_days', 'journal_max_bytes', 'journal_days'):
        number(data.get(key), key, 1)
    for category, upper in (('disk', 100), ('inodes', 100), ('certificate', None)):
        warn = number(data.get(category + '_warning'), category + '_warning', 1)
        crit = number(data.get(category + '_critical'), category + '_critical', 1)
        require(crit < warn and (upper is None or warn < upper), category)
    require(data.get('timezone') == 'Europe/Berlin', 'timezone')
    require(data.get('daily_time') == '09:00', 'daily_time')
    return data


DATA_CLASSES = ('auth_tokens', 'abuse_counters', 'sessions', 'registered', 'verified',
                'rejected', 'registration_evidence', 'active_persons', 'inactive_persons',
                'assignments', 'audits', 'journals', 'local_snapshots', 'dropbox_dated',
                'dropbox_latest', 'restore_copies', 'operator_alerts')
PROCESSORS = ('netcup', 'ionos', 'twilio', 'dropbox', 'dns', 'tls', 'alert_mailbox')


def validate_policy(data):
    require(isinstance(data, dict) and data.get('schema_version') == 1, 'policy.schema_version')
    require(bool(re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9._-]{0,79}', text(data.get('version'), 'policy.version'))), 'policy.version')
    classes = data.get('data_classes', {})
    require(isinstance(classes, dict) and set(classes) == set(DATA_CLASSES), 'policy.data_classes')
    for name, entry in classes.items():
        require(isinstance(entry, dict), 'policy.' + name)
        for key in ('owner', 'purpose', 'legal_basis', 'location', 'access', 'trigger', 'period', 'disposition'):
            text(entry.get(key), 'policy.' + name + '.' + key)
    require(classes['audits']['disposition'] == 'preserve_append_only', 'policy.audits.disposition')
    processors = data.get('processors', {})
    require(isinstance(processors, dict) and set(processors) == set(PROCESSORS), 'policy.processors')
    for name, entry in processors.items():
        require(isinstance(entry, dict) and type(entry.get('enabled')) is bool, 'processor.' + name)
        for key in ('owner', 'purpose', 'data', 'transfer', 'locations', 'review', 'retention', 'incident_contact'):
            text(entry.get(key), 'processor.' + name + '.' + key)
    approval(data, {k: v for k, v in data.items() if k != 'approval'})
    return data


def validate_legal(data):
    require(isinstance(data, dict) and data.get('schema_version') == 1, 'legal.schema_version')
    text(data.get('version'), 'legal.version')
    text(data.get('policy_sha256'), 'legal.policy_sha256')
    pages = data.get('pages', {})
    require(isinstance(pages, dict) and set(pages) == {'impressum', 'datenschutz'}, 'legal.pages')
    for name, page in pages.items():
        require(isinstance(page, dict), 'legal.page')
        text(page.get('title'), 'legal.title')
        blocks = page.get('blocks')
        require(isinstance(blocks, list) and 0 < len(blocks) <= 500, 'legal.blocks')
        for block in blocks:
            require(isinstance(block, dict), 'legal.block')
            kind = block.get('type')
            require(kind in {'heading', 'paragraph', 'list', 'link'}, 'legal.block.type')
            if kind == 'list':
                require(isinstance(block.get('items'), list) and bool(block['items']), 'legal.items')
                for item in block['items']:
                    text(item, 'legal.item')
            else:
                text(block.get('text'), 'legal.text')
            if kind == 'link':
                url = text(block.get('url'), 'legal.url')
                parsed = urlsplit(url)
                require(parsed.scheme in {'https', 'mailto'} and not any(c.isspace() for c in url), 'legal.url')
                require(bool(parsed.netloc) if parsed.scheme == 'https' else bool(parsed.path), 'legal.url')
    approval(data, {k: v for k, v in data.items() if k != 'approval'})
    return data


COMPONENTS = {'application', 'backup', 'cleanup', 'monitor'}


def write_success(component, directory=None, now=None):
    require(component in COMPONENTS, 'component')
    directory = directory or os.environ.get('NULIGAHELPER_STATE_DIR')
    if not directory:
        return  # Local development has no production markers.
    stamp = now or datetime.now(timezone.utc)
    require(stamp.tzinfo is not None, 'timestamp')
    folder = Path(directory)
    descriptor, temporary = tempfile.mkstemp(prefix='.success-', dir=folder)
    try:
        with os.fdopen(descriptor, 'w') as stream:
            json.dump({'schema_version': 1, 'completed_at': stamp.isoformat()}, stream)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, folder / (component + '.success.json'))
        fd = os.open(folder, os.O_RDONLY | os.O_DIRECTORY)
        try:
            os.fsync(fd)
        finally:
            os.close(fd)
    finally:
        Path(temporary).unlink(missing_ok=True)


def success_age(component, directory, now=None):
    require(component in COMPONENTS, 'component')
    data = read_json(Path(directory) / (component + '.success.json'))
    try:
        stamp = datetime.fromisoformat(data['completed_at'])
        require(data.get('schema_version') == 1 and stamp.tzinfo is not None, 'marker')
        age = ((now or datetime.now(timezone.utc)) - stamp).total_seconds()
        require(age >= 0, 'marker.future')
        return age
    except (KeyError, TypeError, ValueError):
        raise ConfigurationError('marker') from None


def event(component, outcome, reason='none'):
    # Call sites supply static identifiers, never exception messages or payloads.
    require(bool(re.fullmatch('[a-z_]+', component)), 'event.component')
    require(outcome in {'success', 'failure', 'started', 'unavailable'}, 'event.outcome')
    require(bool(re.fullmatch('[A-Za-z_]+', reason)), 'event.reason')
    logging.getLogger('nuligahelper.operations').log(
        logging.ERROR if outcome in {'failure', 'unavailable'} else logging.INFO,
        'operation=%s outcome=%s reason=%s', component, outcome, reason)


LAUNCH_CHECKS = ('release', 'sibling_changes', 'permissions', 'secrets', 'backup_restore',
                'log_redaction_retention', 'health_alerts', 'timers_markers', 'clock_disk_tls',
                'privacy_processors', 'legal_content', 'observation', 'rollback')


def validate_launch(data, operations, policy, legal):
    require(isinstance(data, dict) and data.get('schema_version') == 1, 'launch.schema_version')
    text(data.get('version'), 'launch.version')
    require(data.get('decision') == 'go', 'launch.decision')
    text(data.get('operator'), 'launch.operator')
    for key, value in [('operations_sha256', operations), ('policy_sha256', policy), ('legal_sha256', legal)]:
        require(data.get(key) == digest(value), 'launch.' + key)
    checks = data.get('checks', {})
    require(isinstance(checks, dict) and set(checks) == set(LAUNCH_CHECKS), 'launch.checks')
    for name, check in checks.items():
        require(isinstance(check, dict) and check.get('complete') is True, 'launch.' + name)
        for key in ('owner', 'evidence', 'verified_at'):
            text(check.get(key), 'launch.' + name + '.' + key)
    return data
