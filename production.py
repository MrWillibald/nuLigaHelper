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


def validate_policy(data):
    require(isinstance(data, dict) and data.get('schema_version') == 2, 'policy.schema_version')
    require(bool(re.fullmatch(r'[A-Za-z0-9][A-Za-z0-9._-]{0,79}', text(data.get('version'), 'policy.version'))), 'policy.version')
    return data


def validate_legal(data):
    require(isinstance(data, dict) and data.get('schema_version') == 2, 'legal.schema_version')
    text(data.get('version'), 'legal.version')
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


LAUNCH_CHECKS = ('backup_restore', 'privacy_notice', 'retention')


def validate_launch(data, operations, policy, legal):
    require(isinstance(data, dict) and data.get('schema_version') == 2, 'launch.schema_version')
    require(data.get('decision') == 'go', 'launch.decision')
    text(data.get('operator'), 'launch.operator')
    try:
        reviewed = datetime.fromisoformat(text(data.get('reviewed_at'), 'launch.reviewed_at'))
        require(reviewed.tzinfo is not None and reviewed <= datetime.now(timezone.utc), 'launch.reviewed_at')
    except ValueError:
        raise ConfigurationError('launch.reviewed_at') from None
    checks = data.get('checks', {})
    require(isinstance(checks, dict) and set(checks) == set(LAUNCH_CHECKS), 'launch.checks')
    for name, check in checks.items():
        require(check is True, 'launch.' + name)
    return data
