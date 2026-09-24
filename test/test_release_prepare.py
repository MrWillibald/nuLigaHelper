"""Offline checks for operator-selected, non-activating release preparation."""

import importlib.util
import copy
import json
from datetime import datetime, timezone
from pathlib import Path
import subprocess
import sys
import tempfile
from types import SimpleNamespace
from unittest.mock import patch

import helpers as h
import db


SOURCE = Path(h.PROJECT_DIR) / 'release-assets/nuligahelper-deploy.py'
SPEC = importlib.util.spec_from_file_location('nuligahelper_release_deploy', SOURCE)
deploy = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(deploy)


def git(*args, cwd=None):
    return subprocess.run(('git', *args), cwd=cwd, text=True, capture_output=True,
                          check=True).stdout.strip()


def synthetic_source(root):
    source = root / 'source'
    source.mkdir()
    git('init', '-b', 'master', str(source))
    git('config', 'user.email', 'operator@club.test', cwd=source)
    git('config', 'user.name', 'Synthetic Operator', cwd=source)
    (source / 'release-assets').mkdir()
    (source / 'test').mkdir()
    (source / 'requirements-production.txt').write_text('')
    (source / 'requirements-test.txt').write_text('')
    (source / 'release-assets/gunicorn.conf.py').write_text('bind = "127.0.0.1:8080"\n')
    (source / 'release-assets/recovery_check.py').write_text('"""Synthetic helper."""\n')
    (source / 'test/run_tests.sh').write_text('#!/bin/bash\nexit 0\n')
    git('add', '.', cwd=source)
    git('commit', '-m', 'Synthetic master release', cwd=source)
    return source, git('rev-parse', 'HEAD', cwd=source)


def configuration(root, source):
    app_root = root / 'application'
    app_root.mkdir()
    prior = app_root / 'prior'
    prior.mkdir()
    (app_root / 'current').symlink_to(prior)
    return dict(source=str(source), app_root=str(app_root),
                source_cache=str(root / 'cache/source.git'),
                service_group='synthetic', python=sys.executable,
                database=str(root / 'live.db'), recovery_dir=str(root / 'recovery'),
                public_health_url='https://club.test/healthz')


def _synthetic_releases(app_root):
    releases = app_root / 'releases'
    releases.mkdir(exist_ok=True)
    return releases


def test_pinned_commit_must_be_on_fetched_master_even_after_branch_moves():
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        source, old = synthetic_source(root)
        cache = root / 'cache/source.git'
        with patch.object(deploy, 'check_cache_directory'):
            assert deploy.fetch_master(cache, str(source)) == old
        assert deploy.select_commit(cache, old, old)[0] == old
        git('checkout', '-b', 'feature', cwd=source)
        (source / 'feature.txt').write_text('not promoted\n')
        git('add', '.', cwd=source)
        git('commit', '-m', 'Unreviewed feature', cwd=source)
        feature = git('rev-parse', 'HEAD', cwd=source)
        git('-C', str(cache), 'fetch', 'origin', 'feature')
        try:
            deploy.select_commit(cache, old, feature)
        except deploy.DeployError as error:
            assert 'not reachable' in str(error)
        else:
            raise AssertionError('non-master commit accepted')
        git('checkout', 'master', cwd=source)
        (source / 'master.txt').write_text('reviewed next release\n')
        git('add', '.', cwd=source)
        git('commit', '-m', 'Advance master', cwd=source)
        with patch.object(deploy, 'check_cache_directory'):
            latest = deploy.fetch_master(cache, str(source))
        assert latest != old
        assert deploy.select_commit(cache, latest, old)[0] == old, \
            'moving master must not change an explicitly pinned candidate'


def test_preparation_failures_keep_current_and_services_untouched():
    for failed_stage in ('fetch', 'install', 'test'):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source, commit = synthetic_source(root)
            config = configuration(root, source)
            app_root = Path(config['app_root'])
            current = app_root / 'current'
            original = current.resolve()
            real_run = deploy.run
            with patch.object(deploy, 'check_headroom'), \
                    patch.object(deploy, 'check_cache_directory'), \
                    patch.object(deploy, 'ensure_releases_directory',
                                 side_effect=lambda app, group: _synthetic_releases(app)), \
                    patch.object(deploy, 'normalize_permissions'), \
                    patch.object(deploy, 'verify_unit_syntax'), \
                    patch.object(deploy, 'run') as command:
                def fail_command(*args, **kwargs):
                    if (failed_stage == 'fetch' and 'fetch' in args) or \
                            (failed_stage == 'install' and 'install' in args) or \
                            (failed_stage == 'test' and args[0] == 'bash'):
                        raise deploy.DeployError('synthetic ' + failed_stage + ' failure')
                    return real_run(*args, **kwargs)
                command.side_effect = fail_command
                try:
                    deploy.prepare(config, commit)
                except deploy.DeployError as error:
                    assert failed_stage + ' failure' in str(error)
                else:
                    raise AssertionError('failed ' + failed_stage + ' prepared a release')
            assert current.resolve() == original
            assert not (app_root / 'releases' / commit).exists()
            releases = app_root / 'releases'
            assert not releases.exists() or not list(releases.glob('.stage-*'))


def test_clean_preparation_records_exact_commit_without_switching_current():
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        source, commit = synthetic_source(root)
        config = configuration(root, source)
        app_root = Path(config['app_root'])
        original = (app_root / 'current').resolve()
        with patch.object(deploy, 'check_headroom'), \
                patch.object(deploy, 'check_cache_directory'), \
                patch.object(deploy, 'ensure_releases_directory',
                             side_effect=lambda app, group: _synthetic_releases(app)), \
                patch.object(deploy, 'normalize_permissions'), \
                patch.object(deploy, 'verify_unit_syntax'):
            record = deploy.prepare(config, commit)
        candidate = app_root / 'releases' / commit
        assert candidate.is_dir()
        assert record['commit'] == commit
        assert record['previous_release'] == str(original)
        assert (candidate / '.prepared.json').is_file()
        shebang = (candidate / 'venv/bin/pip').read_text().splitlines()[0]
        assert shebang.startswith('#!' + str(candidate / 'venv/bin/python')) and \
            Path(shebang[2:]).exists(), \
            'venv entry points must embed the final release path: ' + \
            shebang
        assert (app_root / 'current').resolve() == original
        assert not (candidate / '.git').exists()
        try:
            with patch.object(deploy, 'check_headroom'), \
                    patch.object(deploy, 'check_cache_directory'), \
                    patch.object(deploy, 'ensure_releases_directory',
                                 side_effect=lambda app, group: _synthetic_releases(app)):
                deploy.prepare(config, commit)
        except deploy.DeployError as error:
            assert 'already exists' in str(error)
        else:
            raise AssertionError('existing candidate was overwritten')


def test_host_lock_serializes_and_current_link_refuses_external_target():
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        lock = root / 'host.lock'
        first = deploy.exclusive_lock(lock)
        try:
            try:
                deploy.exclusive_lock(lock)
            except deploy.DeployError as error:
                assert 'host lock' in str(error)
            else:
                raise AssertionError('concurrent deployment accepted')
        finally:
            import os
            os.close(first)
        source, _ = synthetic_source(root)
        config = configuration(root, source)
        current = Path(config['app_root']) / 'current'
        current.unlink()
        current.symlink_to(source)
        try:
            deploy.prior_release(Path(config['app_root']))
        except deploy.DeployError as error:
            assert 'outside' in str(error)
        else:
            raise AssertionError('external current target accepted')


def test_host_lock_refuses_a_symlink_or_non_private_file():
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        other = root / 'other'
        other.write_text('untouched')
        link = root / 'lock'
        link.symlink_to(other)
        try:
            deploy.exclusive_lock(link)
        except OSError:
            pass
        else:
            raise AssertionError('symlink lock was followed')
        assert other.read_text() == 'untouched'
        link.unlink()
        link.write_text('')
        link.chmod(0o644)
        try:
            deploy.exclusive_lock(link)
        except deploy.DeployError as error:
            assert 'not private' in str(error)
        else:
            raise AssertionError('non-private lock file accepted')


def test_snapshot_headroom_refuses_before_entering_maintenance():
    with tempfile.TemporaryDirectory() as directory:
        config, commit = _maintenance_fixture(Path(directory))
        controller = SyntheticController()
        with patch.object(deploy, 'check_headroom'), \
                patch.object(deploy, 'private_recovery_directory'), \
                patch.object(deploy.shutil, 'disk_usage',
                             return_value=SimpleNamespace(free=1)):
            try:
                deploy.enter_maintenance(config, commit, controller)
            except deploy.DeployError as error:
                assert 'snapshot' in str(error)
            else:
                raise AssertionError('insufficient snapshot headroom accepted')
        assert not controller.events


def test_versioned_units_pass_preparation_syntax_check():
    deploy.verify_unit_syntax(Path(h.PROJECT_DIR))


def test_every_reviewed_unit_uses_only_current_and_one_environment_file():
    units = Path(h.PROJECT_DIR) / 'release-assets/systemd'
    for name in deploy.SERVICE_UNITS:
        content = (units / name).read_text()
        deploy.unit_paths_safe(content, name)
    web = (units / 'nuligahelper-web.service').read_text()
    for unsafe in (web.replace('/current/', '/ea0ab00/', 1),
                   web.replace('release-assets/gunicorn.conf.py',
                               'deploy/gunicorn.conf.py'),
                   web.replace('EnvironmentFile=/etc/nuligahelper/web.env',
                               'EnvironmentFile=/etc/nuligahelper/other.env')):
        try:
            deploy.unit_paths_safe(unsafe, 'nuligahelper-web.service')
        except deploy.DeployError:
            pass
        else:
            raise AssertionError('mixed or legacy installed unit paths accepted')


def test_failed_tool_output_does_not_leak_secret_canary():
    canary = 'synthetic-contact-and-secret-canary'
    try:
        deploy.run(sys.executable, '-c',
                   f'import sys; print({canary!r}, file=sys.stderr); sys.exit(7)')
    except deploy.DeployError as error:
        assert canary not in str(error)
        assert 'exit 7' in str(error)
    else:
        raise AssertionError('failed command reported success')


def test_atomic_current_switch_never_exposes_an_incomplete_candidate():
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        source, _ = synthetic_source(root)
        config = configuration(root, source)
        app_root = Path(config['app_root'])
        current = app_root / 'current'
        prior = current.resolve()
        commit = 'b' * 40
        candidate = app_root / 'releases' / commit
        candidate.mkdir(parents=True)
        try:
            deploy.switch_current(app_root, commit)
        except deploy.DeployError:
            pass
        else:
            raise AssertionError('unprepared release switched into current')
        assert current.resolve() == prior
        (candidate / '.prepared.json').write_text(json.dumps({
            'schema_version': 1, 'commit': commit, 'tree': 'a' * 40}))
        with patch.object(deploy.os, 'replace', side_effect=OSError('synthetic')):
            try:
                deploy.switch_current(app_root, commit)
            except OSError:
                pass
            else:
                raise AssertionError('failed atomic replace reported success')
        assert current.resolve() == prior
        assert not list(app_root.glob('.current-next-*'))
        assert deploy.switch_current(app_root, commit) == (str(prior), str(candidate))
        assert current.resolve() == candidate


def test_read_only_plan_contains_no_configuration_or_contact_canary():
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        source, _ = synthetic_source(root)
        config = configuration(root, source)
        commit = 'c' * 40
        candidate = Path(config['app_root']) / 'releases' / commit
        candidate.mkdir(parents=True)
        (candidate / '.prepared.json').write_text(json.dumps({
            'schema_version': 1, 'commit': commit, 'tree': 'a' * 40}))
        Path(config['database']).write_bytes(b'synthetic database marker')
        with patch.object(deploy, '_unit_state', return_value={
                'LoadState': 'loaded', 'ActiveState': 'active'}), \
                patch.object(deploy, 'schema_probe', return_value={
                    'gate': 'ready', 'revision': '0003_game_day_task_blocks'}):
            plan = deploy.inspect_plan(config, commit)
        serialized = json.dumps(plan)
        assert plan['activation_available'] is True
        assert plan['selected_commit'] == commit
        assert 'synthetic-contact-and-secret-canary' not in serialized
        assert 'public_health_url' not in serialized


def test_candidate_schema_probe_uses_its_own_runtime_without_secrets():
    with tempfile.TemporaryDirectory() as directory:
        database = Path(directory) / 'live.db'
        db.initialize_db(db.make_engine(str(database)))
        result = deploy.schema_probe(Path(h.PROJECT_DIR), database)
        assert result['gate'] == 'ready'
        assert result['revision'] == '0003_game_day_task_blocks'


def test_local_web_verification_checks_release_listener_schema_and_health():
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        config, commit = _maintenance_fixture(root)
        app_root = Path(config['app_root'])
        candidate = app_root / 'releases' / commit
        (app_root / 'current').unlink()
        (app_root / 'current').symlink_to(candidate)
        controller = SyntheticController()
        controller.loopback_listener = lambda: True
        with patch.object(deploy, 'schema_probe', return_value={'gate': 'ready'}), \
                patch.object(deploy, 'check_health') as health:
            deploy.verify_local_web(config, commit, controller)
        health.assert_called_once_with(config['public_health_url'], public=False)
        controller.loopback_listener = lambda: False
        with patch.object(deploy, 'schema_probe') as probe, \
                patch.object(deploy, 'check_health') as health:
            try:
                deploy.verify_local_web(config, commit, controller)
            except deploy.DeployError as error:
                assert 'loopback-only' in str(error)
            else:
                raise AssertionError('public-bound or absent listener accepted')
        assert not probe.called and not health.called


def test_health_checks_reject_unexpected_responses_without_redirects():
    class Response:
        status = 302

        def read(self, maximum):
            return b'ok\n'

    class Connection:
        def __init__(self, *args, **kwargs):
            self.headers = None
            self.closed = False

        def request(self, method, path, headers):
            assert method == 'GET' and path == '/healthz'
            self.headers = headers

        def getresponse(self):
            return Response()

        def close(self):
            self.closed = True

    with patch.object(deploy.http.client, 'HTTPConnection', Connection), \
            patch.object(deploy.http.client, 'HTTPSConnection', Connection):
        for public in (False, True):
            try:
                deploy.check_health('https://club.test/healthz', public=public)
            except deploy.DeployError as error:
                assert 'check failed' in str(error)
            else:
                raise AssertionError('redirecting health endpoint accepted')


def test_persistent_calendar_timer_catchup_requires_a_decision():
    def state(last):
        return {'LastTriggerUSec': last}

    after = datetime.fromisoformat('2026-09-24T10:00:00+02:00')
    before = datetime.fromisoformat('2026-09-24T08:00:00+02:00')
    assert deploy.pending_calendar_catchup(
        state('Wed 2026-09-23 09:00:08 CEST'), now=after)
    assert not deploy.pending_calendar_catchup(
        state('Thu 2026-09-24 09:00:08 CEST'), now=after)
    assert not deploy.pending_calendar_catchup(
        state('Wed 2026-09-23 09:00:08 CEST'), now=before)
    assert deploy.pending_calendar_catchup(
        state('Tue 2026-09-22 09:00:08 CEST'), now=before), \
        'a missed firing yesterday still matters before today at 09:00'
    assert deploy.pending_calendar_catchup(state('n/a'), now=after)


class SyntheticController:
    def __init__(self, *, daily_polls=0, open_handles=False):
        self.active = {name: 'active' for name in
                       (*deploy.TIMERS, 'caddy.service', 'nuligahelper-web.service')}
        self.events = []
        self.daily_polls = daily_polls
        self.open_handles = open_handles

    def state(self, unit):
        if unit == 'nuligahelper-daily.service' and self.daily_polls:
            self.daily_polls -= 1
            return {'ActiveState': 'active'}
        return {'ActiveState': self.active.get(unit, 'inactive'),
                'UnitFileState': 'enabled', 'LastTriggerUSec': 'n/a'}

    def systemctl(self, *arguments):
        self.events.append(arguments)
        self.active[arguments[-1]] = 'active' if arguments[0] == 'start' or \
            arguments[:2] == ('enable', '--now') else 'inactive'

    def loopback_listener(self):
        return True

    def database_handles(self, database):
        self.events.append(('inspect-handles', str(database)))
        return self.open_handles


def test_maintenance_waits_for_active_job_without_killing_it():
    controller = SyntheticController(daily_polls=2)
    elapsed = [0.0]
    def sleep(seconds): elapsed[0] += seconds
    result = deploy.quiesce_writers(Path('/synthetic/live.db'), controller,
                                    wait_seconds=10, clock=lambda: elapsed[0],
                                    sleep=sleep)
    assert result['writers_quiescent'] is True
    assert elapsed[0] >= 2
    assert controller.events[:3] == [
        ('disable', '--now', timer) for timer in deploy.TIMERS]
    assert ('stop', 'nuligahelper-daily.service') not in controller.events
    assert controller.events[-1] == ('inspect-handles', '/synthetic/live.db')


def test_maintenance_timeout_leaves_timers_paused_and_ingress_unchanged():
    controller = SyntheticController(daily_polls=100)
    elapsed = [0.0]
    def sleep(seconds): elapsed[0] += seconds
    try:
        deploy.quiesce_writers(Path('/synthetic/live.db'), controller,
                               wait_seconds=3, clock=lambda: elapsed[0], sleep=sleep)
    except deploy.DeployError as error:
        assert 'still active' in str(error)
    else:
        raise AssertionError('active daily job was interrupted for cutover')
    assert all(controller.active[timer] == 'inactive' for timer in deploy.TIMERS)
    assert controller.active['caddy.service'] == 'active'
    assert controller.active['nuligahelper-web.service'] == 'active'
    assert not any(event[0] == 'stop' for event in controller.events)


def test_open_database_handle_blocks_cutover_after_web_stops():
    controller = SyntheticController(open_handles=True)
    try:
        deploy.quiesce_writers(Path('/synthetic/live.db'), controller,
                               wait_seconds=1, clock=lambda: 0,
                               sleep=lambda seconds: None)
    except deploy.DeployError as error:
        assert 'open handle' in str(error)
    else:
        raise AssertionError('open database handle accepted')
    assert controller.active['caddy.service'] == 'inactive'
    assert controller.active['nuligahelper-web.service'] == 'inactive'


def test_late_started_database_job_blocks_cutover():
    class LateJobController(SyntheticController):
        def __init__(self):
            super().__init__()
            self.daily_checks = 0

        def state(self, unit):
            if unit == 'nuligahelper-daily.service':
                self.daily_checks += 1
                if self.daily_checks > 1:
                    return {'ActiveState': 'active'}
            return super().state(unit)

    controller = LateJobController()
    try:
        deploy.quiesce_writers(Path('/synthetic/live.db'), controller,
                               wait_seconds=1, clock=lambda: 0,
                               sleep=lambda seconds: None)
    except deploy.DeployError as error:
        assert 'started during maintenance' in str(error)
    else:
        raise AssertionError('late-started writer was accepted')
    assert controller.active['caddy.service'] == 'inactive'
    assert not any(event[0] == 'stop' and
                   event[-1] == 'nuligahelper-daily.service'
                   for event in controller.events)


def test_private_deployment_record_allowlists_fields_and_excludes_canaries():
    with tempfile.TemporaryDirectory() as directory:
        recovery = Path(directory)
        identifier = 'd' * 32
        record = dict(schema_version=1, deployment_id=identifier,
                      source_commit='a' * 40, source_tree='b' * 40,
                      previous_release='/opt/nuligahelper/ea0ab00',
                      previous_commit='ea0ab00' + 'c' * 33,
                      timer_before=None,
                      schema_before='0003_game_day_task_blocks', schema_after=None,
                      snapshot_path=None, checks=['fetched_master', 'offline_suite'],
                      outcome='maintenance', public_reopened=False,
                      timers_paused=True, timer_decision=None,
                      pending_catchup=[],
                      updated_at=datetime.now(timezone.utc).isoformat())
        with patch.object(deploy, 'private_recovery_directory'):
            path = deploy.write_deployment_record(recovery, record)
            assert path.stat().st_mode & 0o777 == 0o600
            assert json.loads(path.read_text()) == record
            for bad in ({**record, 'secret': 'synthetic-contact-and-secret-canary'},
                        {**record, 'previous_release':
                         '/opt/nuligahelper/synthetic-contact-and-secret-canary'},
                        {**record, 'timer_before': {
                            timer: {'ActiveState': 'active', 'UnitFileState': 'enabled',
                                    'LastTriggerUSec': 'synthetic-contact-and-secret-canary'}
                            for timer in deploy.TIMERS}},
                        {**record, 'checks': ['synthetic_contact_and_secret_canary']}):
                try:
                    deploy.write_deployment_record(recovery, bad)
                except deploy.DeployError:
                    pass
                else:
                    raise AssertionError('secret/contact canary reached a record')
        assert 'synthetic-contact-and-secret-canary' not in path.read_text()


def test_legacy_previous_commit_must_match_release_directory():
    with patch.object(deploy, 'run', return_value='ea0ab00' + 'a' * 33):
        assert deploy.prior_commit('/opt/nuligahelper/ea0ab00') == \
            'ea0ab00' + 'a' * 33
    with patch.object(deploy, 'run', return_value='f' * 40):
        try:
            deploy.prior_commit('/opt/nuligahelper/ea0ab00')
        except deploy.DeployError as error:
            assert 'does not match' in str(error)
        else:
            raise AssertionError('mismatched legacy release identity accepted')


def _maintenance_fixture(root):
    source, _ = synthetic_source(root)
    config = configuration(root, source)
    commit = 'd' * 40
    candidate = Path(config['app_root']) / 'releases' / commit
    candidate.mkdir(parents=True)
    (candidate / '.prepared.json').write_text(json.dumps({
        'schema_version': 1, 'commit': commit, 'tree': 'e' * 40}))
    Path(config['recovery_dir']).mkdir()
    Path(config['database']).write_bytes(b'synthetic database marker')
    return config, commit


def test_maintenance_records_snapshot_only_after_writers_quiesce():
    with tempfile.TemporaryDirectory() as directory:
        config, commit = _maintenance_fixture(Path(directory))
        events, records = [], []
        controller = SyntheticController()
        def snapshot(*args):
            events.append('snapshot')
            assert controller.active['caddy.service'] == 'inactive'
            assert controller.active['nuligahelper-web.service'] == 'inactive'
            assert all(controller.active[timer] == 'inactive' for timer in deploy.TIMERS)
            return Path(config['recovery_dir']) / ('snapshot-' + args[-1] + '.db')
        with patch.object(deploy, 'private_recovery_directory'), \
                patch.object(deploy, 'check_headroom'), \
                patch.object(deploy, 'prior_commit', return_value='a' * 40), \
                patch.object(deploy, 'write_deployment_record',
                             side_effect=lambda folder, record: records.append(copy.deepcopy(record))), \
                patch.object(deploy, 'schema_probe', side_effect=[
                    {'gate': 'ready', 'revision': '0003_game_day_task_blocks'},
                    {'gate': 'ready', 'revision': '0003_game_day_task_blocks'}]), \
                patch.object(deploy, 'snapshot_probe', side_effect=snapshot):
            record = deploy.enter_maintenance(config, commit, controller)
        assert events == ['snapshot']
        assert record['timers_paused'] is True
        assert record['snapshot_path'].endswith('.db')
        assert 'schema_ready' in record['checks']
        assert records[-1]['outcome'] == 'maintenance'


def test_maintenance_refuses_unsafe_schema_before_stopping_anything():
    with tempfile.TemporaryDirectory() as directory:
        config, commit = _maintenance_fixture(Path(directory))
        controller = SyntheticController()
        with patch.object(deploy, 'private_recovery_directory'), \
                patch.object(deploy, 'check_headroom'), \
                patch.object(deploy, 'prior_commit', return_value='a' * 40), \
                patch.object(deploy, 'schema_probe', return_value={
                    'gate': 'refused', 'revision': ''}), \
                patch.object(deploy, 'snapshot_probe') as snapshot:
            try:
                deploy.enter_maintenance(config, commit, controller)
            except deploy.DeployError:
                pass
            else:
                raise AssertionError('unsafe schema entered maintenance')
        assert not controller.events and not snapshot.called


def test_snapshot_failure_keeps_maintenance_and_records_failure():
    with tempfile.TemporaryDirectory() as directory:
        config, commit = _maintenance_fixture(Path(directory))
        controller = SyntheticController()
        records = []
        with patch.object(deploy, 'private_recovery_directory'), \
                patch.object(deploy, 'check_headroom'), \
                patch.object(deploy, 'prior_commit', return_value='a' * 40), \
                patch.object(deploy, 'write_deployment_record',
                             side_effect=lambda folder, record: records.append(copy.deepcopy(record))), \
                patch.object(deploy, 'schema_probe', return_value={
                    'gate': 'ready', 'revision': '0003_game_day_task_blocks'}), \
                patch.object(deploy, 'snapshot_probe',
                             side_effect=deploy.DeployError('synthetic snapshot failure')):
            try:
                deploy.enter_maintenance(config, commit, controller)
            except deploy.DeployError:
                pass
            else:
                raise AssertionError('snapshot failure allowed activation')
        assert records[-1]['outcome'] == 'failed'
        assert records[-1]['timers_paused'] is True
        assert (Path(config['app_root']) / 'current').resolve().name == 'prior'


def _ready_activation(root):
    config, commit = _maintenance_fixture(root)
    previous = str((Path(config['app_root']) / 'current').resolve())
    record = dict(schema_version=1, deployment_id='d' * 32,
                  source_commit=commit, source_tree='e' * 40,
                  previous_release=previous, previous_commit='a' * 40,
                  timer_before=None, schema_before='0003_game_day_task_blocks',
                  schema_after='0003_game_day_task_blocks',
                  snapshot_path=str(Path(config['recovery_dir']) /
                                    ('snapshot-' + 'd' * 32 + '.db')),
                  checks=['writers_quiescent', 'snapshot_validated', 'schema_ready'],
                  outcome='maintenance', public_reopened=False,
                  timers_paused=True, timer_decision=None,
                  pending_catchup=[],
                  updated_at=datetime.now(timezone.utc).isoformat())
    controller = SyntheticController()
    for unit in controller.active:
        controller.active[unit] = 'inactive'
    return config, commit, record, controller


def test_activation_switches_only_after_gates_and_keeps_timers_paused():
    with tempfile.TemporaryDirectory() as directory:
        config, commit, record, controller = _ready_activation(Path(directory))
        app_root = Path(config['app_root'])
        with patch.object(deploy, 'validate_deployment_record'), \
                patch.object(deploy, 'write_deployment_record'), \
                patch.object(deploy, 'validate_recovery_snapshot') as snapshot, \
                patch.object(deploy, 'schema_probe', return_value={
                    'gate': 'ready', 'revision': '0003_game_day_task_blocks'}), \
                patch.object(deploy, 'check_health') as health:
            result = deploy.complete_activation(config, record, controller)
        assert snapshot.called
        assert (app_root / 'current').resolve() == app_root / 'releases' / commit
        assert result['outcome'] == 'public_ready'
        assert result['public_reopened'] is True
        assert result['timers_paused'] is True
        assert ('start', 'nuligahelper-web.service') in controller.events
        assert ('start', 'caddy.service') in controller.events
        assert all(controller.active[timer] == 'inactive' for timer in deploy.TIMERS)
        assert [call.kwargs['public'] for call in health.call_args_list] == [False, True]


def test_failed_public_health_closes_ingress_and_does_not_resume_timers():
    with tempfile.TemporaryDirectory() as directory:
        config, commit, record, controller = _ready_activation(Path(directory))
        def health(url, *, public):
            if public:
                raise deploy.DeployError('synthetic public health failure')
        with patch.object(deploy, 'validate_deployment_record'), \
                patch.object(deploy, 'write_deployment_record'), \
                patch.object(deploy, 'validate_recovery_snapshot'), \
                patch.object(deploy, 'schema_probe', return_value={
                    'gate': 'ready', 'revision': '0003_game_day_task_blocks'}), \
                patch.object(deploy, 'check_health', side_effect=health):
            try:
                deploy.complete_activation(config, record, controller)
            except deploy.DeployError as error:
                assert 'public health failure' in str(error)
            else:
                raise AssertionError('failed public health accepted activation')
        assert record['outcome'] == 'failed'
        assert record['public_reopened'] is True, \
            'a possible accepted write must block automatic data restoration'
        assert controller.active['caddy.service'] == 'inactive'
        assert all(controller.active[timer] == 'inactive' for timer in deploy.TIMERS)


def test_timer_catchup_is_recorded_before_any_timer_resumes():
    with tempfile.TemporaryDirectory() as directory:
        config, commit, record, controller = _ready_activation(Path(directory))
        record['outcome'] = 'public_ready'
        record['public_reopened'] = True
        record['timer_before'] = {
            timer: {'ActiveState': 'active', 'UnitFileState': 'enabled',
                    'LastTriggerUSec': 'Wed 2026-09-23 09:00:08 CEST'}
            for timer in deploy.TIMERS}
        controller.active['caddy.service'] = 'active'
        controller.active['nuligahelper-web.service'] = 'active'
        saved = []
        with patch.object(deploy, 'validate_deployment_record'), \
                patch.object(deploy, 'write_deployment_record',
                             side_effect=lambda folder, item: saved.append(copy.deepcopy(item))), \
                patch.object(deploy, 'check_health'):
            held = deploy.resume_timers(
                config, record, controller, 'hold',
                now=datetime.fromisoformat('2026-09-24T10:00:00+02:00'))
            assert set(held['pending_catchup']) == set(deploy.CALENDAR_TIMERS)
            assert held['timer_decision'] == 'hold'
            assert not controller.events
            resumed = deploy.resume_timers(
                config, record, controller, 'run',
                now=datetime.fromisoformat('2026-09-24T10:00:00+02:00'))
        assert saved[-2]['timer_decision'] == 'run', \
            'the catch-up decision must be durable before any enable --now'
        assert resumed['outcome'] == 'accepted'
        assert resumed['timers_paused'] is False
        assert all(controller.active[timer] == 'active' for timer in deploy.TIMERS)
        assert not any('nuligahelper-daily.service' in action for action in controller.events)


def test_code_only_rollback_restores_prior_release_without_touching_database():
    with tempfile.TemporaryDirectory() as directory:
        config, commit, record, controller = _ready_activation(Path(directory))
        app_root = Path(config['app_root'])
        original_data = Path(config['database']).read_bytes()
        deploy.switch_current(app_root, commit)
        record['outcome'] = 'failed'
        with patch.object(deploy, 'validate_deployment_record'), \
                patch.object(deploy, 'write_deployment_record'), \
                patch.object(deploy, 'schema_probe', return_value={
                    'gate': 'ready', 'revision': '0003_game_day_task_blocks'}), \
                patch.object(deploy, 'previous_schema_head',
                             return_value='0003_game_day_task_blocks'), \
                patch.object(deploy, 'verify_previous_web_compatibility'), \
                patch.object(deploy, 'check_health'):
            result = deploy.rollback_code_only(config, record, controller)
        assert (app_root / 'current').resolve() == Path(record['previous_release'])
        assert Path(config['database']).read_bytes() == original_data
        assert result['outcome'] == 'rolled_back'
        assert result['timers_paused'] is True
        assert ('start', 'nuligahelper-daily.service') not in controller.events


def test_code_rollback_refuses_migrated_or_post_traffic_database():
    for changed_schema, public_reopened in ((True, False), (False, True)):
        with tempfile.TemporaryDirectory() as directory:
            config, commit, record, controller = _ready_activation(Path(directory))
            app_root = Path(config['app_root'])
            deploy.switch_current(app_root, commit)
            record['outcome'] = 'failed'
            record['public_reopened'] = public_reopened
            if changed_schema:
                record['schema_after'] = '0004_synthetic_revision'
            with patch.object(deploy, 'validate_deployment_record'):
                try:
                    deploy.rollback_code_only(config, record, controller)
                except deploy.DeployError as error:
                    assert 'rollback' in str(error)
                else:
                    raise AssertionError('unsafe code-only rollback accepted')
            assert (app_root / 'current').resolve().name == commit
            assert not controller.events


if __name__ == '__main__':
    h.run_all(globals())
