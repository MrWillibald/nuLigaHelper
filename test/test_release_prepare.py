"""Offline checks for operator-selected, non-activating release preparation."""

import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
from unittest.mock import patch

import helpers as h


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
                service_group='synthetic', python=sys.executable)


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


def test_versioned_units_pass_preparation_syntax_check():
    deploy.verify_unit_syntax(Path(h.PROJECT_DIR))


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


if __name__ == '__main__':
    h.run_all(globals())
