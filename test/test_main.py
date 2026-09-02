"""Offline tests for daily-job orchestration, locking, and failure handling."""

import contextlib
import datetime
import io
import logging
import os
import subprocess
import sys
import tempfile
from typing import Any
from unittest.mock import patch

import helpers as h
import backup
import daily_lock
import db
import main


TODAY = datetime.date(2026, 9, 2)


def _config(
    database_path: str,
    *,
    retention: Any = None,
    include_retention: bool = False,
) -> dict[str, Any]:
    dropbox_config: dict[str, Any] = {
        "dropbox_token": "synthetic-token",
        "dropbox_folder": "synthetic-backups",
    }
    if include_retention:
        dropbox_config["dated_retention"] = retention
    return {
        "club": {
            "database": {"path": database_path},
            "dropbox": dropbox_config,
            "info": {"url": "offline-test"},
        }
    }


def _events():
    return db.SyncEvents(
        shifts=[
            db.ShiftEvent(
                game_id=1,
                game_nr=1001,
                old_date="01.09.2026",
                old_time="10:00",
                new_date="02.09.2026",
                new_time="11:00",
            )
        ],
        referee_alerts=[
            db.RefereeEvent(game_id=1, game_nr=1001, date="02.09.2026", time="11:00")
        ],
        new_games=[db.GameEvent(1, 1001, "test:1001", "BL mD")],
    )


def _expect_error(error_type, callback):
    try:
        callback()
    except error_type as error:
        return error
    raise AssertionError(f"expected {error_type.__name__}")


@contextlib.contextmanager
def _captured_logs():
    stream = io.StringIO()
    handler = logging.StreamHandler(stream)
    root = logging.getLogger()
    old_level = root.level
    root.addHandler(handler)
    root.setLevel(logging.DEBUG)
    try:
        yield stream
    finally:
        root.removeHandler(handler)
        root.setLevel(old_level)


class _FakeSession:
    def __init__(self, state, name):
        self.state = state
        self.name = name
        self.write_transaction = False
        self.closed = False

    def __enter__(self):
        self.state["order"].append(f"{self.name}:enter")
        return self

    def commit(self):
        self.state["order"].append(f"{self.name}:commit")
        self.write_transaction = False

    def __exit__(self, _error_type, _error, _traceback):
        self.closed = True
        self.write_transaction = False
        self.state["order"].append(f"{self.name}:exit")


class _FakeNotifier:
    def __init__(self, state, session, notification_error):
        self.state = state
        self.session = session
        self.notification_error = notification_error
        self.state["order"].append("notifier:init")

    def _call(self, name):
        assert self.state["lock_active"], f"{name} ran without the daily lock"
        assert not any(
            session.write_transaction for session in self.state["sessions"]
        ), f"{name} ran during a write transaction"
        self.state["order"].append(name)
        self.state["notification_calls"].append(name)
        if name == "notify_shifts" and self.notification_error is not None:
            raise self.notification_error
        return 1

    def notify_shifts(self, _events):
        return self._call("notify_shifts")

    def notify_referee_alert(self, _event):
        return self._call("notify_referee_alert")

    def notify_new_games(self, _events):
        return self._call("notify_new_games")

    def notify_game_day(self, _date):
        return self._call("notify_game_day")

    def notify_referees_for_date(self, _date):
        return self._call("notify_referees_for_date")

    def notify_service_early(self, _date):
        return self._call("notify_service_early")

    def notify_pre(self, _date):
        return self._call("notify_pre")


@contextlib.contextmanager
def _fake_job(
    database_path,
    *,
    backup_error=None,
    notification_error=None,
    games_exist=True,
    include_retention=False,
    retention=None,
    real_lock=False,
):
    state: dict[str, Any] = {
        "order": [],
        "sessions": [],
        "lock_active": False,
        "lock_paths": [],
        "backup_calls": [],
        "notification_calls": [],
    }
    events = _events()
    real_daily_run_lock = daily_lock.daily_run_lock

    @contextlib.contextmanager
    def fake_lock(path):
        state["lock_paths"].append(os.fspath(path))
        state["lock_active"] = True
        state["order"].append("lock:enter")
        try:
            yield
        finally:
            state["order"].append("lock:exit")
            state["lock_active"] = False

    @contextlib.contextmanager
    def tracked_real_lock(path):
        state["lock_paths"].append(os.fspath(path))
        with real_daily_run_lock(path):
            state["lock_active"] = True
            state["order"].append("lock:enter")
            try:
                yield
            finally:
                state["order"].append("lock:exit")
                state["lock_active"] = False

    def make_engine(path):
        assert state["lock_active"], "engine construction happened before the lock"
        state["order"].append("make_engine")
        state["engine_path"] = path
        return "synthetic-engine"

    def init_db(_engine):
        assert state["lock_active"], "database initialization happened before the lock"
        state["order"].append("init_db")

    def scrape(_info, _season_year):
        assert state["lock_active"], "scraping happened before the lock"
        assert not state["sessions"], "scraping happened with a database session open"
        state["order"].append("scrape")
        return [{"source_key": "offline"}]

    def session_factory(_engine):
        session = _FakeSession(state, f"session{len(state['sessions'])}")
        state["sessions"].append(session)
        return session

    def sync_games(session, _scraped, _season_year):
        assert state["lock_active"], "sync happened before the lock"
        session.write_transaction = True
        state["order"].append("sync")
        return events

    def run_backup(path, folder, **options):
        assert state["lock_active"], "backup happened without the daily lock"
        assert state["sessions"][0].closed, "sync session was not closed before backup"
        assert not any(
            session.write_transaction for session in state["sessions"]
        ), "backup happened during a write transaction"
        state["order"].append("backup")
        state["backup_calls"].append((path, folder, options))
        if backup_error is not None:
            raise backup_error
        paths = backup.DropboxBackupPaths(
            folder="/synthetic-backups",
            latest="/synthetic-backups/daily.db",
            dated="/synthetic-backups/daily-2026-09-02.db",
        )
        return backup.BackupResult(paths=paths, deleted_paths=(), byte_count=123)

    def notifier_factory(_club_cfg, session, _season_year):
        assert state["sessions"][0].closed
        assert session is state["sessions"][1]
        return _FakeNotifier(state, session, notification_error)

    def get_games_on_date(session, date):
        assert state["lock_active"], "notification query happened without the lock"
        assert session is state["sessions"][1]
        assert not any(item.write_transaction for item in state["sessions"])
        state["order"].append(f"games:{date}")
        return [object()] if games_exist else []

    lock_factory = tracked_real_lock if real_lock else fake_lock
    patches: list[Any] = [
        patch.object(
            main.common,
            "load_config",
            return_value=_config(
                database_path,
                retention=retention,
                include_retention=include_retention,
            ),
        ),
        patch.object(main.common, "effective_today", return_value=TODAY),
        patch.object(main.db, "make_engine", side_effect=make_engine),
        patch.object(main.db, "init_db", side_effect=init_db),
        patch.object(main, "fetch_home_games", side_effect=scrape),
        patch.object(main.db, "Session", side_effect=session_factory),
        patch.object(main.db, "sync_games", side_effect=sync_games),
        patch.object(
            main.backup,
            "backup_database_to_dropbox",
            side_effect=run_backup,
        ),
        patch.object(main, "Notifier", side_effect=notifier_factory),
        patch.object(main.db, "get_games_on_date", side_effect=get_games_on_date),
        patch.object(main.daily_lock, "daily_run_lock", lock_factory),
    ]

    with contextlib.ExitStack() as stack:
        for active_patch in patches:
            stack.enter_context(active_patch)
        yield state


def test_new_game_filter_keeps_duplicate_numbers_independent():
    events = db.SyncEvents(new_games=[
        db.GameEvent(1, 555, "meeting:101", "GE"),
        db.GameEvent(2, 555, "meeting:102", "BL mD"),
    ])
    result = main.reportable_new_games(events)
    assert [event.game_id for event in result] == [2]


def test_successful_daily_path_orders_lock_sync_backup_and_notifications():
    with tempfile.TemporaryDirectory() as directory:
        database_path = os.path.join(directory, "daily.db")
        with _fake_job(database_path, games_exist=True) as state:
            assert main.main() is None

    assert state["order"] == [
        "lock:enter",
        "make_engine",
        "init_db",
        "scrape",
        "session0:enter",
        "sync",
        "session0:commit",
        "session0:exit",
        "backup",
        "session1:enter",
        "notifier:init",
        "notify_shifts",
        "notify_referee_alert",
        "notify_new_games",
        "games:03.09.2026",
        "notify_game_day",
        "notify_referees_for_date",
        "games:09.09.2026",
        "notify_service_early",
        "notify_pre",
        "session1:exit",
        "lock:exit",
    ]
    assert state["engine_path"] == os.path.realpath(database_path)
    assert state["lock_paths"] == [os.path.realpath(database_path)]
    assert len(state["backup_calls"]) == 1
    backup_path, folder, options = state["backup_calls"][0]
    assert backup_path == os.path.realpath(database_path)
    assert folder == "synthetic-backups"
    assert options["backup_date"] == TODAY
    assert "retention" not in options, "omitted retention must use backup.py's default"
    assert callable(options["client_factory"])
    assert all(session.closed for session in state["sessions"])


def test_explicit_retention_is_forwarded_to_backup_api():
    with tempfile.TemporaryDirectory() as directory:
        database_path = os.path.join(directory, "retained.db")
        with _fake_job(
            database_path,
            games_exist=False,
            include_retention=True,
            retention=5,
        ) as state:
            main.main()
    assert state["backup_calls"][0][2]["retention"] == 5


def test_overlap_fails_before_database_or_external_side_effects():
    with tempfile.TemporaryDirectory() as directory:
        database_path = os.path.join(directory, "overlap.db")
        with daily_lock.daily_run_lock(database_path):
            with _fake_job(database_path, real_lock=True) as state:
                error = _expect_error(
                    daily_lock.DailyRunAlreadyActiveError,
                    main.main,
                )

    assert "Refusing to overlap" in str(error)
    assert state["order"] == []
    assert state["backup_calls"] == []
    assert state["notification_calls"] == []
    assert state["sessions"] == []


def test_leftover_unlocked_file_runs_and_lock_is_released_afterward():
    with tempfile.TemporaryDirectory() as directory:
        database_path = os.path.join(directory, "stale.db")
        lock_path = daily_lock.daily_lock_path(database_path)
        lock_path.write_text("left behind by an old process", encoding="utf-8")

        with _fake_job(database_path, games_exist=False, real_lock=True) as state:
            main.main()

        with daily_lock.daily_run_lock(database_path) as acquired:
            assert acquired.lock_path == lock_path

    assert "backup" in state["order"]
    assert "notify_shifts" in state["notification_calls"]


def test_lock_for_a_different_database_is_independent():
    with tempfile.TemporaryDirectory() as directory:
        first_path = os.path.join(directory, "first.db")
        second_path = os.path.join(directory, "second.db")
        with daily_lock.daily_run_lock(first_path):
            with _fake_job(second_path, games_exist=False, real_lock=True) as state:
                main.main()

    assert state["engine_path"] == os.path.realpath(second_path)
    assert "backup" in state["order"]


def test_backup_failure_is_retained_while_notifications_continue():
    cleanup_error = OSError("synthetic cleanup failure")
    backup_error = backup.BackupError(
        backup.FailureStage.LATEST_UPLOAD,
        "synthetic latest upload failure",
        cause=OSError("Dropbox unavailable"),
        completed_stages=(backup.FailureStage.DATED_UPLOAD,),
        secondary_failures=(
            backup.StageFailure(
                backup.FailureStage.LOCAL_CLEANUP,
                "synthetic leftover",
                cleanup_error,
            ),
        ),
    )
    with tempfile.TemporaryDirectory() as directory:
        database_path = os.path.join(directory, "backup-failure.db")
        with _captured_logs() as logs:
            with _fake_job(database_path, backup_error=backup_error) as state:
                error = _expect_error(main.DailyJobError, main.main)

    assert [stage for stage, _failure in error.failures] == ["backup:latest_upload"]
    assert state["notification_calls"] == [
        "notify_shifts",
        "notify_referee_alert",
        "notify_new_games",
        "notify_game_day",
        "notify_referees_for_date",
        "notify_service_early",
        "notify_pre",
    ]
    output = logs.getvalue()
    assert "stage latest_upload" in output
    assert "completed: dated_upload" in output
    assert "Secondary backup failure: local_cleanup" in output
    assert "Traceback" not in output


def test_combined_backup_and_notification_failures_are_both_visible_without_retry():
    backup_error = backup.BackupError(
        backup.FailureStage.RETENTION_PAGINATION,
        "synthetic pagination failure",
        completed_stages=(
            backup.FailureStage.DATED_UPLOAD,
            backup.FailureStage.LATEST_UPLOAD,
        ),
    )
    notification_error = RuntimeError("synthetic notification failure")
    with tempfile.TemporaryDirectory() as directory:
        database_path = os.path.join(directory, "combined-failure.db")
        with _captured_logs() as logs:
            with _fake_job(
                database_path,
                backup_error=backup_error,
                notification_error=notification_error,
            ) as state:
                error = _expect_error(main.DailyJobError, main.main)

    assert [stage for stage, _failure in error.failures] == [
        "backup:retention_pagination",
        "notifications",
    ]
    assert len(state["backup_calls"]) == 1
    assert state["notification_calls"] == ["notify_shifts"], (
        "a fatal notification call must not be retried or replay the sequence"
    )
    output = logs.getvalue()
    assert "retention_pagination" in output
    assert "dated_upload, latest_upload" in output
    assert "Notification delivery failed: synthetic notification failure" in output
    assert "2 fatal stage(s)" in output


def test_overlap_is_shell_visible_and_nonzero():
    with tempfile.TemporaryDirectory() as directory:
        database_path = os.path.join(directory, "subprocess-overlap.db")
        child_code = "\n".join([
            "from unittest.mock import patch",
            "import common",
            "import runpy",
            f"config = {repr(_config(database_path))}",
            "with patch.object(common, 'load_config', return_value=config):",
            "    runpy.run_path('main.py', run_name='__main__')",
        ])
        environment = os.environ.copy()
        environment["NULIGAHELPER_SECRET"] = "synthetic-subprocess-secret"
        with daily_lock.daily_run_lock(database_path):
            result = subprocess.run(
                [sys.executable, "-c", child_code],
                cwd=h.PROJECT_DIR,
                env=environment,
                text=True,
                capture_output=True,
                check=False,
            )

    assert result.returncode != 0
    assert "Another daily run already owns the lock" in result.stderr
    assert "Refusing to overlap" in result.stderr
    assert "Traceback" not in result.stderr


def test_run_sh_propagates_daily_job_failure_status():
    with tempfile.TemporaryDirectory() as directory:
        activate_directory = os.path.join(directory, "venv", "bin")
        fake_bin = os.path.join(directory, "fake-bin")
        os.makedirs(activate_directory)
        os.makedirs(fake_bin)

        activate_path = os.path.join(activate_directory, "activate")
        with open(activate_path, "w", encoding="utf-8") as target:
            target.write(f'export PATH="{fake_bin}:$PATH"\n')

        fake_python = os.path.join(fake_bin, "python")
        with open(fake_python, "w", encoding="utf-8") as target:
            target.write(
                "#!/bin/sh\n"
                "echo 'Refusing to overlap: synthetic daily lock owner' >&2\n"
                "exit 73\n"
            )
        os.chmod(fake_python, 0o755)

        environment = os.environ.copy()
        environment["NULIGAHELPER_SECRET"] = "synthetic-launcher-secret"
        result = subprocess.run(
            ["bash", os.path.join(h.PROJECT_DIR, "run.sh")],
            cwd=directory,
            env=environment,
            text=True,
            capture_output=True,
            check=False,
        )

    assert result.returncode == 73
    assert "Refusing to overlap" in result.stderr


if __name__ == "__main__":
    h.run_all(dict(globals()))
