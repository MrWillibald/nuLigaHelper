"""Process-level daily-run lock tests (Linux/Raspberry Pi only)."""

from pathlib import Path
import subprocess
import sys
import tempfile

import helpers as h
from daily_lock import (
    DailyRunAlreadyActiveError,
    DailyRunLockPathError,
    daily_lock_path,
    daily_run_lock,
)


def test_same_database_overlap_fails_before_side_effects():
    with tempfile.TemporaryDirectory(prefix="daily-lock-") as directory:
        database_path = Path(directory) / "daily.db"
        side_effects: list[str] = []

        def invoke_daily_work():
            with daily_run_lock(database_path):
                side_effects.append("scrape/sync/backup/notify")

        with daily_run_lock(database_path) as held_lock:
            try:
                invoke_daily_work()
            except DailyRunAlreadyActiveError as exc:
                assert exc.database_path == database_path.resolve()
                assert exc.lock_path == held_lock.lock_path
                assert "Another daily run" in str(exc)
            else:
                raise AssertionError("overlapping work for one database must fail fast")

        assert side_effects == [], \
            "lock refusal must happen before integration side effects are invoked"


def test_leftover_unlocked_file_is_harmless():
    with tempfile.TemporaryDirectory(prefix="daily-lock-") as directory:
        database_path = Path(directory) / "daily.db"
        lock_path = daily_lock_path(database_path)
        _ = lock_path.write_text(
            "left over from an earlier process", encoding="utf-8"
        )

        with daily_run_lock(database_path) as acquired:
            assert acquired.lock_path == lock_path

        assert lock_path.exists(), "an unlocked stale file does not need deletion"


def test_different_database_paths_are_independent():
    with tempfile.TemporaryDirectory(prefix="daily-lock-") as directory:
        first_database = Path(directory) / "first.db"
        second_database = Path(directory) / "second.db"

        with (
            daily_run_lock(first_database) as first_lock,
            daily_run_lock(second_database) as second_lock,
        ):
            assert first_lock.lock_path != second_lock.lock_path


def test_canonical_database_aliases_share_one_lock():
    with tempfile.TemporaryDirectory(prefix="daily-lock-") as directory:
        root = Path(directory)
        database_path = root / "daily.db"
        database_path.touch()
        alias_path = root / "daily-alias.db"
        alias_path.symlink_to(database_path)

        with daily_run_lock(database_path):
            try:
                with daily_run_lock(alias_path):
                    raise AssertionError("a symlink alias must not acquire a separate lock")
            except DailyRunAlreadyActiveError as exc:
                assert exc.database_path == database_path.resolve()
                assert exc.lock_path == daily_lock_path(database_path)


def test_context_exit_releases_lock():
    with tempfile.TemporaryDirectory(prefix="daily-lock-") as directory:
        database_path = Path(directory) / "daily.db"

        with daily_run_lock(database_path):
            pass

        with daily_run_lock(database_path):
            pass


def test_process_exit_releases_lock_without_cleanup():
    with tempfile.TemporaryDirectory(prefix="daily-lock-") as directory:
        database_path = Path(directory) / "daily.db"
        child_code = (
            "import os, sys; "
            "from daily_lock import daily_run_lock; "
            "guard = daily_run_lock(sys.argv[1]); "
            "guard.__enter__(); "
            "os._exit(0)"
        )
        completed = subprocess.run(
            [sys.executable, "-c", child_code, str(database_path)],
            cwd=h.PROJECT_DIR,
            check=False,
            capture_output=True,
            text=True,
        )
        assert completed.returncode == 0, completed.stderr

        with daily_run_lock(database_path):
            pass


def test_missing_parent_error_is_actionable_and_does_not_create_it():
    with tempfile.TemporaryDirectory(prefix="daily-lock-") as directory:
        missing_parent = Path(directory) / "missing"
        database_path = missing_parent / "daily.db"

        try:
            with daily_run_lock(database_path):
                pass
        except DailyRunLockPathError as exc:
            message = str(exc)
            assert str(missing_parent) in message
            assert "does not exist" in message
            assert "Linux/Raspberry Pi" in message
        else:
            raise AssertionError("a missing lock parent must raise a typed path error")

        assert not missing_parent.exists()


if __name__ == "__main__":
    h.run_all(dict(globals()))
