"""Offline tests for WAL-safe backup publication and guarded restoration."""

from __future__ import annotations

import datetime as dt
import sqlite3
import tempfile
import time
from contextlib import closing
from pathlib import Path
from types import SimpleNamespace

import helpers as h  # pyright: ignore[reportImplicitRelativeImport]

import backup
import common
import requests


def _expect_error(error_type, callback):
    try:
        callback()
    except error_type as exc:
        return exc
    raise AssertionError(f"expected {error_type.__name__}")


def _create_database(path: Path, values=(1,)) -> None:
    with closing(sqlite3.connect(path)) as connection:
        connection.execute("CREATE TABLE items (value INTEGER NOT NULL)")
        connection.executemany(
            "INSERT INTO items (value) VALUES (?)",
            [(value,) for value in values],
        )
        connection.commit()


def _create_foreign_key_invalid_database(path: Path) -> None:
    with closing(sqlite3.connect(path)) as connection:
        connection.executescript(
            """
            PRAGMA foreign_keys=OFF;
            CREATE TABLE parent (id INTEGER PRIMARY KEY);
            CREATE TABLE child (
                id INTEGER PRIMARY KEY,
                parent_id INTEGER NOT NULL REFERENCES parent(id)
            );
            INSERT INTO child (id, parent_id) VALUES (1, 999);
            """
        )


def _database_values(path: Path) -> list[int]:
    with closing(sqlite3.connect(path)) as connection:
        return [row[0] for row in connection.execute("SELECT value FROM items ORDER BY value")]


def _snapshot_file(directory: Path, source: Path, name: str = "candidate.db") -> Path:
    candidate = directory / name
    candidate.write_bytes(backup.snapshot_database(source))
    return candidate


class FakeDropbox:
    def __init__(self, pages=()):
        self.pages = list(pages)
        self.uploads = []
        self.deleted = []
        self.listed_folders = []
        self.continued_cursors = []

    def files_upload(self, data, path, mode):
        self.uploads.append((data, path, mode))

    def files_list_folder(self, folder):
        self.listed_folders.append(folder)
        if not self.pages:
            return SimpleNamespace(entries=[], has_more=False, cursor="done")
        return self.pages.pop(0)

    def files_list_folder_continue(self, cursor):
        self.continued_cursors.append(cursor)
        return self.pages.pop(0)

    def files_delete_v2(self, path):
        self.deleted.append(path)


def _page(names, has_more=False, cursor="cursor"):
    return SimpleNamespace(
        entries=[SimpleNamespace(name=name) for name in names],
        has_more=has_more,
        cursor=cursor,
    )


def test_retention_validation_and_exact_naming_edges():
    assert backup.validate_dated_retention(None) == 14
    assert backup.validate_dated_retention(1) == 1
    assert backup.validate_dated_retention(99) == 99
    for invalid in (0, -1, True, False, 1.0, "14", [], {}):
        error = _expect_error(
            backup.BackupError,
            lambda invalid=invalid: backup.validate_dated_retention(invalid),
        )
        assert error.stage is backup.FailureStage.CONFIGURATION

    date = dt.date(2026, 9, 2)
    paths = backup.dropbox_backup_paths(
        "/Club Backups/sqlite/", "/var/lib/app/archive.tar.db", date
    )
    assert paths.folder == "/Club Backups/sqlite"
    assert paths.latest == "/Club Backups/sqlite/archive.tar.db"
    assert paths.dated == "/Club Backups/sqlite/archive.tar-2026-09-02.db"
    assert backup.dated_backup_name("database", date) == "database-2026-09-02"
    assert backup.dated_backup_name(".database", date) == ".database-2026-09-02"
    assert backup.dropbox_backup_paths("", "club.db", date).latest == "/club.db"
    assert (
        backup.dropbox_backup_paths("Backups", "club.db", date)
        == backup.dropbox_backup_paths("Backups", "club.db", date)
    ), "same-day reruns must resolve to exactly the same object paths"

    original_effective_today = common.effective_today
    common.effective_today = lambda: date
    try:
        assert backup.dropbox_backup_paths("Backups", "club.db").dated.endswith(
            "/club-2026-09-02.db"
        )
    finally:
        common.effective_today = original_effective_today

    for folder in ("a//b", "a/../b", "a\\b"):
        error = _expect_error(
            backup.BackupError,
            lambda folder=folder: backup.dropbox_backup_paths(folder, "club.db", date),
        )
        assert error.stage is backup.FailureStage.CONFIGURATION


def test_wal_snapshot_contains_committed_and_excludes_uncommitted_rows():
    with tempfile.TemporaryDirectory() as raw_directory:
        directory = Path(raw_directory)
        source = directory / "live.db"
        writer = sqlite3.connect(source)
        try:
            assert writer.execute("PRAGMA journal_mode=WAL").fetchone()[0] == "wal"
            writer.execute("PRAGMA wal_autocheckpoint=0")
            writer.execute("CREATE TABLE items (value INTEGER NOT NULL)")
            writer.execute("INSERT INTO items VALUES (1)")
            writer.commit()
            writer.execute("INSERT INTO items VALUES (2)")
            writer.commit()
            assert Path(f"{source}-wal").stat().st_size > 32

            writer.execute("BEGIN")
            writer.execute("INSERT INTO items VALUES (3)")
            data = backup.snapshot_database(source)
            writer.rollback()

            standalone = directory / "standalone.db"
            standalone.write_bytes(data)
            backup.validate_snapshot(standalone)
            assert _database_values(standalone) == [1, 2]
            assert not Path(f"{standalone}-wal").exists()
            assert not Path(f"{standalone}-shm").exists()
        finally:
            writer.close()


def test_snapshot_observes_only_complete_committed_writer_states():
    with tempfile.TemporaryDirectory() as raw_directory:
        directory = Path(raw_directory)
        source = directory / "live.db"
        writer = sqlite3.connect(source)
        try:
            writer.execute("PRAGMA journal_mode=WAL")
            writer.execute("PRAGMA wal_autocheckpoint=0")
            writer.execute("CREATE TABLE items (value INTEGER NOT NULL)")
            writer.executemany("INSERT INTO items VALUES (?)", [(1,), (2,)])
            writer.commit()

            writer.execute("BEGIN IMMEDIATE")
            writer.executemany("INSERT INTO items VALUES (?)", [(3,), (4,)])
            before_commit = directory / "before.db"
            before_commit.write_bytes(backup.snapshot_database(source))
            assert _database_values(before_commit) == [1, 2]

            writer.commit()
            after_commit = directory / "after.db"
            after_commit.write_bytes(backup.snapshot_database(source))
            assert _database_values(after_commit) == [1, 2, 3, 4]
        finally:
            writer.close()


def test_snapshot_validation_rejects_corruption_foreign_keys_and_sidecars():
    with tempfile.TemporaryDirectory() as raw_directory:
        directory = Path(raw_directory)
        corrupt = directory / "corrupt.db"
        corrupt.write_bytes(b"not a SQLite database")
        error = _expect_error(
            backup.SnapshotValidationError,
            lambda: backup.validate_snapshot(corrupt),
        )
        assert error.stage is backup.FailureStage.SNAPSHOT_VALIDATION
        assert error.cause is not None

        invalid_fk = directory / "invalid-fk.db"
        _create_foreign_key_invalid_database(invalid_fk)
        error = _expect_error(
            backup.SnapshotValidationError,
            lambda: backup.validate_snapshot(invalid_fk),
        )
        assert "foreign_key_check" in str(error.cause)

        valid = directory / "valid.db"
        _create_database(valid)
        Path(f"{valid}-wal").write_bytes(b"ambiguous sidecar")
        error = _expect_error(
            backup.SnapshotValidationError,
            lambda: backup.validate_snapshot(valid),
        )
        assert "self-contained" in str(error.cause)


def test_permanently_busy_snapshot_respects_a_feasible_deadline_bound():
    with tempfile.TemporaryDirectory() as raw_directory:
        source = Path(raw_directory) / "locked.db"
        _create_database(source)
        locker = sqlite3.connect(source, timeout=0.05, isolation_level=None)
        try:
            locker.execute("PRAGMA journal_mode=DELETE")
            locker.execute("BEGIN EXCLUSIVE")
            started = time.monotonic()
            error = _expect_error(
                backup.BackupError,
                lambda: backup.snapshot_database(
                    source,
                    busy_timeout_seconds=0.05,
                    deadline_seconds=0.20,
                    retry_sleep_seconds=0.01,
                ),
            )
            elapsed = time.monotonic() - started
            assert elapsed < 1.5, f"busy backup exceeded its feasible bound: {elapsed:.3f}s"
            assert error.stage in {
                backup.FailureStage.SNAPSHOT_DEADLINE,
                backup.FailureStage.SNAPSHOT_COPY,
            }
        finally:
            locker.rollback()
            locker.close()


def test_snapshot_read_and_cleanup_failures_keep_all_stages_visible():
    with tempfile.TemporaryDirectory() as raw_directory:
        directory = Path(raw_directory)
        source = directory / "live.db"
        _create_database(source)
        attempted_removals = []

        def fail_read(path):
            assert path != source, "backup must read the temporary snapshot, not the live file"
            raise OSError("injected read failure")

        def fail_cleanup(path):
            attempted_removals.append(path)
            raise PermissionError("injected cleanup failure")

        error = _expect_error(
            backup.BackupError,
            lambda: backup.snapshot_database(
                source,
                read_snapshot=fail_read,
                remove_file=fail_cleanup,
            ),
        )
        assert error.stage is backup.FailureStage.SNAPSHOT_READ
        assert [failure.stage for failure in error.secondary_failures] == [
            backup.FailureStage.LOCAL_CLEANUP
        ]
        assert attempted_removals
        for artifact in directory.glob(".live.db.snapshot-*"):
            artifact.unlink()


def test_snapshot_temporary_artifacts_are_removed_on_success_and_validation_failure():
    with tempfile.TemporaryDirectory() as raw_directory:
        directory = Path(raw_directory)
        valid = directory / "valid.db"
        _create_database(valid)
        backup.snapshot_database(valid)
        assert list(directory.glob(".valid.db.snapshot-*")) == []

        invalid = directory / "invalid.db"
        _create_foreign_key_invalid_database(invalid)
        error = _expect_error(
            backup.SnapshotValidationError,
            lambda: backup.snapshot_database(invalid),
        )
        assert error.stage is backup.FailureStage.SNAPSHOT_VALIDATION
        assert list(directory.glob(".invalid.db.snapshot-*")) == []


def test_publication_is_dated_first_latest_second_with_identical_bytes():
    date = dt.date(2026, 9, 2)
    client = FakeDropbox([_page(["club-2026-09-02.db"])])
    data = b"validated snapshot bytes"
    result = backup.publish_snapshot(
        client,
        data,
        "Backups",
        "/srv/club.db",
        retention=14,
        backup_date=date,
        upload_mode="overwrite-test",
    )
    assert [call[1] for call in client.uploads] == [
        "/Backups/club-2026-09-02.db",
        "/Backups/club.db",
    ]
    assert client.uploads[0][0] is data and client.uploads[1][0] is data
    assert [call[2] for call in client.uploads] == ["overwrite-test", "overwrite-test"]
    assert result.paths.latest == "/Backups/club.db"


def test_paginated_retention_prunes_only_anchored_oldest_database_names():
    pages = [
        _page(
            [
                "club-2026-09-01.db",
                "club.db",
                "clubhouse-2020-01-01.db",
                "club-2024-01-01.db.bak",
            ],
            has_more=True,
            cursor="page-2",
        ),
        _page(
            [
                "club-2026-09-04.db",
                "club-2026-09-03.db",
                "club-2026-99-99.db",
                "notes.txt",
            ],
            has_more=True,
            cursor="page-3",
        ),
        _page(
            [
                "club-2026-09-02.db",
                "club-copy-2020-01-01.db",
                "club-2026-9-1.db",
            ]
        ),
    ]
    client = FakeDropbox(pages)
    deleted = backup.prune_dated_backups(client, "/Backups/", "club.db", 2)
    assert client.listed_folders == ["/Backups"]
    assert client.continued_cursors == ["page-2", "page-3"]
    assert deleted == (
        "/Backups/club-2026-09-01.db",
        "/Backups/club-2026-09-02.db",
    )
    assert client.deleted == list(deleted)


def test_transient_dropbox_failures_retry_each_idempotent_operation():
    class FlakyDropbox(FakeDropbox):
        def __init__(self):
            super().__init__([
                _page(["club-2026-09-01.db"], has_more=True, cursor="next"),
                _page(["club-2026-09-02.db"]),
            ])
            self.failed_once = set()

        def _fail_once(self, operation):
            if operation not in self.failed_once:
                self.failed_once.add(operation)
                raise requests.exceptions.ConnectionError("temporary DNS failure")

        def files_upload(self, data, path, mode):
            operation = "dated" if "2026-09-02" in path else "latest"
            self._fail_once(operation)
            return super().files_upload(data, path, mode)

        def files_list_folder(self, folder):
            self._fail_once("listing")
            return super().files_list_folder(folder)

        def files_list_folder_continue(self, cursor):
            self._fail_once("pagination")
            return super().files_list_folder_continue(cursor)

        def files_delete_v2(self, path):
            self._fail_once("deletion")
            return super().files_delete_v2(path)

    sleeps = []
    client = FlakyDropbox()
    result = backup.publish_snapshot(
        client,
        b"data",
        "Backups",
        "club.db",
        retention=1,
        backup_date=dt.date(2026, 9, 2),
        retry_delays=(0.01,),
        sleep=sleeps.append,
    )

    assert client.failed_once == {
        "dated", "latest", "listing", "pagination", "deletion",
    }
    assert sleeps == [0.01] * 5
    assert [call[1] for call in client.uploads] == [
        "/Backups/club-2026-09-02.db",
        "/Backups/club.db",
    ]
    assert result.deleted_paths == ("/Backups/club-2026-09-01.db",)


def test_transient_dropbox_failure_exhaustion_preserves_stage():
    class OfflineDropbox(FakeDropbox):
        def __init__(self):
            super().__init__()
            self.attempts = 0

        def files_upload(self, data, path, mode):
            self.attempts += 1
            raise requests.exceptions.ConnectionError("DNS remains unavailable")

    sleeps = []
    client = OfflineDropbox()
    error = _expect_error(
        backup.BackupError,
        lambda: backup.publish_snapshot(
            client,
            b"data",
            "Backups",
            "club.db",
            retention=1,
            backup_date=dt.date(2026, 9, 2),
            retry_delays=(0.01, 0.02, 0.04),
            sleep=sleeps.append,
        ),
    )

    assert error.stage is backup.FailureStage.DATED_UPLOAD
    assert isinstance(error.cause, requests.exceptions.ConnectionError)
    assert client.attempts == 4
    assert sleeps == [0.01, 0.02, 0.04]


def test_publication_fault_stages_and_partial_outcomes_are_precise():
    class FaultClient(FakeDropbox):
        def __init__(self, fault):
            super().__init__([_page([], has_more=fault == "pagination")])
            self.fault = fault

        def files_upload(self, data, path, mode):
            self.uploads.append((data, path, mode))
            if self.fault == "dated" and len(self.uploads) == 1:
                raise RuntimeError("dated rejected")
            if self.fault == "latest" and len(self.uploads) == 2:
                raise RuntimeError("latest rejected")

        def files_list_folder(self, folder):
            if self.fault == "listing":
                raise RuntimeError("list rejected")
            return super().files_list_folder(folder)

        def files_list_folder_continue(self, cursor):
            if self.fault == "pagination":
                raise RuntimeError("continuation rejected")
            return super().files_list_folder_continue(cursor)

        def files_delete_v2(self, path):
            if self.fault == "delete":
                raise RuntimeError("delete rejected")
            super().files_delete_v2(path)

    cases = [
        ("dated", backup.FailureStage.DATED_UPLOAD, 1, ()),
        (
            "latest",
            backup.FailureStage.LATEST_UPLOAD,
            2,
            (backup.FailureStage.DATED_UPLOAD,),
        ),
        (
            "listing",
            backup.FailureStage.RETENTION_LIST,
            2,
            (backup.FailureStage.DATED_UPLOAD, backup.FailureStage.LATEST_UPLOAD),
        ),
        (
            "pagination",
            backup.FailureStage.RETENTION_PAGINATION,
            2,
            (
                backup.FailureStage.DATED_UPLOAD,
                backup.FailureStage.LATEST_UPLOAD,
                backup.FailureStage.RETENTION_LIST,
            ),
        ),
    ]
    for fault, expected_stage, upload_count, completed in cases:
        client = FaultClient(fault)
        error = _expect_error(
            backup.BackupError,
            lambda client=client: backup.publish_snapshot(
                client,
                b"data",
                "Backups",
                "club.db",
                retention=1,
                backup_date=dt.date(2026, 9, 2),
            ),
        )
        assert error.stage is expected_stage
        assert len(client.uploads) == upload_count
        assert error.completed_stages == completed

    delete_client = FaultClient("delete")
    delete_client.pages = [_page(["club-2026-09-01.db", "club-2026-09-02.db"])]
    error = _expect_error(
        backup.BackupError,
        lambda: backup.publish_snapshot(
            delete_client,
            b"data",
            "Backups",
            "club.db",
            retention=1,
            backup_date=dt.date(2026, 9, 2),
        ),
    )
    assert error.stage is backup.FailureStage.RETENTION_DELETE
    assert len(delete_client.uploads) == 2


def test_invalid_snapshot_and_client_construction_never_start_network_work():
    with tempfile.TemporaryDirectory() as raw_directory:
        directory = Path(raw_directory)
        invalid = directory / "invalid.db"
        _create_foreign_key_invalid_database(invalid)
        factory_calls = []

        def factory():
            factory_calls.append(True)
            return FakeDropbox()

        error = _expect_error(
            backup.BackupError,
            lambda: backup.backup_database_to_dropbox(
                invalid,
                "Backups",
                client_factory=factory,
                backup_date=dt.date(2026, 9, 2),
            ),
        )
        assert error.stage is backup.FailureStage.SNAPSHOT_VALIDATION
        assert factory_calls == []

        valid = directory / "valid.db"
        _create_database(valid)

        def failing_factory():
            raise RuntimeError("offline fake construction failure")

        error = _expect_error(
            backup.BackupError,
            lambda: backup.backup_database_to_dropbox(
                valid,
                "Backups",
                client_factory=failing_factory,
                backup_date=dt.date(2026, 9, 2),
            ),
        )
        assert error.stage is backup.FailureStage.DROPBOX_CLIENT
        assert backup.FailureStage.SNAPSHOT_READ in error.completed_stages


def test_restore_requires_confirmation_and_validates_before_any_mutation():
    with tempfile.TemporaryDirectory() as raw_directory:
        directory = Path(raw_directory)
        target = directory / "active.db"
        _create_database(target, (7,))
        wal = Path(f"{target}-wal")
        shm = Path(f"{target}-shm")
        wal.write_bytes(b"old wal")
        shm.write_bytes(b"old shm")
        original = {path: path.read_bytes() for path in (target, wal, shm)}

        valid_source = directory / "source.db"
        _create_database(valid_source, (1, 2))
        candidate = _snapshot_file(directory, valid_source)
        error = _expect_error(
            backup.RestoreError,
            lambda: backup.install_snapshot(
                candidate, target, confirmed_stopped=False
            ),
        )
        assert error.stage is backup.FailureStage.RESTORE_CONFIRMATION
        assert {path: path.read_bytes() for path in original} == original

        invalid = directory / "invalid.db"
        invalid.write_bytes(b"broken")
        error = _expect_error(
            backup.RestoreError,
            lambda: backup.install_snapshot(
                invalid, target, confirmed_stopped=True
            ),
        )
        assert error.stage is backup.FailureStage.RESTORE_VALIDATION
        assert {path: path.read_bytes() for path in original} == original
        assert list(directory.glob(".active.db.rollback-*")) == []
        assert list(directory.glob(".active.db.restore-*")) == []


def test_restore_preserves_active_set_and_installs_standalone_snapshot():
    with tempfile.TemporaryDirectory() as raw_directory:
        directory = Path(raw_directory)
        wal_source = directory / "wal-source.db"
        writer = sqlite3.connect(wal_source)
        try:
            writer.execute("PRAGMA journal_mode=WAL")
            writer.execute("PRAGMA wal_autocheckpoint=0")
            writer.execute("CREATE TABLE items (value INTEGER NOT NULL)")
            writer.executemany("INSERT INTO items VALUES (?)", [(10,), (20,)])
            writer.commit()
            assert Path(f"{wal_source}-wal").exists()
            candidate = _snapshot_file(directory, wal_source)
        finally:
            writer.close()

        target = directory / "active.db"
        _create_database(target, (99,))
        wal = Path(f"{target}-wal")
        shm = Path(f"{target}-shm")
        wal.write_bytes(b"preserved wal")
        shm.write_bytes(b"preserved shm")
        old_bytes = {path.name: path.read_bytes() for path in (target, wal, shm)}

        result = backup.install_snapshot(candidate, target, confirmed_stopped=True)
        backup.validate_snapshot(target)
        assert _database_values(target) == [10, 20]
        assert not wal.exists() and not shm.exists()
        assert result.rollback_directory.is_dir()
        assert {path.name for path in result.preserved_paths} == set(old_bytes)
        for name, data in old_bytes.items():
            assert (result.rollback_directory / name).read_bytes() == data


def test_restore_install_failure_recovers_database_and_sidecars():
    with tempfile.TemporaryDirectory() as raw_directory:
        directory = Path(raw_directory)
        source = directory / "source.db"
        _create_database(source, (1, 2))
        candidate = _snapshot_file(directory, source)

        target = directory / "active.db"
        _create_database(target, (77,))
        wal = Path(f"{target}-wal")
        shm = Path(f"{target}-shm")
        wal.write_bytes(b"rollback wal")
        shm.write_bytes(b"rollback shm")
        original = {path: path.read_bytes() for path in (target, wal, shm)}

        def fail_install(_source, _target):
            raise OSError("injected atomic replace failure")

        error = _expect_error(
            backup.RestoreError,
            lambda: backup.install_snapshot(
                candidate,
                target,
                confirmed_stopped=True,
                install_replace=fail_install,
            ),
        )
        assert error.stage is backup.FailureStage.RESTORE_INSTALL
        assert error.secondary_failures == ()
        assert {path: path.read_bytes() for path in original} == original
        assert list(directory.glob(".active.db.restore-*")) == []
        assert list(directory.glob(".active.db.rollback-*")) == []


if __name__ == "__main__":
    h.run_all(dict(globals()))
