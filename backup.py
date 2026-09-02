"""Safe online SQLite backup publication and offline restore helpers.

The module deliberately has no knowledge of application configuration or Dropbox
credentials.  Callers inject an authenticated Dropbox-compatible client factory,
which also keeps the helpers fully testable offline.
"""

from __future__ import annotations

import datetime as dt
import logging
import math
import os
import re
import shutil
import sqlite3
import tempfile
import time
from collections.abc import Callable, Iterable
from contextlib import closing
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Any

import requests

import common

try:
    from dropbox import exceptions as _dropbox_exceptions
    from dropbox import files as _dropbox_files
except ImportError:  # pragma: no cover - production requirements include Dropbox
    _dropbox_exceptions = None
    _dropbox_files = None

DROPBOX_OVERWRITE_MODE: Any = (
    _dropbox_files.WriteMode.overwrite if _dropbox_files is not None else "overwrite"
)


DEFAULT_DATED_RETENTION = 14
SQLITE_BUSY_TIMEOUT_SECONDS = 5.0
DEFAULT_BACKUP_DEADLINE_SECONDS = 30.0
DEFAULT_DROPBOX_RETRY_DELAYS = (1.0, 2.0, 4.0)


class FailureStage(str, Enum):
    """Machine-readable stages for backup and restore failures."""

    CONFIGURATION = "configuration"
    SNAPSHOT_CREATE = "snapshot_create"
    SNAPSHOT_COPY = "snapshot_copy"
    SNAPSHOT_DEADLINE = "snapshot_deadline"
    SNAPSHOT_VALIDATION = "snapshot_validation"
    SNAPSHOT_READ = "snapshot_read"
    LOCAL_CLEANUP = "local_cleanup"
    DROPBOX_CLIENT = "dropbox_client"
    DATED_UPLOAD = "dated_upload"
    LATEST_UPLOAD = "latest_upload"
    RETENTION_LIST = "retention_list"
    RETENTION_PAGINATION = "retention_pagination"
    RETENTION_DELETE = "retention_delete"
    RESTORE_CONFIRMATION = "restore_confirmation"
    RESTORE_VALIDATION = "restore_validation"
    RESTORE_COPY = "restore_copy"
    RESTORE_PRESERVE = "restore_preserve"
    RESTORE_INSTALL = "restore_install"
    RESTORE_RECOVERY = "restore_recovery"
    RESTORE_CLEANUP = "restore_cleanup"


@dataclass(frozen=True)
class StageFailure:
    """One failure, including secondary cleanup or recovery failures."""

    stage: FailureStage
    message: str
    exception: BaseException | None = None

    def __str__(self) -> str:
        if self.exception is None:
            return f"{self.stage.value}: {self.message}"
        return f"{self.stage.value}: {self.message}: {self.exception}"


class StagedOperationError(RuntimeError):
    """Base error that preserves a primary stage and all secondary failures."""

    def __init__(
        self,
        stage: FailureStage,
        message: str,
        *,
        cause: BaseException | None = None,
        secondary_failures: Iterable[StageFailure] = (),
        completed_stages: Iterable[FailureStage] = (),
    ) -> None:
        self.failure = StageFailure(stage, message, cause)
        self.secondary_failures = tuple(secondary_failures)
        self.completed_stages = tuple(completed_stages)
        super().__init__(str(self.failure))

    @property
    def stage(self) -> FailureStage:
        return self.failure.stage

    @property
    def cause(self) -> BaseException | None:
        return self.failure.exception

    @property
    def failures(self) -> tuple[StageFailure, ...]:
        return (self.failure, *self.secondary_failures)

    def add_secondary(self, failures: Iterable[StageFailure]) -> None:
        self.secondary_failures = (*self.secondary_failures, *failures)

    def add_completed_prefix(self, stages: Iterable[FailureStage]) -> None:
        prefix = tuple(stages)
        self.completed_stages = (*prefix, *self.completed_stages)


class BackupError(StagedOperationError):
    """A snapshot, publication, retention, or local-cleanup failure."""


class SnapshotValidationError(BackupError):
    """A candidate is not a self-contained, valid SQLite snapshot."""


class RestoreError(StagedOperationError):
    """A guarded restore validation, installation, or recovery failure."""


@dataclass(frozen=True)
class DropboxBackupPaths:
    folder: str
    latest: str
    dated: str


@dataclass(frozen=True)
class PublicationResult:
    paths: DropboxBackupPaths
    deleted_paths: tuple[str, ...]


@dataclass(frozen=True)
class BackupResult:
    paths: DropboxBackupPaths
    deleted_paths: tuple[str, ...]
    byte_count: int


@dataclass(frozen=True)
class RestoreResult:
    restored_path: Path
    rollback_directory: Path
    preserved_paths: tuple[Path, ...]


class _SnapshotDeadlineExpired(TimeoutError):
    pass


def validate_dated_retention(value: object | None) -> int:
    """Return a positive dated-retention count, defaulting an absent value to 14."""
    if value is None:
        return DEFAULT_DATED_RETENTION
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise BackupError(
            FailureStage.CONFIGURATION,
            "dated_retention must be a positive integer",
        )
    return value


def _database_basename(database_path: os.PathLike[str] | str) -> str:
    name = Path(database_path).name
    if not name or name in {".", ".."} or "\x00" in name:
        raise BackupError(
            FailureStage.CONFIGURATION,
            "database path must have a valid basename",
        )
    return name


def _normalize_dropbox_folder(folder: str) -> str:
    if not isinstance(folder, str) or "\x00" in folder or "\\" in folder:
        raise BackupError(
            FailureStage.CONFIGURATION,
            "Dropbox folder must be a slash-separated string",
        )
    stripped = folder.strip("/")
    if not stripped:
        return ""
    parts = stripped.split("/")
    if any(part in {"", ".", ".."} for part in parts):
        raise BackupError(
            FailureStage.CONFIGURATION,
            "Dropbox folder contains an invalid path component",
        )
    return "/" + "/".join(parts)


def _validate_backup_date(value: dt.date) -> dt.date:
    if type(value) is not dt.date:
        raise BackupError(
            FailureStage.CONFIGURATION,
            "backup date must be a datetime.date",
        )
    return value


def dated_backup_name(
    database_path: os.PathLike[str] | str,
    backup_date: dt.date,
) -> str:
    """Return ``<stem>-YYYY-MM-DD<suffix>`` for a database basename."""
    name = _database_basename(database_path)
    date_value = _validate_backup_date(backup_date)
    suffix = Path(name).suffix
    stem = name[: -len(suffix)] if suffix else name
    return f"{stem}-{date_value.isoformat()}{suffix}"


def dropbox_backup_paths(
    folder: str,
    database_path: os.PathLike[str] | str,
    backup_date: dt.date | None = None,
) -> DropboxBackupPaths:
    """Build exact latest and dated Dropbox paths for this database."""
    folder_path = _normalize_dropbox_folder(folder)
    date_value = common.effective_today() if backup_date is None else backup_date
    date_value = _validate_backup_date(date_value)
    basename = _database_basename(database_path)
    dated_name = dated_backup_name(basename, date_value)
    prefix = folder_path or ""
    return DropboxBackupPaths(
        folder=folder_path,
        latest=f"{prefix}/{basename}",
        dated=f"{prefix}/{dated_name}",
    )


def _readonly_uri(path: Path, *, immutable: bool = False) -> str:
    query = "mode=ro&immutable=1" if immutable else "mode=ro"
    return f"{path.resolve().as_uri()}?{query}"


def validate_snapshot(
    snapshot_path: os.PathLike[str] | str,
    *,
    connect: Callable[..., sqlite3.Connection] = sqlite3.connect,
) -> None:
    """Require an independent SQLite file with clean integrity and FK checks."""
    path = Path(snapshot_path)
    try:
        if not path.is_file():
            raise FileNotFoundError(path)
        sidecars = [Path(f"{path}-wal"), Path(f"{path}-shm")]
        if any(sidecar.exists() for sidecar in sidecars):
            raise ValueError("snapshot must be self-contained and have no WAL sidecars")

        with closing(
            connect(
                _readonly_uri(path, immutable=True),
                uri=True,
                timeout=SQLITE_BUSY_TIMEOUT_SECONDS,
            )
        ) as connection:
            quick_check = connection.execute("PRAGMA quick_check").fetchall()
            if quick_check != [("ok",)]:
                raise ValueError(f"quick_check returned {quick_check!r}")
            foreign_key_rows = connection.execute("PRAGMA foreign_key_check").fetchall()
            if foreign_key_rows:
                preview = foreign_key_rows[:10]
                raise ValueError(f"foreign_key_check returned {preview!r}")
    except SnapshotValidationError:
        raise
    except BaseException as exc:
        raise SnapshotValidationError(
            FailureStage.SNAPSHOT_VALIDATION,
            f"snapshot validation failed for {path.name}",
            cause=exc,
        ) from exc


def _positive_seconds(value: object, name: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise BackupError(FailureStage.CONFIGURATION, f"{name} must be positive")
    converted = float(value)
    if not math.isfinite(converted) or converted <= 0:
        raise BackupError(FailureStage.CONFIGURATION, f"{name} must be positive")
    return converted


def _snapshot_artifacts(path: Path) -> tuple[Path, ...]:
    return (
        path,
        Path(f"{path}-wal"),
        Path(f"{path}-shm"),
        Path(f"{path}-journal"),
    )


def _cleanup_paths(
    paths: Iterable[Path],
    remove_file: Callable[[Path], None],
    stage: FailureStage,
) -> list[StageFailure]:
    failures: list[StageFailure] = []
    for path in paths:
        try:
            if path.exists() or path.is_symlink():
                remove_file(path)
        except FileNotFoundError:
            continue
        except BaseException as exc:
            failures.append(StageFailure(stage, f"could not remove {path.name}", exc))
    return failures


def snapshot_database(
    database_path: os.PathLike[str] | str,
    *,
    busy_timeout_seconds: float = SQLITE_BUSY_TIMEOUT_SECONDS,
    deadline_seconds: float = DEFAULT_BACKUP_DEADLINE_SECONDS,
    pages_per_step: int = 128,
    retry_sleep_seconds: float = 0.05,
    connect: Callable[..., sqlite3.Connection] = sqlite3.connect,
    monotonic: Callable[[], float] = time.monotonic,
    remove_file: Callable[[Path], None] | None = None,
    read_snapshot: Callable[[Path], bytes] | None = None,
) -> bytes:
    """Create, validate, read, and clean a unique online SQLite snapshot.

    ``Connection.backup`` provides the committed WAL-aware view.  The source and
    destination use a five-second busy policy by default, while the progress
    callback enforces a separate overall deadline (30 seconds by default).
    """
    source_path = Path(database_path)
    busy_timeout = _positive_seconds(busy_timeout_seconds, "busy timeout")
    deadline = _positive_seconds(deadline_seconds, "backup deadline")
    if isinstance(pages_per_step, bool) or not isinstance(pages_per_step, int) or pages_per_step <= 0:
        raise BackupError(FailureStage.CONFIGURATION, "pages_per_step must be positive")
    sleep_seconds = _positive_seconds(retry_sleep_seconds, "retry sleep")

    def default_remove(path: Path) -> None:
        path.unlink()

    def default_read(path: Path) -> bytes:
        return path.read_bytes()

    remover = remove_file or default_remove
    reader = read_snapshot or default_read

    snapshot_path: Path | None = None
    descriptor: int | None = None
    primary_error: BackupError | None = None
    snapshot_bytes: bytes | None = None
    current_stage = FailureStage.SNAPSHOT_CREATE

    try:
        if not source_path.is_file():
            raise FileNotFoundError(source_path)
        descriptor, raw_path = tempfile.mkstemp(
            prefix=f".{source_path.name}.snapshot-",
            suffix=".db",
            dir=source_path.parent,
        )
        snapshot_path = Path(raw_path)
        os.close(descriptor)
        descriptor = None

        current_stage = FailureStage.SNAPSHOT_COPY
        started = monotonic()
        deadline_at = started + deadline

        def check_deadline() -> None:
            if monotonic() >= deadline_at:
                raise _SnapshotDeadlineExpired(
                    f"online backup exceeded {deadline:g} seconds"
                )

        def progress(_status: int, _remaining: int, _total: int) -> None:
            check_deadline()

        with closing(
            connect(
                _readonly_uri(source_path),
                uri=True,
                timeout=busy_timeout,
            )
        ) as source, closing(connect(str(snapshot_path), timeout=busy_timeout)) as target:
            busy_ms = max(1, round(busy_timeout * 1000))
            source.execute(f"PRAGMA busy_timeout={busy_ms}")
            target.execute(f"PRAGMA busy_timeout={busy_ms}")
            check_deadline()
            source.backup(
                target,
                pages=pages_per_step,
                progress=progress,
                sleep=sleep_seconds,
            )
            check_deadline()

        current_stage = FailureStage.SNAPSHOT_VALIDATION
        validate_snapshot(snapshot_path, connect=connect)

        current_stage = FailureStage.SNAPSHOT_READ
        snapshot_bytes = reader(snapshot_path)
    except SnapshotValidationError as exc:
        primary_error = exc
    except _SnapshotDeadlineExpired as exc:
        primary_error = BackupError(
            FailureStage.SNAPSHOT_DEADLINE,
            "online SQLite snapshot did not finish before its deadline",
            cause=exc,
        )
    except BackupError as exc:
        primary_error = exc
    except BaseException as exc:
        primary_error = BackupError(
            current_stage,
            f"backup failed while processing {source_path.name}",
            cause=exc,
        )
    finally:
        cleanup_failures: list[StageFailure] = []
        if descriptor is not None:
            try:
                os.close(descriptor)
            except BaseException as exc:
                cleanup_failures.append(
                    StageFailure(
                        FailureStage.LOCAL_CLEANUP,
                        "could not close temporary snapshot descriptor",
                        exc,
                    )
                )
        if snapshot_path is not None:
            cleanup_failures.extend(
                _cleanup_paths(
                    _snapshot_artifacts(snapshot_path),
                    remover,
                    FailureStage.LOCAL_CLEANUP,
                )
            )

        if primary_error is not None:
            primary_error.add_secondary(cleanup_failures)
        elif cleanup_failures:
            first, *rest = cleanup_failures
            primary_error = BackupError(
                first.stage,
                first.message,
                cause=first.exception,
                secondary_failures=rest,
                completed_stages=(
                    FailureStage.SNAPSHOT_COPY,
                    FailureStage.SNAPSHOT_VALIDATION,
                    FailureStage.SNAPSHOT_READ,
                ),
            )

    if primary_error is not None:
        raise primary_error from primary_error.cause
    assert snapshot_bytes is not None
    return snapshot_bytes


def _dated_name_pattern(database_path: os.PathLike[str] | str) -> re.Pattern[str]:
    name = _database_basename(database_path)
    suffix = Path(name).suffix
    stem = name[: -len(suffix)] if suffix else name
    return re.compile(
        rf"^{re.escape(stem)}-(?P<date>\d{{4}}-\d{{2}}-\d{{2}}){re.escape(suffix)}$"
    )


def _matching_dated_entries(
    entries: Iterable[Any],
    database_path: os.PathLike[str] | str,
    folder: str,
) -> list[tuple[dt.date, str]]:
    pattern = _dated_name_pattern(database_path)
    matches: list[tuple[dt.date, str]] = []
    for entry in entries:
        name = getattr(entry, "name", None)
        if not isinstance(name, str):
            continue
        match = pattern.fullmatch(name)
        if match is None:
            continue
        try:
            parsed_date = dt.date.fromisoformat(match.group("date"))
        except ValueError:
            continue
        matches.append((parsed_date, f"{folder}/{name}" if folder else f"/{name}"))
    return matches


def _is_retryable_dropbox_error(error: Exception) -> bool:
    if isinstance(error, (requests.exceptions.ConnectionError, requests.exceptions.Timeout)):
        return True
    if _dropbox_exceptions is None:
        return False
    if isinstance(
        error,
        (
            _dropbox_exceptions.InternalServerError,
            _dropbox_exceptions.RateLimitError,
        ),
    ):
        return True
    if isinstance(error, _dropbox_exceptions.HttpError) and not isinstance(
        error,
        (_dropbox_exceptions.AuthError, _dropbox_exceptions.BadInputError),
    ):
        return error.status_code in {408, 429} or error.status_code >= 500
    return False


def _retry_dropbox_call(
    operation: Callable[[], Any],
    operation_name: str,
    *,
    retry_delays: tuple[float, ...],
    sleep: Callable[[float], None],
) -> Any:
    total_attempts = len(retry_delays) + 1
    for attempt in range(1, total_attempts + 1):
        try:
            return operation()
        except Exception as error:
            if attempt >= total_attempts or not _is_retryable_dropbox_error(error):
                raise
            delay = retry_delays[attempt - 1]
            provider_backoff = getattr(error, "backoff", None)
            if isinstance(provider_backoff, (int, float)) and provider_backoff > delay:
                delay = min(float(provider_backoff), 30.0)
            logging.warning(
                "Transient Dropbox %s failure (%s); retrying attempt %d/%d in %.1f seconds",
                operation_name,
                type(error).__name__,
                attempt + 1,
                total_attempts,
                delay,
            )
            sleep(delay)
    raise AssertionError("Dropbox retry loop terminated unexpectedly")


def prune_dated_backups(
    client: Any,
    folder: str,
    database_path: os.PathLike[str] | str,
    retention: object | None = None,
    *,
    retry_delays: tuple[float, ...] = DEFAULT_DROPBOX_RETRY_DELAYS,
    sleep: Callable[[float], None] = time.sleep,
) -> tuple[str, ...]:
    """Fully paginate a folder and delete only this database's oldest dated names."""
    keep = validate_dated_retention(retention)
    folder_path = _normalize_dropbox_folder(folder)
    completed: list[FailureStage] = []
    try:
        page = _retry_dropbox_call(
            lambda: client.files_list_folder(folder_path),
            "folder listing",
            retry_delays=retry_delays,
            sleep=sleep,
        )
        completed.append(FailureStage.RETENTION_LIST)
    except BaseException as exc:
        raise BackupError(
            FailureStage.RETENTION_LIST,
            f"could not list Dropbox folder {folder_path or '/'}",
            cause=exc,
        ) from exc

    entries = list(getattr(page, "entries", ()))
    while bool(getattr(page, "has_more", False)):
        try:
            cursor = page.cursor
            page = _retry_dropbox_call(
                lambda cursor=cursor: client.files_list_folder_continue(cursor),
                "folder-list pagination",
                retry_delays=retry_delays,
                sleep=sleep,
            )
            entries.extend(getattr(page, "entries", ()))
        except BaseException as exc:
            raise BackupError(
                FailureStage.RETENTION_PAGINATION,
                f"could not continue listing Dropbox folder {folder_path or '/'}",
                cause=exc,
                completed_stages=completed,
            ) from exc

    matches = sorted(_matching_dated_entries(entries, database_path, folder_path))
    delete_count = max(0, len(matches) - keep)
    deleted: list[str] = []
    for _date, remote_path in matches[:delete_count]:
        try:
            _retry_dropbox_call(
                lambda remote_path=remote_path: client.files_delete_v2(remote_path),
                "retention deletion",
                retry_delays=retry_delays,
                sleep=sleep,
            )
            deleted.append(remote_path)
        except BaseException as exc:
            raise BackupError(
                FailureStage.RETENTION_DELETE,
                f"could not delete expired dated backup {remote_path}",
                cause=exc,
                completed_stages=completed,
            ) from exc
    return tuple(deleted)


def publish_snapshot(
    client: Any,
    snapshot_bytes: bytes,
    folder: str,
    database_path: os.PathLike[str] | str,
    *,
    retention: object | None = None,
    backup_date: dt.date | None = None,
    upload_mode: Any = DROPBOX_OVERWRITE_MODE,
    retry_delays: tuple[float, ...] = DEFAULT_DROPBOX_RETRY_DELAYS,
    sleep: Callable[[float], None] = time.sleep,
) -> PublicationResult:
    """Publish identical bytes dated-first/latest-second, then enforce retention."""
    keep = validate_dated_retention(retention)
    paths = dropbox_backup_paths(folder, database_path, backup_date)
    if not isinstance(snapshot_bytes, bytes):
        raise BackupError(
            FailureStage.CONFIGURATION,
            "snapshot_bytes must be bytes",
        )

    try:
        _retry_dropbox_call(
            lambda: client.files_upload(snapshot_bytes, paths.dated, mode=upload_mode),
            "dated upload",
            retry_delays=retry_delays,
            sleep=sleep,
        )
    except BaseException as exc:
        raise BackupError(
            FailureStage.DATED_UPLOAD,
            f"could not upload dated backup {paths.dated}",
            cause=exc,
        ) from exc

    try:
        _retry_dropbox_call(
            lambda: client.files_upload(snapshot_bytes, paths.latest, mode=upload_mode),
            "latest upload",
            retry_delays=retry_delays,
            sleep=sleep,
        )
    except BaseException as exc:
        raise BackupError(
            FailureStage.LATEST_UPLOAD,
            f"dated backup was uploaded but latest update failed for {paths.latest}",
            cause=exc,
            completed_stages=(FailureStage.DATED_UPLOAD,),
        ) from exc

    try:
        deleted = prune_dated_backups(
            client,
            paths.folder,
            database_path,
            keep,
            retry_delays=retry_delays,
            sleep=sleep,
        )
    except BackupError as exc:
        exc.add_completed_prefix(
            (FailureStage.DATED_UPLOAD, FailureStage.LATEST_UPLOAD)
        )
        raise
    return PublicationResult(paths, deleted)


def backup_database_to_dropbox(
    database_path: os.PathLike[str] | str,
    folder: str,
    *,
    client_factory: Callable[[], Any],
    retention: object | None = None,
    backup_date: dt.date | None = None,
    upload_mode: Any = DROPBOX_OVERWRITE_MODE,
    busy_timeout_seconds: float = SQLITE_BUSY_TIMEOUT_SECONDS,
    deadline_seconds: float = DEFAULT_BACKUP_DEADLINE_SECONDS,
    retry_delays: tuple[float, ...] = DEFAULT_DROPBOX_RETRY_DELAYS,
    sleep: Callable[[float], None] = time.sleep,
) -> BackupResult:
    """Create a validated local snapshot and publish it through an injected client."""
    keep = validate_dated_retention(retention)
    paths = dropbox_backup_paths(folder, database_path, backup_date)
    data = snapshot_database(
        database_path,
        busy_timeout_seconds=busy_timeout_seconds,
        deadline_seconds=deadline_seconds,
    )
    try:
        client = client_factory()
    except BaseException as exc:
        raise BackupError(
            FailureStage.DROPBOX_CLIENT,
            "could not create Dropbox client",
            cause=exc,
            completed_stages=(
                FailureStage.SNAPSHOT_COPY,
                FailureStage.SNAPSHOT_VALIDATION,
                FailureStage.SNAPSHOT_READ,
            ),
        ) from exc

    result = publish_snapshot(
        client,
        data,
        paths.folder,
        database_path,
        retention=keep,
        backup_date=backup_date,
        upload_mode=upload_mode,
        retry_delays=retry_delays,
        sleep=sleep,
    )
    return BackupResult(result.paths, result.deleted_paths, len(data))


def _fsync_directory(directory: Path) -> None:
    descriptor = os.open(directory, os.O_RDONLY | getattr(os, "O_DIRECTORY", 0))
    try:
        os.fsync(descriptor)
    finally:
        os.close(descriptor)


def _active_database_paths(target: Path) -> tuple[Path, Path, Path]:
    return target, Path(f"{target}-wal"), Path(f"{target}-shm")


def _recover_preserved_set(
    active_paths: tuple[Path, Path, Path],
    original_names: set[str],
    rollback_directory: Path,
) -> list[StageFailure]:
    failures: list[StageFailure] = []
    for active_path in active_paths:
        preserved_path = rollback_directory / active_path.name
        try:
            if active_path.name in original_names:
                if preserved_path.exists():
                    os.replace(preserved_path, active_path)
            elif active_path.exists() or active_path.is_symlink():
                active_path.unlink()
        except BaseException as exc:
            failures.append(
                StageFailure(
                    FailureStage.RESTORE_RECOVERY,
                    f"could not recover {active_path.name}",
                    exc,
                )
            )
    if not failures:
        try:
            _fsync_directory(active_paths[0].parent)
        except BaseException as exc:
            failures.append(
                StageFailure(
                    FailureStage.RESTORE_RECOVERY,
                    "could not fsync recovered database directory",
                    exc,
                )
            )
    return failures


def install_snapshot(
    snapshot_path: os.PathLike[str] | str,
    target_database_path: os.PathLike[str] | str,
    *,
    confirmed_stopped: bool,
    install_replace: Callable[[os.PathLike[str] | str, os.PathLike[str] | str], None] = os.replace,
    connect: Callable[..., sqlite3.Connection] = sqlite3.connect,
) -> RestoreResult:
    """Validate and atomically install a standalone snapshot while users are stopped.

    The caller must explicitly confirm that all database users are stopped.  On
    success the old database/WAL/SHM set remains in the returned unique rollback
    directory.  Any failure after preservation attempts to put that complete set
    back before raising a staged ``RestoreError``.
    """
    if confirmed_stopped is not True:
        raise RestoreError(
            FailureStage.RESTORE_CONFIRMATION,
            "restore requires explicit confirmation that all database users are stopped",
        )

    candidate = Path(snapshot_path)
    target = Path(target_database_path)
    try:
        if candidate.resolve() == target.resolve():
            raise ValueError("snapshot and target database must be different files")
        validate_snapshot(candidate, connect=connect)
    except BaseException as exc:
        if isinstance(exc, RestoreError):
            raise
        raise RestoreError(
            FailureStage.RESTORE_VALIDATION,
            f"restore candidate {candidate.name} failed validation",
            cause=exc,
        ) from exc

    if not target.parent.is_dir():
        raise RestoreError(
            FailureStage.RESTORE_COPY,
            f"target directory does not exist: {target.parent}",
        )

    temporary_path: Path | None = None
    rollback_directory: Path | None = None
    active_paths = _active_database_paths(target)
    original_names = {path.name for path in active_paths if path.exists()}
    preserved_paths: list[Path] = []
    mutation_started = False
    current_stage = FailureStage.RESTORE_COPY
    primary_error: RestoreError | None = None

    try:
        descriptor, raw_temporary = tempfile.mkstemp(
            prefix=f".{target.name}.restore-",
            suffix=".db",
            dir=target.parent,
        )
        temporary_path = Path(raw_temporary)
        with os.fdopen(descriptor, "wb") as destination, candidate.open("rb") as source:
            shutil.copyfileobj(source, destination)
            destination.flush()
            os.fsync(destination.fileno())

        current_stage = FailureStage.RESTORE_PRESERVE
        rollback_directory = Path(
            tempfile.mkdtemp(
                prefix=f".{target.name}.rollback-",
                dir=target.parent,
            )
        )
        mutation_started = True
        for active_path in active_paths:
            if active_path.exists():
                preserved_path = rollback_directory / active_path.name
                os.replace(active_path, preserved_path)
                preserved_paths.append(preserved_path)
        _fsync_directory(target.parent)

        current_stage = FailureStage.RESTORE_INSTALL
        install_replace(temporary_path, target)
        temporary_path = None
        _fsync_directory(target.parent)
    except BaseException as exc:
        primary_error = RestoreError(
            current_stage,
            f"restore failed while installing {candidate.name}",
            cause=exc,
        )
    finally:
        secondary: list[StageFailure] = []
        if primary_error is not None and mutation_started and rollback_directory is not None:
            secondary.extend(
                _recover_preserved_set(
                    active_paths,
                    original_names,
                    rollback_directory,
                )
            )
        if temporary_path is not None:
            secondary.extend(
                _cleanup_paths(
                    (temporary_path,),
                    lambda path: path.unlink(),
                    FailureStage.RESTORE_CLEANUP,
                )
            )
        if (
            primary_error is not None
            and rollback_directory is not None
            and not secondary
        ):
            try:
                rollback_directory.rmdir()
            except BaseException as exc:
                secondary.append(
                    StageFailure(
                        FailureStage.RESTORE_CLEANUP,
                        f"could not remove empty rollback directory {rollback_directory.name}",
                        exc,
                    )
                )
        if primary_error is not None:
            primary_error.add_secondary(secondary)

    if primary_error is not None:
        raise primary_error from primary_error.cause
    assert rollback_directory is not None
    return RestoreResult(target, rollback_directory, tuple(preserved_paths))


__all__ = [
    "BackupError",
    "BackupResult",
    "DEFAULT_BACKUP_DEADLINE_SECONDS",
    "DEFAULT_DATED_RETENTION",
    "DROPBOX_OVERWRITE_MODE",
    "DropboxBackupPaths",
    "FailureStage",
    "PublicationResult",
    "RestoreError",
    "RestoreResult",
    "SQLITE_BUSY_TIMEOUT_SECONDS",
    "SnapshotValidationError",
    "StageFailure",
    "backup_database_to_dropbox",
    "dated_backup_name",
    "dropbox_backup_paths",
    "install_snapshot",
    "prune_dated_backups",
    "publish_snapshot",
    "snapshot_database",
    "validate_dated_retention",
    "validate_snapshot",
]
