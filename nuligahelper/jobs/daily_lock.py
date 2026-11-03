"""Linux/Raspberry Pi process lock for the daily job.

The lock is advisory and therefore coordinates nuLigaHelper processes that use
this helper.  It intentionally leaves the lock file in place: ownership belongs
to the open file descriptor and is released by the kernel when the descriptor
is closed or the process exits.
"""

from __future__ import annotations

import errno
import fcntl
import os
from collections.abc import Generator
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path


class DailyRunLockError(RuntimeError):
    """Base class for actionable daily-run lock failures."""


class DailyRunAlreadyActiveError(DailyRunLockError):
    """Raised when another process already owns a database's daily-run lock."""

    database_path: Path
    lock_path: Path

    def __init__(self, database_path: Path, lock_path: Path) -> None:
        self.database_path = database_path
        self.lock_path = lock_path
        message = f"Another daily run already owns the lock for database '{database_path}'"
        message += f" (lock file '{lock_path}'). Refusing to overlap."
        super().__init__(message)


class DailyRunLockPathError(DailyRunLockError):
    """Raised when the configured path cannot provide a usable lock file."""

    database_path: Path | None
    lock_path: Path | None

    def __init__(
        self,
        message: str,
        *,
        database_path: Path | None = None,
        lock_path: Path | None = None,
    ) -> None:
        self.database_path = database_path
        self.lock_path = lock_path
        super().__init__(message)


@dataclass(frozen=True)
class DailyRunLock:
    """Canonical database and lock paths for an acquired daily-run lock."""

    database_path: Path
    lock_path: Path


def canonical_database_path(database_path: str | os.PathLike[str]) -> Path:
    """Return the absolute, symlink-resolved identity of *database_path*."""
    try:
        raw_path = os.fspath(database_path)
    except TypeError as exc:
        raise DailyRunLockPathError(
            "The database path must be a non-empty filesystem path."
        ) from exc

    if not raw_path.strip():
        raise DailyRunLockPathError(
            "The database path must be a non-empty text filesystem path."
        )

    try:
        return Path(raw_path).expanduser().resolve(strict=False)
    except (OSError, RuntimeError) as exc:
        message = f"Cannot resolve database path '{raw_path}': {exc}. "
        message += "Configure an existing local Linux/Raspberry Pi filesystem path."
        raise DailyRunLockPathError(message) from exc


def daily_lock_path(database_path: str | os.PathLike[str]) -> Path:
    """Return ``<canonical database path>.daily.lock``."""
    canonical_path = canonical_database_path(database_path)
    return Path(f"{canonical_path}.daily.lock")


@contextmanager
def daily_run_lock(
    database_path: str | os.PathLike[str],
) -> Generator[DailyRunLock, None, None]:
    """Acquire the database's non-blocking exclusive daily-run lock.

    The database's parent directory must already exist and be writable.  The
    helper does not create it because doing so could hide a bad production
    database configuration.  The lock file itself may safely remain after use.
    ``fcntl.flock`` limits this guard to the supported local Linux/Pi setup.
    """
    canonical_path = canonical_database_path(database_path)
    lock_path = Path(f"{canonical_path}.daily.lock")
    parent = lock_path.parent

    if not parent.exists():
        message = f"Cannot create daily-run lock '{lock_path}': parent directory "
        message += f"'{parent}' does not exist. Configure the database in an existing, "
        message += "writable local Linux/Raspberry Pi directory."
        raise DailyRunLockPathError(
            message, database_path=canonical_path, lock_path=lock_path
        )
    if not parent.is_dir():
        message = f"Cannot create daily-run lock '{lock_path}': parent path '{parent}' "
        message += "is not a directory. Configure a writable local Linux/Raspberry Pi "
        message += "database directory."
        raise DailyRunLockPathError(
            message, database_path=canonical_path, lock_path=lock_path
        )

    flags = os.O_RDWR | os.O_CREAT
    if hasattr(os, "O_CLOEXEC"):
        flags |= os.O_CLOEXEC

    try:
        descriptor = os.open(lock_path, flags, 0o600)
    except OSError as exc:
        message = f"Cannot open daily-run lock '{lock_path}': {exc}. Ensure '{parent}' "
        message += "is writable and is on a local Linux/Raspberry Pi filesystem."
        raise DailyRunLockPathError(
            message, database_path=canonical_path, lock_path=lock_path
        ) from exc

    try:
        try:
            fcntl.flock(descriptor, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except OSError as exc:
            if exc.errno in (errno.EACCES, errno.EAGAIN):
                raise DailyRunAlreadyActiveError(canonical_path, lock_path) from exc
            message = f"Cannot acquire daily-run lock '{lock_path}': {exc}. Ensure the "
            message += "database and lock are on a local Linux/Raspberry Pi filesystem."
            raise DailyRunLockPathError(
                message, database_path=canonical_path, lock_path=lock_path
            ) from exc

        yield DailyRunLock(canonical_path, lock_path)
    finally:
        os.close(descriptor)
