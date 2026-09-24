"""Candidate-runtime SQLite recovery and schema probes for a stopped cutover.

The deploy orchestrator must first quiesce all known writers and close public
ingress. This module does not stop services or authorize activation by itself.
"""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import shutil
import stat

import backup
import schema_migrations


class RecoveryCheckError(RuntimeError):
    pass


def _private_directory(path: Path) -> None:
    if path.is_symlink() or not path.is_dir():
        raise RecoveryCheckError("recovery directory is unavailable")
    info = path.stat()
    if info.st_uid != 0 or stat.S_IMODE(info.st_mode) != 0o700:
        raise RecoveryCheckError("recovery directory must be root-owned mode 0700")


def _fsync_directory(path: Path) -> None:
    descriptor = os.open(path, os.O_RDONLY | os.O_DIRECTORY)
    try:
        os.fsync(descriptor)
    finally:
        os.close(descriptor)


def durable_snapshot(database: Path, destination: Path, *, verify_private=True) -> Path:
    """Stream the validated SQLite backup into a durable root-private file."""
    if not database.is_absolute() or not database.is_file() or database.is_symlink():
        raise RecoveryCheckError("approved database file is missing or invalid")
    if not destination.is_absolute() or destination.exists() or destination.is_symlink():
        raise RecoveryCheckError("snapshot destination must be a new absolute path")
    if verify_private:
        _private_directory(destination.parent)

    def publish(source: Path) -> bytes:
        descriptor = os.open(destination, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        try:
            with os.fdopen(descriptor, "wb") as target, source.open("rb") as stream:
                shutil.copyfileobj(stream, target, length=1024 * 1024)
                target.flush()
                os.fsync(target.fileno())
        except BaseException:
            destination.unlink(missing_ok=True)
            raise
        return b""  # snapshot_database's reader hook requires a bytes result.

    try:
        backup.snapshot_database(database, read_snapshot=publish)
        backup.validate_snapshot(destination)
        _fsync_directory(destination.parent)
        return destination
    except BaseException:
        destination.unlink(missing_ok=True)
        raise


def schema_gate(database: Path) -> dict[str, str]:
    """Classify without creating a database or running Alembic migrations."""
    if not database.is_absolute() or database.is_symlink():
        raise RecoveryCheckError("approved database path is invalid")
    heads = schema_migrations.head_revisions()
    if len(heads) != 1:
        raise RecoveryCheckError("candidate migration history has no single head")
    state = schema_migrations.inspect_schema(database)
    if state.kind == "head" and state.revision == heads[0]:
        gate = "ready"
    elif state.kind in {"baseline", "versioned"}:
        gate = "migration_required"
    else:
        gate = "refused"
    return {"gate": gate, "state": state.kind, "revision": state.revision or "",
            "candidate_head": heads[0]}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("snapshot", "schema"))
    parser.add_argument("--database", type=Path, required=True)
    parser.add_argument("--destination", type=Path)
    args = parser.parse_args()
    if os.geteuid() != 0:
        parser.error("recovery probes must run as root in the stopped cutover")
    try:
        if args.action == "snapshot":
            if args.destination is None:
                parser.error("snapshot requires --destination")
            path = durable_snapshot(args.database, args.destination)
            print(json.dumps({"snapshot": str(path), "validated": True}, sort_keys=True))
        else:
            if args.destination is not None:
                parser.error("schema does not accept --destination")
            print(json.dumps(schema_gate(args.database), sort_keys=True))
        return 0
    except (RecoveryCheckError, backup.BackupError, OSError) as exc:
        # Avoid exception messages; SQLite or filesystem errors may contain
        # sensitive source paths or content.
        print(json.dumps({"error": type(exc).__name__}, sort_keys=True))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
