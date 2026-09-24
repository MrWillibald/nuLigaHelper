#!/usr/bin/env python3
"""Root-operated release preparation. Activation is intentionally not implemented yet.

This command never changes ``current``, systemd state, ingress, or the database.
Install a reviewed copy outside the service-readable application tree. Until the
activation workflow is implemented and rehearsed, this command supports only
``prepare`` and ``inspect``; it cannot cut over production by accident.
"""

from __future__ import annotations

import argparse
import fcntl
import grp
import io
import json
import os
from pathlib import Path, PurePosixPath
import re
import shutil
import stat
import subprocess
import sys
import tarfile
import tempfile
from datetime import datetime, timezone
from uuid import uuid4


SHA = re.compile(r"[0-9a-f]{40}\Z")
DEFAULT_CONFIG = Path("/etc/nuligahelper/deployment.json")
DEFAULT_LOCK = Path("/run/lock/nuligahelper-deploy.lock")
MIN_FREE_BYTES = 2 * 1024**3
MIN_AVAILABLE_MEMORY_KIB = 256 * 1024


class DeployError(RuntimeError):
    """A non-secret, operator-actionable deployment refusal."""


def run(*args: str, cwd: Path | None = None, env: dict[str, str] | None = None) -> str:
    result = subprocess.run(args, cwd=cwd, env=env, text=True, stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE, check=False)
    if result.returncode:
        # Never echo stderr: Git and pip errors may include a credential-bearing URL.
        raise DeployError(f"command failed: {Path(args[0]).name} (exit {result.returncode})")
    return result.stdout.strip()


def safe_environment() -> dict[str, str]:
    """No application, Git, pip, or provider credentials reach the test suite."""
    return {"PATH": "/usr/local/bin:/usr/bin:/bin", "LANG": "C.UTF-8",
            "HOME": "/nonexistent", "GIT_TERMINAL_PROMPT": "0",
            "PYTHONDONTWRITEBYTECODE": "1"}


def git_environment() -> dict[str, str]:
    """Allow only root's read-only Git/SSH setup, never application secrets."""
    environment = safe_environment()
    environment["HOME"] = "/root"
    return environment


def load_config(path: Path) -> dict[str, object]:
    try:
        info = path.lstat()
        if path.is_symlink() or info.st_uid != 0 or info.st_mode & 0o077:
            raise DeployError("deployment config must be root-owned mode 0600")
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise DeployError("deployment config is unavailable or invalid") from exc
    required = {"source", "app_root", "source_cache", "service_group", "python"}
    if not isinstance(data, dict) or set(data) != required:
        raise DeployError("deployment config keys are invalid")
    for key in required:
        if not isinstance(data[key], str) or not data[key]:
            raise DeployError(f"deployment config {key} is invalid")
    for key in ("app_root", "source_cache", "python"):
        value = Path(data[key])
        if not value.is_absolute() or ".." in value.parts:
            raise DeployError(f"deployment config {key} must be absolute")
    if data["app_root"] != "/opt/nuligahelper":
        raise DeployError("app_root differs from the approved production root")
    if Path(data["source_cache"]).is_relative_to(data["app_root"]):
        raise DeployError("source cache must be outside the application tree")
    if not re.fullmatch(r"[a-z_][a-z0-9_-]*", data["service_group"]):
        raise DeployError("service_group is invalid")
    if any(marker in data["source"].lower() for marker in ("@", "token", "password")) and \
            data["source"].startswith("https://"):
        raise DeployError("source URL must not contain an embedded credential")
    return data


def exclusive_lock(path: Path):
    path.parent.mkdir(mode=0o700, parents=True, exist_ok=True)
    descriptor = os.open(path, os.O_CREAT | os.O_RDWR | os.O_CLOEXEC, 0o600)
    try:
        fcntl.flock(descriptor, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except BlockingIOError as exc:
        os.close(descriptor)
        raise DeployError("another deployment attempt holds the host lock") from exc
    return descriptor


def available_memory_kib() -> int:
    for line in Path("/proc/meminfo").read_text().splitlines():
        if line.startswith("MemAvailable:"):
            return int(line.split()[1])
    raise DeployError("available memory could not be measured")


def check_headroom(path: Path) -> None:
    if shutil.disk_usage(path).free < MIN_FREE_BYTES:
        raise DeployError("insufficient free disk space for a second release and snapshot")
    if available_memory_kib() < MIN_AVAILABLE_MEMORY_KIB:
        raise DeployError("insufficient available memory for release preparation")


def check_cache_directory(path: Path) -> None:
    """Git objects and credentials must not be controlled by the service user."""
    for folder in (path.parent, path):
        if folder.is_symlink():
            raise DeployError("source cache path must not be a symlink")
        if folder.exists():
            info = folder.stat()
            if not stat.S_ISDIR(info.st_mode) or info.st_uid != 0 or info.st_mode & 0o077:
                raise DeployError("source cache must be root-owned and private")


def fetch_master(cache: Path, source: str) -> str:
    check_cache_directory(cache)
    cache.parent.mkdir(mode=0o700, parents=True, exist_ok=True)
    env = git_environment()
    if not cache.exists():
        cache.mkdir(mode=0o700)
        run("git", "init", "--bare", str(cache), env=env)
        run("git", "-C", str(cache), "remote", "add", "origin", source, env=env)
    check_cache_directory(cache)
    if not (cache / "HEAD").is_file():
        raise DeployError("source cache is not a bare Git repository")
    configured = run("git", "-C", str(cache), "remote", "get-url", "origin", env=env)
    if configured != source:
        raise DeployError("source cache origin differs from approved source")
    run("git", "-C", str(cache), "fetch", "--no-tags", "origin",
        "+refs/heads/master:refs/remotes/origin/master", env=env)
    master = run("git", "-C", str(cache), "rev-parse", "--verify",
                 "refs/remotes/origin/master^{commit}", env=env)
    if not SHA.fullmatch(master):
        raise DeployError("fetched master did not resolve to a full commit")
    return master


def select_commit(cache: Path, master: str, requested: str | None) -> tuple[str, str]:
    selected = requested or master
    if not SHA.fullmatch(selected):
        raise DeployError("requested commit must be a full lowercase SHA-1")
    env = git_environment()
    resolved = run("git", "-C", str(cache), "rev-parse", "--verify",
                   f"{selected}^{{commit}}", env=env)
    if resolved != selected:
        raise DeployError("requested commit does not identify a commit")
    result = subprocess.run(("git", "-C", str(cache), "merge-base", "--is-ancestor",
                             selected, master), env=env, capture_output=True, check=False)
    if result.returncode:
        raise DeployError("requested commit is not reachable from fetched master")
    tree = run("git", "-C", str(cache), "show", "-s", "--format=%T", selected, env=env)
    if not SHA.fullmatch(tree):
        raise DeployError("selected tree identity is invalid")
    return selected, tree


def export_commit(cache: Path, commit: str, stage: Path) -> None:
    archive = subprocess.run(("git", "-C", str(cache), "archive", "--format=tar", commit),
                             env=git_environment(), stdout=subprocess.PIPE,
                             stderr=subprocess.PIPE, check=False)
    if archive.returncode:
        raise DeployError("source archive failed")
    with tarfile.open(fileobj=io.BytesIO(archive.stdout), mode="r:") as bundle:
        members = bundle.getmembers()
        for member in members:
            path = PurePosixPath(member.name)
            if path.is_absolute() or ".." in path.parts or not (member.isfile() or member.isdir()):
                raise DeployError("source archive contains an unsafe member")
            if any(part in {"venv", "deploy", ".git"} for part in path.parts) or \
                    path.name in {"config.json", ".nuligahelper_secret"}:
                raise DeployError("source archive contains local or private state")
        bundle.extractall(stage, members=members, filter="data")


def prior_release(app_root: Path) -> str | None:
    current = app_root / "current"
    if not current.is_symlink():
        if current.exists():
            raise DeployError("current exists but is not a symlink")
        return None
    target = current.resolve(strict=True)
    if not target.is_relative_to(app_root.resolve()):
        raise DeployError("current points outside the approved application root")
    return str(target)


def prepared_candidate(app_root: Path, commit: str) -> Path:
    """Return only a complete, unsymlinked candidate with a matching marker."""
    if not SHA.fullmatch(commit):
        raise DeployError("candidate commit must be a full SHA")
    candidate = app_root / "releases" / commit
    if candidate.is_symlink() or not candidate.is_dir():
        raise DeployError("candidate release is unavailable")
    marker = candidate / ".prepared.json"
    if marker.is_symlink():
        raise DeployError("candidate preparation marker is invalid")
    try:
        record = json.loads(marker.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise DeployError("candidate preparation marker is unavailable") from exc
    tree = record.get("tree") if isinstance(record, dict) else None
    if not isinstance(tree, str) or record.get("schema_version") != 1 or \
            record.get("commit") != commit or not SHA.fullmatch(tree):
        raise DeployError("candidate preparation marker does not match the release")
    return candidate


def switch_current(app_root: Path, commit: str) -> tuple[str, str]:
    """Atomically replace an internal link; caller must first pass cutover gates.

    This is intentionally not exposed as a CLI action until writer quiescence,
    snapshot, schema, and ingress gates are implemented and rehearsed.
    """
    candidate = prepared_candidate(app_root, commit)
    previous = prior_release(app_root)
    if previous is None or previous == str(candidate):
        raise DeployError("current release is missing or already selected")
    temporary = app_root / (".current-next-" + uuid4().hex)
    try:
        os.symlink(candidate, temporary)
        os.replace(temporary, app_root / "current")
        directory_fd = os.open(app_root, os.O_RDONLY | os.O_DIRECTORY)
        try:
            os.fsync(directory_fd)
        finally:
            os.close(directory_fd)
    finally:
        temporary.unlink(missing_ok=True)
    return previous, str(candidate)


def normalize_permissions(stage: Path, group_name: str) -> None:
    gid = grp.getgrnam(group_name).gr_gid
    for root, dirs, files in os.walk(stage, followlinks=False):
        for name in dirs + files:
            path = Path(root) / name
            if path.is_symlink():
                target = path.resolve(strict=True)
                if not str(target).startswith(("/usr/", str(stage) + "/")):
                    raise DeployError("candidate contains an external symlink")
                os.lchown(path, 0, gid)
                continue
            mode = path.stat().st_mode
            os.chown(path, 0, gid)
            os.chmod(path, 0o750 if path.is_dir() or mode & 0o100 else 0o640)
    os.chown(stage, 0, gid)
    os.chmod(stage, 0o750)


def ensure_releases_directory(app_root: Path, group_name: str) -> Path:
    releases = app_root / "releases"
    gid = grp.getgrnam(group_name).gr_gid
    if releases.is_symlink():
        raise DeployError("releases directory must not be a symlink")
    if not releases.exists():
        releases.mkdir(mode=0o750)
        os.chown(releases, 0, gid)
        os.chmod(releases, 0o750)
    info = releases.stat()
    if not stat.S_ISDIR(info.st_mode) or (info.st_uid, info.st_gid) != (0, gid) or \
            stat.S_IMODE(info.st_mode) != 0o750:
        raise DeployError("releases directory is not root-owned and service-readable")
    return releases


def verify_unit_syntax(stage: Path) -> None:
    """Parse every versioned unit against the prepared executable paths."""
    assets = stage / "release-assets/systemd"
    units = sorted(path for path in assets.glob("nuligahelper-*")
                   if path.suffix in {".service", ".timer"})
    if len([unit for unit in units if unit.suffix == ".service"]) != 9:
        raise DeployError("candidate does not include the complete service set")
    with tempfile.TemporaryDirectory(prefix="nuligahelper-units-") as directory:
        temporary = Path(directory)
        for unit in units:
            content = unit.read_text(encoding="utf-8")
            if unit.suffix == ".service" and "/opt/nuligahelper/current" not in content:
                raise DeployError("candidate unit does not use the current release link")
            (temporary / unit.name).write_text(
                content.replace("/opt/nuligahelper/current", str(stage)), encoding="utf-8")
        run("systemd-analyze", "verify", *(str(temporary / unit.name) for unit in units),
            env=safe_environment())


def prepare(config: dict[str, object], requested: str | None = None) -> dict[str, object]:
    app_root = Path(config["app_root"])
    cache = Path(config["source_cache"])
    if app_root.is_symlink() or not app_root.is_dir():
        raise DeployError("approved application root is unavailable")
    check_headroom(app_root)
    master = fetch_master(cache, str(config["source"]))
    commit, tree = select_commit(cache, master, requested)
    releases = ensure_releases_directory(app_root, str(config["service_group"]))
    destination = releases / commit
    if destination.exists():
        raise DeployError("candidate release already exists; inspect it before retrying")
    previous = prior_release(app_root)
    # Virtualenv console scripts embed absolute interpreter paths. Build at the
    # final SHA path, never in a temporary directory that will be renamed.
    # An incomplete path has no trusted marker and is removed on failure.
    destination.mkdir(mode=0o700)
    try:
        export_commit(cache, commit, destination)
        if not (destination / "requirements-production.txt").is_file() or \
                not (destination / "requirements-test.txt").is_file() or \
                not (destination / "release-assets/gunicorn.conf.py").is_file() or \
                not (destination / "release-assets/recovery_check.py").is_file():
            raise DeployError("selected commit lacks required release assets")
        env = safe_environment()
        run(str(config["python"]), "-m", "venv", str(destination / "venv"), env=env)
        interpreter = str(destination / "venv/bin/python")
        run(interpreter, "-m", "pip", "install", "--disable-pip-version-check",
            "-r", str(destination / "requirements-test.txt"), cwd=destination, env=env)
        run(interpreter, "-m", "pip", "check", cwd=destination, env=env)
        run("bash", str(destination / "test/run_tests.sh"), cwd=destination, env=env)
        run(interpreter, "-m", "compileall", "-q", str(destination), cwd=destination, env=env)
        run(interpreter, "-m", "py_compile", str(destination / "release-assets/gunicorn.conf.py"),
            cwd=destination, env=env)
        verify_unit_syntax(destination)
        record = {"schema_version": 1, "commit": commit, "tree": tree,
                  "fetched_master": master, "previous_release": previous,
                  "prepared_at": datetime.now(timezone.utc).isoformat(),
                  "checks": ["pip-check", "offline-suite", "python-syntax", "systemd-syntax"]}
        marker = destination / ".prepared.json"
        with marker.open("x", encoding="utf-8") as handle:
            json.dump(record, handle, sort_keys=True)
            handle.flush()
            os.fsync(handle.fileno())
        normalize_permissions(destination, str(config["service_group"]))
        directory_fd = os.open(releases, os.O_RDONLY | os.O_DIRECTORY)
        try:
            os.fsync(directory_fd)
        finally:
            os.close(directory_fd)
        return record
    except BaseException:
        # This is the unique destination we just created; never remove a prior
        # candidate, current link target, or live database on failure.
        if destination.exists():
            shutil.rmtree(destination)
        raise


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG)
    parser.add_argument("--lock", type=Path, default=DEFAULT_LOCK)
    subcommands = parser.add_subparsers(dest="action", required=True)
    selection = subcommands.add_parser("prepare", help="fetch and prepare, never activate")
    selection.add_argument("--sha", help="full commit SHA reachable from fetched master")
    subcommands.add_parser("inspect", help="show configured paths, no secret values")
    args = parser.parse_args()
    if os.geteuid() != 0:
        parser.error("deployment command must run as root")
    try:
        config = load_config(args.config)
        if args.action == "inspect":
            print(json.dumps({"app_root": config["app_root"],
                              "current": prior_release(Path(config["app_root"])),
                              "source_cache": config["source_cache"]}, sort_keys=True))
            return 0
        lock_fd = exclusive_lock(args.lock)
        try:
            result = prepare(config, args.sha)
        finally:
            os.close(lock_fd)
        print(json.dumps(result, sort_keys=True))
        return 0
    except (DeployError, OSError) as exc:
        print(f"deployment refused: {exc if isinstance(exc, DeployError) else type(exc).__name__}",
              file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
