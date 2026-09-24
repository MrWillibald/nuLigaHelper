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
import http.client
import io
import json
import os
from pathlib import Path, PurePosixPath
import re
import shutil
import ssl
import stat
import subprocess
import sys
import tarfile
import tempfile
import time
from datetime import datetime, time as daytime, timedelta, timezone
from uuid import uuid4
from urllib.parse import urlsplit
from zoneinfo import ZoneInfo


SHA = re.compile(r"[0-9a-f]{40}\Z")
DEFAULT_CONFIG = Path("/etc/nuligahelper/deployment.json")
DEFAULT_LOCK = Path("/run/lock/nuligahelper-deploy.lock")
APPROVED_DATABASE = Path("/var/lib/nuligahelper/nuliga_helper.db")
MIN_FREE_BYTES = 2 * 1024**3
MIN_AVAILABLE_MEMORY_KIB = 256 * 1024
RECORD_OUTCOMES = {"maintenance", "migration_required", "failed",
                   "web_ready", "public_ready", "accepted", "rolled_back"}
RECORD_CHECKS = {"fetched_master", "offline_suite", "syntax", "writers_quiescent",
                 "snapshot_validated", "schema_ready", "web_ready", "public_ready",
                 "timers_reviewed", "rollback_ready"}
TIMERS = ("nuligahelper-daily.timer", "nuligahelper-cleanup.timer",
          "nuligahelper-monitor.timer")
CALENDAR_TIMERS = TIMERS[:2]
DATABASE_USERS = ("nuligahelper-daily.service", "nuligahelper-cleanup.service",
                  "nuligahelper-monitor.service", "nuligahelper-preview.service",
                  "nuligahelper-preflight.service", "nuligahelper-launch-check.service",
                  "nuligahelper-alert-test.service")


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
    required = {"source", "app_root", "source_cache", "service_group", "python",
                "database", "recovery_dir", "public_health_url"}
    if not isinstance(data, dict) or set(data) != required:
        raise DeployError("deployment config keys are invalid")
    for key in required:
        if not isinstance(data[key], str) or not data[key]:
            raise DeployError(f"deployment config {key} is invalid")
    for key in ("app_root", "source_cache", "python", "database", "recovery_dir"):
        value = Path(data[key])
        if not value.is_absolute() or ".." in value.parts:
            raise DeployError(f"deployment config {key} must be absolute")
    if data["app_root"] != "/opt/nuligahelper":
        raise DeployError("app_root differs from the approved production root")
    if Path(data["source_cache"]).is_relative_to(data["app_root"]):
        raise DeployError("source cache must be outside the application tree")
    if data["database"] != str(APPROVED_DATABASE):
        raise DeployError("database differs from the approved production path")
    if Path(data["recovery_dir"]).is_relative_to(data["app_root"]) or \
            Path(data["recovery_dir"]).is_relative_to(APPROVED_DATABASE.parent):
        raise DeployError("recovery directory must be outside code and live state")
    if not re.fullmatch(r"[a-z_][a-z0-9_-]*", data["service_group"]):
        raise DeployError("service_group is invalid")
    if any(marker in data["source"].lower() for marker in ("@", "token", "password")) and \
            data["source"].startswith("https://"):
        raise DeployError("source URL must not contain an embedded credential")
    try:
        health = urlsplit(data["public_health_url"])
    except ValueError as exc:
        raise DeployError("public_health_url is invalid") from exc
    if health.scheme != "https" or not health.hostname or health.username or \
            health.password or health.path != "/healthz" or health.query or \
            health.fragment or health.hostname.endswith(".invalid"):
        raise DeployError("public_health_url must be an approved HTTPS health endpoint")
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


def prior_commit(previous: str) -> str:
    """Identify both a prepared release and the one-time legacy Git checkout."""
    path = Path(previous)
    if path.parent == Path("/opt/nuligahelper/releases") and SHA.fullmatch(path.name):
        return path.name
    if path.parent != Path("/opt/nuligahelper") or \
            not re.fullmatch(r"[0-9a-f]{7,40}", path.name):
        raise DeployError("previous release identity is unavailable")
    commit = run("git", "-C", previous, "rev-parse", "--verify", "HEAD^{commit}",
                 env=git_environment())
    if not SHA.fullmatch(commit) or not commit.startswith(path.name):
        raise DeployError("legacy release checkout does not match its directory")
    return commit


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


def _unit_state(unit: str) -> dict[str, str]:
    properties = ("LoadState", "ActiveState", "UnitFileState", "Result", "LastTriggerUSec")
    result = subprocess.run(("systemctl", "show", unit,
                             *(f"--property={item}" for item in properties)),
                            env=safe_environment(), text=True, capture_output=True,
                            timeout=15, check=False)
    if result.returncode:
        raise DeployError("systemd state is unavailable")
    state = dict(line.split("=", 1) for line in result.stdout.splitlines() if "=" in line)
    if state.get("LoadState") != "loaded":
        raise DeployError("required systemd unit is not loaded")
    return {item: state.get(item, "") for item in properties}


class HostController:
    """Bounded, non-interactive systemd and database-handle operations."""

    def state(self, unit: str) -> dict[str, str]:
        return _unit_state(unit)

    def systemctl(self, *arguments: str) -> None:
        result = subprocess.run(("systemctl", *arguments), env=safe_environment(),
                                text=True, capture_output=True, timeout=60, check=False)
        if result.returncode:
            raise DeployError("systemd operation failed")

    def database_handles(self, database: Path) -> bool:
        for path in (database, Path(str(database) + "-wal"),
                     Path(str(database) + "-shm")):
            if not path.exists():
                continue
            result = subprocess.run(("lsof", "-t", "--", str(path)),
                                    env=safe_environment(), capture_output=True,
                                    timeout=15, check=False)
            if result.returncode == 0:
                return True
            if result.returncode != 1:
                raise DeployError("database open-handle inspection failed")
        return False

    def loopback_listener(self) -> bool:
        result = subprocess.run(("ss", "-ltnH", "sport = :8080"),
                                env=safe_environment(), text=True,
                                capture_output=True, timeout=15, check=False)
        if result.returncode:
            raise DeployError("loopback listener inspection failed")
        listeners = [line.split() for line in result.stdout.splitlines()]
        if len(listeners) != 1 or len(listeners[0]) < 4:
            return False
        return listeners[0][0] == "LISTEN" and \
            listeners[0][3] == "127.0.0.1:8080"


def check_health(public_url: str, *, public: bool) -> None:
    """Check the exact readiness response without redirects or credential output."""
    parsed = urlsplit(public_url)
    if parsed.scheme != "https" or not parsed.hostname or parsed.path != "/healthz" or \
            parsed.username or parsed.password or parsed.query or parsed.fragment:
        raise DeployError("public health endpoint is invalid")
    if public:
        connection = http.client.HTTPSConnection(
            parsed.hostname, parsed.port or 443, timeout=8,
            context=ssl.create_default_context())
        headers = {"Host": parsed.hostname}
    else:
        connection = http.client.HTTPConnection("127.0.0.1", 8080, timeout=5)
        headers = {"Host": parsed.hostname, "X-Forwarded-For": "127.0.0.1",
                   "X-Forwarded-Proto": "https", "X-Forwarded-Host": parsed.hostname}
    try:
        connection.request("GET", "/healthz", headers=headers)
        response = connection.getresponse()
        body = response.read(64)
        if response.status != 200 or body != b"ok\n":
            raise DeployError("public health check failed" if public else
                              "local readiness check failed")
    except (OSError, http.client.HTTPException) as exc:
        raise DeployError("public health check failed" if public else
                          "local readiness check failed") from exc
    finally:
        connection.close()


def verify_local_web(config: dict[str, object], commit: str,
                     controller: HostController) -> None:
    app_root = Path(config["app_root"])
    if prior_release(app_root) != str(prepared_candidate(app_root, commit)):
        raise DeployError("current link does not identify the selected release")
    if controller.state("nuligahelper-web.service").get("ActiveState") != "active":
        raise DeployError("web service is not active")
    if not controller.loopback_listener():
        raise DeployError("web listener is not loopback-only")
    if schema_probe(prepared_candidate(app_root, commit), Path(config["database"]))["gate"] != "ready":
        raise DeployError("candidate database revision is not ready")
    check_health(str(config["public_health_url"]), public=False)


def quiesce_writers(database: Path, controller: HostController, *,
                    wait_seconds: float = 20 * 60,
                    clock=time.monotonic, sleep=time.sleep) -> dict[str, object]:
    """Pause timers, wait for jobs, close ingress/web, then check open handles.

    On any failure, keep stopped/disabled units paused for operator recovery.
    Never terminate an already-running oneshot notification or backup job.
    """
    before = {unit: controller.state(unit) for unit in
              (*TIMERS, "caddy.service", "nuligahelper-web.service")}
    for timer in TIMERS:
        controller.systemctl("disable", "--now", timer)
        if controller.state(timer).get("ActiveState") != "inactive":
            raise DeployError("timer did not stop cleanly")
    deadline = clock() + wait_seconds
    for service in DATABASE_USERS:
        while controller.state(service).get("ActiveState") in {"active", "activating"}:
            if clock() >= deadline:
                raise DeployError("database job is still active; timers remain paused")
            sleep(min(2.0, max(0.0, deadline - clock())))
    controller.systemctl("stop", "caddy.service")
    controller.systemctl("stop", "nuligahelper-web.service")
    if controller.state("caddy.service").get("ActiveState") != "inactive" or \
            controller.state("nuligahelper-web.service").get("ActiveState") != "inactive":
        raise DeployError("ingress or web service did not stop cleanly")
    for service in DATABASE_USERS:
        if controller.state(service).get("ActiveState") != "inactive":
            raise DeployError("database job started during maintenance entry")
    if controller.database_handles(database):
        raise DeployError("database still has an open handle; keep services paused")
    return {"units_before": before, "writers_quiescent": True}


def pending_calendar_catchup(timer_state: dict[str, str], *,
                             now: datetime | None = None) -> bool:
    """Conservatively flag a missed 09:00 Europe/Berlin persistent trigger.

    Unknown last-trigger text is treated as pending after 09:00. This is an
    operator decision aid, never permission to start a timer automatically.
    """
    local = (now or datetime.now(timezone.utc)).astimezone(ZoneInfo("Europe/Berlin"))
    today_fire = datetime.combine(local.date(), daytime(9), local.tzinfo)
    latest_fire = today_fire if local >= today_fire else today_fire - timedelta(days=1)
    last = timer_state.get("LastTriggerUSec", "")
    match = re.search(r"\b([0-9]{4}-[0-9]{2}-[0-9]{2}) ([0-9]{2}:[0-9]{2}:[0-9]{2})\b",
                      last)
    if not match:
        return True
    try:
        triggered = datetime.fromisoformat(f"{match[1]}T{match[2]}").replace(
            tzinfo=local.tzinfo)
    except ValueError:
        return True
    return triggered < latest_fire


def schema_probe(candidate: Path, database: Path) -> dict[str, str]:
    environment = safe_environment()
    environment["PYTHONPATH"] = str(candidate)
    output = run(str(candidate / "venv/bin/python"), "-B",
                 str(candidate / "release-assets/recovery_check.py"), "schema",
                 "--database", str(database), env=environment)
    try:
        result = json.loads(output)
    except ValueError as exc:
        raise DeployError("candidate schema probe returned invalid output") from exc
    if result.get("gate") not in {"ready", "migration_required", "refused"}:
        raise DeployError("candidate schema probe returned invalid state")
    return result


def snapshot_probe(candidate: Path, database: Path, recovery_dir: Path,
                   deployment_id: str) -> Path:
    """Invoke the candidate's WAL-safe snapshot helper after writers stop."""
    private_recovery_directory(recovery_dir)
    if not re.fullmatch(r"[0-9a-f]{32}", deployment_id):
        raise DeployError("deployment identifier is invalid")
    destination = recovery_dir / f"snapshot-{deployment_id}.db"
    environment = safe_environment()
    environment["PYTHONPATH"] = str(candidate)
    output = run(str(candidate / "venv/bin/python"), "-B",
                 str(candidate / "release-assets/recovery_check.py"), "snapshot",
                 "--database", str(database), "--destination", str(destination),
                 env=environment)
    try:
        result = json.loads(output)
    except ValueError as exc:
        raise DeployError("candidate snapshot probe returned invalid output") from exc
    if result != {"snapshot": str(destination), "validated": True}:
        raise DeployError("candidate snapshot probe returned invalid state")
    info = destination.stat()
    if info.st_uid != 0 or stat.S_IMODE(info.st_mode) != 0o600 or info.st_size == 0:
        raise DeployError("validated snapshot has unexpected ownership or mode")
    return destination


def inspect_plan(config: dict[str, object], commit: str) -> dict[str, object]:
    """Report non-sensitive state without stopping a service or changing data."""
    app_root = Path(config["app_root"])
    candidate = prepared_candidate(app_root, commit)
    database = Path(config["database"])
    if database.is_symlink() or not database.is_file() or database.stat().st_size == 0:
        raise DeployError("approved production database is unavailable")
    units = {unit: _unit_state(unit) for unit in (
        "nuligahelper-web.service", "nuligahelper-daily.service",
        "nuligahelper-cleanup.service", "nuligahelper-daily.timer",
        "nuligahelper-cleanup.timer", "nuligahelper-monitor.timer", "caddy.service")}
    return {"selected_commit": commit, "candidate": str(candidate),
            "current": prior_release(app_root), "database": str(database),
            "database_size_bytes": database.stat().st_size,
            "schema": schema_probe(candidate, database), "units": units,
            "action": "read_only_plan", "activation_available": False}


def private_recovery_directory(path: Path) -> None:
    if not path.is_absolute() or ".." in path.parts or path.is_symlink() or not path.is_dir():
        raise DeployError("root-private recovery directory is unavailable")
    info = path.stat()
    if info.st_uid != 0 or stat.S_IMODE(info.st_mode) != 0o700:
        raise DeployError("recovery directory must be root-owned mode 0700")


def write_deployment_record(recovery_dir: Path, record: dict[str, object]) -> Path:
    """Atomically persist an allowlisted, root-only operational record."""
    private_recovery_directory(recovery_dir)
    keys = {"schema_version", "deployment_id", "source_commit", "source_tree",
            "previous_release", "previous_commit", "timer_before",
            "schema_before", "schema_after", "snapshot_path",
            "checks", "outcome", "public_reopened", "timers_paused", "updated_at"}
    if set(record) != keys or record["schema_version"] != 1:
        raise DeployError("deployment record fields are invalid")
    identifier = record["deployment_id"]
    if not isinstance(identifier, str) or not re.fullmatch(r"[0-9a-f]{32}", identifier):
        raise DeployError("deployment record identifier is invalid")
    if not all(isinstance(record[key], str) and SHA.fullmatch(record[key])
               for key in ("source_commit", "source_tree", "previous_commit")):
        raise DeployError("deployment record source identity is invalid")
    previous = record["previous_release"]
    if not isinstance(previous, str) or ".." in Path(previous).parts or \
            not Path(previous).is_relative_to('/opt/nuligahelper'):
        raise DeployError("deployment record previous release is invalid")
    parts = Path(previous).relative_to('/opt/nuligahelper').parts
    if not ((len(parts) == 1 and re.fullmatch(r"[0-9a-f]{7,40}", parts[0])) or
            (len(parts) == 2 and parts[0] == "releases" and SHA.fullmatch(parts[1]))):
        raise DeployError("deployment record previous release is invalid")
    if not record["previous_commit"].startswith(parts[-1]):
        raise DeployError("deployment record previous commit does not match its path")
    timer_before = record["timer_before"]
    if timer_before is not None:
        if not isinstance(timer_before, dict) or set(timer_before) != set(TIMERS):
            raise DeployError("deployment record timer state is invalid")
        for state in timer_before.values():
            if not isinstance(state, dict) or set(state) != {
                    "ActiveState", "UnitFileState", "LastTriggerUSec"} or \
                    state["ActiveState"] not in {"active", "inactive", "failed"} or \
                    state["UnitFileState"] not in {"enabled", "disabled", "static"} or \
                    not isinstance(state["LastTriggerUSec"], str) or \
                    not re.fullmatch(
                        r"(?:|n/a|[A-Z][a-z]{2} [0-9]{4}-[0-9]{2}-[0-9]{2} "
                        r"[0-9]{2}:[0-9]{2}:[0-9]{2} [A-Z]{2,5})",
                                     state["LastTriggerUSec"]):
                raise DeployError("deployment record timer state is invalid")
    snapshot = record["snapshot_path"]
    if snapshot is not None and (not isinstance(snapshot, str) or
                                 Path(snapshot) != recovery_dir / f"snapshot-{identifier}.db"):
        raise DeployError("deployment record snapshot path is invalid")
    for field in ("schema_before", "schema_after"):
        value = record[field]
        if value is not None and (not isinstance(value, str) or
                                  not re.fullmatch(
                                      r"(?:[0-9]{4}_[a-z0-9_]{1,64}|[0-9a-f]{12,40})",
                                      value)):
            raise DeployError("deployment record schema identity is invalid")
    if not isinstance(record["outcome"], str) or \
            record["outcome"] not in RECORD_OUTCOMES or \
            type(record["public_reopened"]) is not bool or \
            record["timers_paused"] is not None and \
            type(record["timers_paused"]) is not bool:
        raise DeployError("deployment record state is invalid")
    checks = record["checks"]
    if not isinstance(checks, list) or any(not isinstance(item, str) or
           item not in RECORD_CHECKS for item in checks):
        raise DeployError("deployment record checks are invalid")
    stamp = record["updated_at"]
    if not isinstance(stamp, str) or not re.fullmatch(r"[0-9T:+.Z-]{20,40}", stamp):
        raise DeployError("deployment record timestamp is invalid")
    target = recovery_dir / f"deployment-{identifier}.json"
    descriptor, raw_temp = tempfile.mkstemp(prefix=".deployment-", dir=recovery_dir)
    temporary = Path(raw_temp)
    try:
        os.fchmod(descriptor, 0o600)
        with os.fdopen(descriptor, "w", encoding="utf-8") as handle:
            json.dump(record, handle, sort_keys=True)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, target)
        directory_fd = os.open(recovery_dir, os.O_RDONLY | os.O_DIRECTORY)
        try:
            os.fsync(directory_fd)
        finally:
            os.close(directory_fd)
        return target
    finally:
        temporary.unlink(missing_ok=True)


def enter_maintenance(config: dict[str, object], commit: str,
                      controller: HostController) -> dict[str, object]:
    """Quiesce, snapshot, and gate schema; never switch or reopen services.

    This operation is not yet exposed by the CLI. Once exposed, any failure
    after timer shutdown must leave maintenance in place for the operator.
    """
    app_root = Path(config["app_root"])
    candidate = prepared_candidate(app_root, commit)
    marker = json.loads((candidate / ".prepared.json").read_text(encoding="utf-8"))
    database = Path(config["database"])
    recovery_dir = Path(config["recovery_dir"])
    private_recovery_directory(recovery_dir)
    check_headroom(app_root)
    check_headroom(recovery_dir)
    previous = prior_release(app_root)
    if previous is None:
        raise DeployError("current release is unavailable")
    previous_sha = prior_commit(previous)
    before = schema_probe(candidate, database)
    if before["gate"] == "refused":
        raise DeployError("database schema is not safe for candidate activation")
    record = {"schema_version": 1, "deployment_id": uuid4().hex,
              "source_commit": commit, "source_tree": marker["tree"],
              "previous_release": previous, "previous_commit": previous_sha,
              "timer_before": None, "schema_before": before["revision"] or None,
              "schema_after": None, "snapshot_path": None,
              "checks": ["fetched_master", "offline_suite", "syntax"],
              "outcome": "maintenance", "public_reopened": False,
              "timers_paused": None,
              "updated_at": datetime.now(timezone.utc).isoformat()}
    write_deployment_record(recovery_dir, record)
    try:
        maintenance = quiesce_writers(database, controller)
        record["timer_before"] = {
            timer: {key: maintenance["units_before"][timer][key] for key in
                    ("ActiveState", "UnitFileState", "LastTriggerUSec")}
            for timer in TIMERS}
        record["timers_paused"] = True
        record["checks"].append("writers_quiescent")
        record["updated_at"] = datetime.now(timezone.utc).isoformat()
        write_deployment_record(recovery_dir, record)
        snapshot = snapshot_probe(candidate, database, recovery_dir,
                                  record["deployment_id"])
        record["snapshot_path"] = str(snapshot)
        record["checks"].append("snapshot_validated")
        record["updated_at"] = datetime.now(timezone.utc).isoformat()
        write_deployment_record(recovery_dir, record)
        after = schema_probe(candidate, database)
        record["schema_after"] = after["revision"] or None
        if after["gate"] == "migration_required":
            record["outcome"] = "migration_required"
        elif after["gate"] == "ready":
            record["checks"].append("schema_ready")
        else:
            raise DeployError("database schema changed unexpectedly during maintenance")
        record["updated_at"] = datetime.now(timezone.utc).isoformat()
        write_deployment_record(recovery_dir, record)
        return record
    except BaseException:
        record["outcome"] = "failed"
        record["updated_at"] = datetime.now(timezone.utc).isoformat()
        write_deployment_record(recovery_dir, record)
        raise


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
    inspection = subcommands.add_parser("inspect", help="read-only host and candidate plan")
    inspection.add_argument("--sha", help="prepared candidate SHA to inspect")
    args = parser.parse_args()
    if os.geteuid() != 0:
        parser.error("deployment command must run as root")
    try:
        config = load_config(args.config)
        if args.action == "inspect":
            if args.sha:
                report = inspect_plan(config, args.sha)
            else:
                report = {"app_root": config["app_root"],
                          "current": prior_release(Path(config["app_root"])),
                          "source_cache": config["source_cache"],
                          "activation_available": False}
            print(json.dumps(report, sort_keys=True))
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
