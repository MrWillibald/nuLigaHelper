#!/usr/bin/env python3
"""Root-operated, operator-invoked release preparation and guarded cutover.

Install a reviewed copy outside the service-readable application tree. Never
run activation until the host-specific units and rollback path are rehearsed.
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
DEFAULT_LOCK = Path("/run/nuligahelper-deploy/lock")
APPROVED_DATABASE = Path("/var/lib/nuligahelper/nuliga_helper.db")
APPROVED_SOURCES = {
    "git@github.com:MrWillibald/nuLigaHelper.git",
    "https://github.com/MrWillibald/nuLigaHelper.git",
}
MIN_FREE_BYTES = 2 * 1024**3
MIN_AVAILABLE_MEMORY_KIB = 256 * 1024
WEB_READY_TIMEOUT_SECONDS = 60
RECORD_OUTCOMES = {"maintenance", "migration_required", "failed",
                   "web_ready", "public_ready", "accepted", "rolled_back"}
RECORD_CHECKS = {"fetched_master", "offline_suite", "syntax", "writers_quiescent",
                 "snapshot_validated", "schema_ready", "web_ready", "public_ready",
                 "timers_reviewed", "rollback_ready", "current_switched"}
TIMERS = ("nuligahelper-daily.timer", "nuligahelper-cleanup.timer",
          "nuligahelper-monitor.timer")
CALENDAR_TIMERS = TIMERS[:2]
DATABASE_USERS = ("nuligahelper-daily.service", "nuligahelper-cleanup.service",
                  "nuligahelper-monitor.service", "nuligahelper-preview.service",
                  "nuligahelper-preflight.service", "nuligahelper-launch-check.service",
                  "nuligahelper-alert-test.service")
SERVICE_UNITS = ("nuligahelper-web.service", *DATABASE_USERS,
                 "nuligahelper-alert@.service")


class DeployError(RuntimeError):
    """A non-secret, operator-actionable deployment refusal."""


def run(*args: str, cwd: Path | None = None, env: dict[str, str] | None = None) -> str:
    try:
        result = subprocess.run(args, cwd=cwd, env=env, text=True, stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE, timeout=30 * 60, check=False)
    except subprocess.TimeoutExpired as exc:
        raise DeployError(f"command timed out: {Path(args[0]).name}") from exc
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
    if data["source"] not in APPROVED_SOURCES:
        raise DeployError("source differs from the approved GitHub repository")
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
    if not path.is_absolute() or ".." in path.parts or path.parent.is_symlink():
        raise DeployError("host lock path is invalid")
    path.parent.mkdir(mode=0o700, exist_ok=True)
    folder = path.parent.stat()
    if folder.st_uid != os.geteuid() or stat.S_IMODE(folder.st_mode) != 0o700:
        raise DeployError("host lock directory is not private")
    descriptor = os.open(path, os.O_CREAT | os.O_RDWR | os.O_CLOEXEC | os.O_NOFOLLOW,
                         0o600)
    try:
        info = os.fstat(descriptor)
        if not stat.S_ISREG(info.st_mode) or info.st_uid != os.geteuid() or \
                stat.S_IMODE(info.st_mode) != 0o600:
            raise DeployError("host lock file is not private")
        fcntl.flock(descriptor, fcntl.LOCK_EX | fcntl.LOCK_NB)
    except BlockingIOError as exc:
        os.close(descriptor)
        raise DeployError("another deployment attempt holds the host lock") from exc
    except BaseException:
        os.close(descriptor)
        raise
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


def check_snapshot_headroom(recovery_dir: Path, database: Path) -> None:
    if database.is_symlink() or not database.is_file() or database.stat().st_size == 0:
        raise DeployError("approved production database is unavailable")
    if shutil.disk_usage(recovery_dir).free < max(
            MIN_FREE_BYTES, database.stat().st_size * 2 + 512 * 1024**2):
        raise DeployError("insufficient free disk space for a durable database snapshot")


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
    try:
        result = subprocess.run(("systemctl", "show", unit,
                                 *(f"--property={item}" for item in properties)),
                                env=safe_environment(), text=True, capture_output=True,
                                timeout=15, check=False)
    except subprocess.TimeoutExpired as exc:
        raise DeployError("systemd state inspection timed out") from exc
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
        try:
            result = subprocess.run(("systemctl", *arguments), env=safe_environment(),
                                    text=True, capture_output=True, timeout=60, check=False)
        except subprocess.TimeoutExpired as exc:
            raise DeployError("systemd operation timed out") from exc
        if result.returncode:
            raise DeployError("systemd operation failed")

    def database_handles(self, database: Path) -> bool:
        for path in (database, Path(str(database) + "-wal"),
                     Path(str(database) + "-shm")):
            if not path.exists():
                continue
            try:
                result = subprocess.run(("lsof", "-t", "--", str(path)),
                                        env=safe_environment(), capture_output=True,
                                        timeout=15, check=False)
            except subprocess.TimeoutExpired as exc:
                raise DeployError("database open-handle inspection timed out") from exc
            if result.returncode == 0:
                return True
            if result.returncode != 1:
                raise DeployError("database open-handle inspection failed")
        return False

    def loopback_listener(self) -> bool:
        try:
            result = subprocess.run(("ss", "-ltnH", "sport = :8080"),
                                    env=safe_environment(), text=True,
                                    capture_output=True, timeout=15, check=False)
        except subprocess.TimeoutExpired as exc:
            raise DeployError("loopback listener inspection timed out") from exc
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


def wait_for_web_ready(config: dict[str, object], controller: HostController, *,
                       wait_seconds: float = WEB_READY_TIMEOUT_SECONDS,
                       clock=time.monotonic, sleep=time.sleep) -> None:
    """Wait for Type=simple Gunicorn readiness before allowing public ingress."""
    deadline = clock() + wait_seconds
    while True:
        state = controller.state("nuligahelper-web.service").get("ActiveState")
        if state in {"failed", "inactive"}:
            raise DeployError("web service failed before becoming ready")
        if state == "active" and controller.loopback_listener():
            try:
                check_health(str(config["public_health_url"]), public=False)
            except DeployError as exc:
                if str(exc) != "local readiness check failed":
                    raise
            else:
                if clock() <= deadline:
                    return
        remaining = deadline - clock()
        if remaining <= 0:
            raise DeployError("web readiness timed out with public ingress closed")
        sleep(min(1.0, remaining))


def verify_local_web(config: dict[str, object], commit: str,
                     controller: HostController) -> None:
    app_root = Path(config["app_root"])
    if prior_release(app_root) != str(prepared_candidate(app_root, commit)):
        raise DeployError("current link does not identify the selected release")
    if schema_probe(prepared_candidate(app_root, commit), Path(config["database"]))["gate"] != "ready":
        raise DeployError("candidate database revision is not ready")
    wait_for_web_ready(config, controller)


def unit_paths_safe(content: str, name: str) -> None:
    lines = [line.strip() for line in content.splitlines()
             if line.startswith(("WorkingDirectory=", "EnvironmentFile=",
                                 "Exec"))]
    if "WorkingDirectory=/opt/nuligahelper/current" not in lines or \
            "EnvironmentFile=/etc/nuligahelper/web.env" not in lines:
        raise DeployError("installed application unit has mixed runtime paths")
    commands = [line for line in lines if line.startswith("Exec")]
    if not any(line.startswith("ExecStart=") for line in commands) or \
            any(not line.startswith(("ExecStart=", "ExecStartPre=")) or
                           "/opt/nuligahelper/current/" not in line or
                           "/opt/nuligahelper/" in line.replace(
                               "/opt/nuligahelper/current/", "")
                           for line in commands):
        raise DeployError("installed application unit has mixed runtime paths")
    if name == "nuligahelper-web.service" and not any(
            "/opt/nuligahelper/current/release-assets/gunicorn.conf.py" in line
            for line in commands):
        raise DeployError("installed web unit still uses legacy Gunicorn configuration")


def verify_installed_units(directory: Path = Path("/etc/systemd/system")) -> None:
    """Refuse activation until every code-running unit uses one release link."""
    for name in SERVICE_UNITS:
        unit = directory / name
        if unit.is_symlink():
            raise DeployError("installed application unit must not be a symlink")
        try:
            info = unit.stat()
            content = unit.read_text(encoding="utf-8")
        except OSError as exc:
            raise DeployError("installed application unit is unavailable") from exc
        if info.st_uid != 0 or info.st_mode & 0o022:
            raise DeployError("installed application unit is writable outside root")
        unit_paths_safe(content, name)
        loaded_name = ("nuligahelper-alert@monitor.service" if name ==
                       "nuligahelper-alert@.service" else name)
        loaded = run("systemctl", "show", loaded_name,
                     "--property=WorkingDirectory,EnvironmentFiles,ExecStart,ExecStartPre,"
                     "ExecCondition,ExecStartPost,ExecReload,ExecStop,ExecStopPost",
                     env=safe_environment())
        properties = dict(line.split("=", 1) for line in loaded.splitlines() if "=" in line)
        if properties.get("WorkingDirectory") != "/opt/nuligahelper/current" or \
                not re.fullmatch(
                    r"/etc/nuligahelper/web\.env \(ignore_errors=(?:yes|no)\)",
                    properties.get("EnvironmentFiles", "")):
            raise DeployError("loaded application unit has mixed runtime paths")
        for field in ("ExecStart", "ExecStartPre"):
            command = properties.get(field, "")
            if field == "ExecStart" and "/opt/nuligahelper/current/" not in command:
                raise DeployError("loaded application unit has mixed runtime paths")
            if "/opt/nuligahelper/" in command.replace(
                    "/opt/nuligahelper/current/", ""):
                raise DeployError("loaded application unit has mixed runtime paths")
        if any(properties.get(field, "") for field in
               ("ExecCondition", "ExecStartPost", "ExecReload", "ExecStop", "ExecStopPost")):
            raise DeployError("loaded application unit has unreviewed command hooks")
    run("systemd-analyze", "verify", *(str(directory / name) for name in SERVICE_UNITS),
        env=safe_environment())


def assert_maintenance_still_closed(database: Path, controller: HostController) -> None:
    for unit in (*TIMERS, "caddy.service", "nuligahelper-web.service", *DATABASE_USERS):
        if controller.state(unit).get("ActiveState") != "inactive":
            raise DeployError("maintenance state changed; keep writers and ingress stopped")
    if controller.database_handles(database):
        raise DeployError("database is open during maintenance")


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


def validate_recovery_snapshot(candidate: Path, database: Path,
                               recovery_dir: Path, record: dict[str, object]) -> None:
    identifier = record["deployment_id"]
    destination = recovery_dir / f"snapshot-{identifier}.db"
    if record["snapshot_path"] != str(destination):
        raise DeployError("deployment recovery snapshot is unavailable")
    environment = safe_environment()
    environment["PYTHONPATH"] = str(candidate)
    output = run(str(candidate / "venv/bin/python"), "-B",
                 str(candidate / "release-assets/recovery_check.py"), "validate",
                 "--database", str(database), "--destination", str(destination),
                 env=environment)
    try:
        result = json.loads(output)
    except ValueError as exc:
        raise DeployError("candidate snapshot validation returned invalid output") from exc
    if result != {"snapshot": str(destination), "validated": True}:
        raise DeployError("candidate snapshot validation returned invalid state")


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
            "action": "read_only_plan", "activation_available": True}


def private_recovery_directory(path: Path) -> None:
    if not path.is_absolute() or ".." in path.parts or path.is_symlink() or not path.is_dir():
        raise DeployError("root-private recovery directory is unavailable")
    info = path.stat()
    if info.st_uid != 0 or stat.S_IMODE(info.st_mode) != 0o700:
        raise DeployError("recovery directory must be root-owned mode 0700")


def validate_deployment_record(recovery_dir: Path, record: dict[str, object]) -> None:
    """Reject unapproved fields or values before writing or consuming state."""
    keys = {"schema_version", "deployment_id", "source_commit", "source_tree",
            "previous_release", "previous_commit", "timer_before",
            "schema_before", "schema_after", "snapshot_path",
            "checks", "outcome", "public_reopened", "timers_paused",
            "timer_decision", "pending_catchup", "updated_at"}
    if set(record) != keys or type(record["schema_version"]) is not int or \
            record["schema_version"] != 1:
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
                    not isinstance(state["ActiveState"], str) or \
                    state["ActiveState"] not in {"active", "inactive", "failed"} or \
                    not isinstance(state["UnitFileState"], str) or \
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
    if record["timer_decision"] is not None and \
            (not isinstance(record["timer_decision"], str) or
             record["timer_decision"] not in {"run", "hold"}):
        raise DeployError("deployment record timer decision is invalid")
    pending = record["pending_catchup"]
    if not isinstance(pending, list) or any(
            not isinstance(item, str) or item not in CALENDAR_TIMERS
            for item in pending) or len(set(pending)) != len(pending):
        raise DeployError("deployment record catch-up state is invalid")
    checks = record["checks"]
    if not isinstance(checks, list) or any(not isinstance(item, str) or
           item not in RECORD_CHECKS for item in checks):
        raise DeployError("deployment record checks are invalid")
    stamp = record["updated_at"]
    if not isinstance(stamp, str) or not re.fullmatch(r"[0-9T:+.Z-]{20,40}", stamp):
        raise DeployError("deployment record timestamp is invalid")


def write_deployment_record(recovery_dir: Path, record: dict[str, object]) -> Path:
    """Atomically persist an allowlisted, root-only operational record."""
    private_recovery_directory(recovery_dir)
    validate_deployment_record(recovery_dir, record)
    identifier = record["deployment_id"]
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


def read_deployment_record(recovery_dir: Path, deployment_id: str) -> dict[str, object]:
    private_recovery_directory(recovery_dir)
    if not re.fullmatch(r"[0-9a-f]{32}", deployment_id):
        raise DeployError("deployment identifier is invalid")
    path = recovery_dir / f"deployment-{deployment_id}.json"
    if path.is_symlink():
        raise DeployError("deployment record must not be a symlink")
    try:
        info = path.stat()
        if info.st_uid != 0 or stat.S_IMODE(info.st_mode) != 0o600:
            raise DeployError("deployment record ownership or mode is invalid")
        record = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise DeployError("deployment record is unavailable or invalid") from exc
    if not isinstance(record, dict):
        raise DeployError("deployment record is invalid")
    validate_deployment_record(recovery_dir, record)
    if record["deployment_id"] != deployment_id:
        raise DeployError("deployment record identity mismatch")
    return record


def enter_maintenance(config: dict[str, object], commit: str,
                      controller: HostController) -> dict[str, object]:
    """Quiesce, snapshot, and gate schema; never switch or reopen services.

    Any failure after timer shutdown leaves maintenance in place for recovery.
    """
    app_root = Path(config["app_root"])
    candidate = prepared_candidate(app_root, commit)
    marker = json.loads((candidate / ".prepared.json").read_text(encoding="utf-8"))
    database = Path(config["database"])
    recovery_dir = Path(config["recovery_dir"])
    private_recovery_directory(recovery_dir)
    check_headroom(app_root)
    check_headroom(recovery_dir)
    check_snapshot_headroom(recovery_dir, database)
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
              "timers_paused": None, "timer_decision": None,
              "pending_catchup": [],
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


def verify_previous_web_compatibility(previous: str, group_name: str) -> None:
    """The new web unit must also start the retained prior release on rollback."""
    gunicorn = Path(previous) / "release-assets/gunicorn.conf.py"
    if gunicorn.is_symlink() or not gunicorn.is_file():
        raise DeployError("previous release lacks the new web unit's Gunicorn file")
    info = gunicorn.stat()
    group = grp.getgrnam(group_name).gr_gid
    if info.st_uid != 0 or info.st_gid != group or \
            not info.st_mode & stat.S_IRGRP or info.st_mode & 0o022:
        raise DeployError("previous Gunicorn file is not safely service-readable")


def complete_activation(config: dict[str, object], record: dict[str, object],
                        controller: HostController) -> dict[str, object]:
    """Switch only after the persisted snapshot and stopped-writer gates pass."""
    app_root = Path(config["app_root"])
    database = Path(config["database"])
    recovery_dir = Path(config["recovery_dir"])
    validate_deployment_record(recovery_dir, record)
    if record["outcome"] not in {"maintenance", "migration_required"} or \
            record["public_reopened"] or record["timers_paused"] is not True:
        raise DeployError("deployment record is not ready for activation")
    commit = record["source_commit"]
    candidate = prepared_candidate(app_root, commit)
    marker = json.loads((candidate / ".prepared.json").read_text(encoding="utf-8"))
    if record["source_tree"] != marker["tree"] or \
            prior_release(app_root) != record["previous_release"]:
        raise DeployError("candidate or previous release changed during maintenance")
    assert_maintenance_still_closed(database, controller)
    validate_recovery_snapshot(candidate, database, recovery_dir, record)
    gate = schema_probe(candidate, database)
    if gate["gate"] != "ready":
        raise DeployError("candidate database revision is not ready; migration remains manual")
    record["schema_after"] = gate["revision"]
    if "schema_ready" not in record["checks"]:
        record["checks"].append("schema_ready")
    record["updated_at"] = datetime.now(timezone.utc).isoformat()
    write_deployment_record(recovery_dir, record)
    try:
        switch_current(app_root, commit)
        record["checks"].append("current_switched")
        record["updated_at"] = datetime.now(timezone.utc).isoformat()
        write_deployment_record(recovery_dir, record)
        controller.systemctl("start", "nuligahelper-web.service")
        verify_local_web(config, commit, controller)
        record["outcome"] = "web_ready"
        record["checks"].append("web_ready")
        record["updated_at"] = datetime.now(timezone.utc).isoformat()
        write_deployment_record(recovery_dir, record)
        # Persist this *before* opening ingress: a crash after start may have
        # accepted writes even if the public health probe never returned.
        record["public_reopened"] = True
        record["updated_at"] = datetime.now(timezone.utc).isoformat()
        write_deployment_record(recovery_dir, record)
        controller.systemctl("start", "caddy.service")
        if controller.state("caddy.service").get("ActiveState") != "active":
            raise DeployError("public ingress did not become active")
        check_health(str(config["public_health_url"]), public=True)
        record["outcome"] = "public_ready"
        record["checks"].append("public_ready")
        record["updated_at"] = datetime.now(timezone.utc).isoformat()
        write_deployment_record(recovery_dir, record)
        return record
    except BaseException:
        if record["public_reopened"]:
            try:
                controller.systemctl("stop", "caddy.service")
            except DeployError:
                pass  # The record remains conservative: public writes may exist.
        record["outcome"] = "failed"
        record["updated_at"] = datetime.now(timezone.utc).isoformat()
        write_deployment_record(recovery_dir, record)
        raise


def resume_timers(config: dict[str, object], record: dict[str, object],
                  controller: HostController, decision: str, *,
                  now: datetime | None = None) -> dict[str, object]:
    """Restore prior timer states only after an explicit catch-up decision."""
    recovery_dir = Path(config["recovery_dir"])
    validate_deployment_record(recovery_dir, record)
    if record["outcome"] not in {"public_ready", "rolled_back"} or \
            record["timers_paused"] is not True or not record["public_reopened"] or \
            record["timer_before"] is None or decision not in {"run", "hold"}:
        raise DeployError("timers cannot resume from this deployment state")
    if controller.state("caddy.service").get("ActiveState") != "active" or \
            controller.state("nuligahelper-web.service").get("ActiveState") != "active" or \
            not controller.loopback_listener():
        raise DeployError("web or public ingress is not ready for scheduled work")
    check_health(str(config["public_health_url"]), public=False)
    check_health(str(config["public_health_url"]), public=True)
    for timer in TIMERS:
        if controller.state(timer).get("ActiveState") != "inactive":
            raise DeployError("a paused timer was restarted outside deployment control")
    timer_before = record["timer_before"]
    pending = [timer for timer in CALENDAR_TIMERS
               if timer_before[timer]["ActiveState"] == "active" and
               timer_before[timer]["UnitFileState"] == "enabled" and
               pending_calendar_catchup(timer_before[timer], now=now)]
    record["pending_catchup"] = pending
    record["timer_decision"] = decision
    if "timers_reviewed" not in record["checks"]:
        record["checks"].append("timers_reviewed")
    record["updated_at"] = datetime.now(timezone.utc).isoformat()
    write_deployment_record(recovery_dir, record)
    if decision == "hold":
        return record
    try:
        for timer in TIMERS:
            before = timer_before[timer]
            if before["UnitFileState"] != "enabled":
                continue
            if before["ActiveState"] == "active":
                controller.systemctl("enable", "--now", timer)
            else:
                controller.systemctl("enable", timer)
            state = controller.state(timer)
            if state.get("UnitFileState") != "enabled" or \
                    state.get("ActiveState") != before["ActiveState"]:
                raise DeployError("timer did not return to its prior state")
        record["timers_paused"] = False
        if record["outcome"] != "rolled_back":
            record["outcome"] = "accepted"
        record["updated_at"] = datetime.now(timezone.utc).isoformat()
        write_deployment_record(recovery_dir, record)
        return record
    except BaseException:
        record["timers_paused"] = None  # A partial enable may already have fired.
        record["outcome"] = "failed"
        record["updated_at"] = datetime.now(timezone.utc).isoformat()
        write_deployment_record(recovery_dir, record)
        raise


def previous_schema_head(previous: str) -> str:
    environment = safe_environment()
    environment["PYTHONPATH"] = previous
    revision = run(str(Path(previous) / "venv/bin/python"), "-B", "-c",
                   "import schema_migrations; "
                   "heads = schema_migrations.head_revisions(); "
                   "assert len(heads) == 1; print(heads[0])",
                   env=environment)
    if not re.fullmatch(r"[0-9]{4}_[a-z0-9_]{1,64}", revision):
        raise DeployError("previous release schema head is invalid")
    return revision


def rollback_code_only(config: dict[str, object], record: dict[str, object],
                       controller: HostController) -> dict[str, object]:
    """Restore prior code only before ingress and only at the same schema."""
    recovery_dir = Path(config["recovery_dir"])
    validate_deployment_record(recovery_dir, record)
    if record["public_reopened"] or record["timers_paused"] is not True:
        raise DeployError("post-traffic or unpaused rollback requires manual recovery")
    app_root = Path(config["app_root"])
    database = Path(config["database"])
    commit = record["source_commit"]
    candidate = prepared_candidate(app_root, commit)
    if prior_release(app_root) != str(candidate):
        raise DeployError("selected release is not active; no code switch to undo")
    if record["schema_before"] is None or \
            record["schema_before"] != record["schema_after"]:
        raise DeployError("database schema changed; code-only rollback is unsafe")
    if controller.state("caddy.service").get("ActiveState") != "inactive":
        raise DeployError("public ingress is not closed")
    controller.systemctl("stop", "nuligahelper-web.service")
    assert_maintenance_still_closed(database, controller)
    gate = schema_probe(candidate, database)
    if gate["gate"] != "ready" or gate["revision"] != record["schema_before"] or \
            previous_schema_head(record["previous_release"]) != gate["revision"]:
        raise DeployError("previous release cannot use the current database revision")
    verify_previous_web_compatibility(record["previous_release"],
                                      str(config["service_group"]))
    atomic_current_link(app_root, Path(record["previous_release"]))
    try:
        controller.systemctl("start", "nuligahelper-web.service")
        wait_for_web_ready(config, controller)
        record["public_reopened"] = True  # Durable before any accepted public write.
        record["updated_at"] = datetime.now(timezone.utc).isoformat()
        write_deployment_record(recovery_dir, record)
        controller.systemctl("start", "caddy.service")
        if controller.state("caddy.service").get("ActiveState") != "active":
            raise DeployError("public ingress did not restart")
        check_health(str(config["public_health_url"]), public=True)
        record["outcome"] = "rolled_back"
        record["checks"].append("rollback_ready")
        record["updated_at"] = datetime.now(timezone.utc).isoformat()
        write_deployment_record(recovery_dir, record)
        return record
    except BaseException:
        if record["public_reopened"]:
            try:
                controller.systemctl("stop", "caddy.service")
            except DeployError:
                pass
        record["outcome"] = "failed"
        record["updated_at"] = datetime.now(timezone.utc).isoformat()
        write_deployment_record(recovery_dir, record)
        raise


def atomic_current_link(app_root: Path, target: Path) -> tuple[str, str]:
    """Replace the internal link as one namespace operation, then sync it."""
    if target.is_symlink() or not target.is_dir() or \
            not target.resolve().is_relative_to(app_root.resolve()):
        raise DeployError("release link target is invalid")
    previous = prior_release(app_root)
    if previous is None or previous == str(target):
        raise DeployError("current release is missing or already selected")
    temporary = app_root / (".current-next-" + uuid4().hex)
    try:
        os.symlink(target, temporary)
        os.replace(temporary, app_root / "current")
        directory_fd = os.open(app_root, os.O_RDONLY | os.O_DIRECTORY)
        try:
            os.fsync(directory_fd)
        finally:
            os.close(directory_fd)
    finally:
        temporary.unlink(missing_ok=True)
    return previous, str(target)


def switch_current(app_root: Path, commit: str) -> tuple[str, str]:
    """Select only a prepared complete candidate; caller enforces cutover gates."""
    return atomic_current_link(app_root, prepared_candidate(app_root, commit))


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
    activation = subcommands.add_parser("activate", help="stop writers and activate a prepared SHA")
    activation.add_argument("--sha", required=True, help="prepared full commit SHA")
    continuation = subcommands.add_parser("continue", help="continue after explicit schema migration")
    continuation.add_argument("--deployment-id", required=True)
    resumption = subcommands.add_parser("resume-timers", help="record catch-up decision and restore timers")
    resumption.add_argument("--deployment-id", required=True)
    resumption.add_argument("--catchup", required=True, choices=("run", "hold"))
    rollback = subcommands.add_parser("rollback-code", help="guarded pre-traffic code-only rollback")
    rollback.add_argument("--deployment-id", required=True)
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
                          "activation_available": True}
            print(json.dumps(report, sort_keys=True))
            return 0
        lock_fd = exclusive_lock(args.lock)
        try:
            controller = HostController()
            if args.action == "prepare":
                result = prepare(config, args.sha)
            elif args.action == "activate":
                verify_installed_units()
                candidate = prepared_candidate(Path(config["app_root"]), args.sha)
                marker = json.loads((candidate / ".prepared.json").read_text(encoding="utf-8"))
                previous = prior_release(Path(config["app_root"]))
                if previous is None or previous != marker.get("previous_release"):
                    raise DeployError("current release changed since candidate preparation")
                verify_previous_web_compatibility(previous, str(config["service_group"]))
                result = enter_maintenance(config, args.sha, controller)
                if result["outcome"] == "migration_required":
                    print(json.dumps(result, sort_keys=True))
                    return 2
                result = complete_activation(config, result, controller)
            else:
                recovery_dir = Path(config["recovery_dir"])
                result = read_deployment_record(recovery_dir, args.deployment_id)
                if args.action == "continue":
                    if result["outcome"] != "migration_required":
                        raise DeployError("only a migration-paused deployment can continue")
                    verify_installed_units()
                    verify_previous_web_compatibility(
                        result["previous_release"], str(config["service_group"]))
                    result = complete_activation(config, result, controller)
                elif args.action == "resume-timers":
                    result = resume_timers(config, result, controller, args.catchup)
                elif args.action == "rollback-code":
                    result = rollback_code_only(config, result, controller)
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
