"""Persistent, privacy-safe authentication abuse controls."""

from __future__ import annotations

import hashlib
import hmac
import ipaddress
import logging
import time
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Iterable, Mapping

from sqlalchemy import delete, select
from sqlalchemy.exc import SQLAlchemyError

import db

LOGGER = logging.getLogger("nuligahelper.security")


@dataclass(frozen=True)
class Policy:
    action: str
    dimension: str
    channel: str
    limit: int
    window_seconds: int


@dataclass(frozen=True)
class Config:
    policies: Mapping[str, Policy]
    retention_seconds: int
    cleanup_batch: int
    cleanup_interval_seconds: int
    trusted_proxies: tuple[ipaddress.IPv4Network | ipaddress.IPv6Network, ...]
    trusted_hops: int
    proxy_error: str


@dataclass(frozen=True)
class Rule:
    policy: str
    subject: str
    prehashed: bool = False


@dataclass(frozen=True)
class Decision:
    allowed: bool
    reason_dimensions: tuple[str, ...] = ()
    storage_error: bool = False


_POLICY_DEFAULTS = {
    "login_client": ("login_request", "client", "any", 10, 900),
    "login_contact_email": ("login_request", "contact", "email", 3, 900),
    "login_contact_sms": ("login_request", "contact", "sms", 2, 900),
    "login_person_email": ("login_request", "person", "email", 3, 900),
    "login_person_sms": ("login_request", "person", "sms", 2, 900),
    "registration_client": ("registration_request", "client", "any", 5, 900),
    "registration_contact_email": ("registration_request", "contact", "email", 3, 900),
    "registration_contact_sms": ("registration_request", "contact", "sms", 2, 900),
    "registration_person_email": ("registration_request", "person", "email", 3, 900),
    "registration_person_sms": ("registration_request", "person", "sms", 2, 900),
    "confirmation_client": ("code_confirmation", "client", "any", 10, 900),
    "confirmation_contact": ("code_confirmation", "contact", "any", 5, 900),
    "confirmation_person": ("code_confirmation", "person", "any", 5, 900),
    "sms_contact_cap": ("sms_delivery", "contact", "sms", 6, 86400),
    "sms_person_cap": ("sms_delivery", "person", "sms", 6, 86400),
    "sms_global_cap": ("sms_delivery", "global", "sms", 30, 86400),
}


def load_config(section: Mapping | None = None) -> Config:
    """Merge and validate the optional ``auth_abuse`` configuration section."""
    raw = dict(section or {})
    unknown = set(raw) - {
        "policies", "retention_seconds", "cleanup_batch",
        "cleanup_interval_seconds", "trusted_proxies", "trusted_hops",
        "proxy_error",
    }
    if unknown:
        raise ValueError(f"Unknown auth_abuse settings: {', '.join(sorted(unknown))}")
    overrides = raw.get("policies", {})
    if not isinstance(overrides, Mapping):
        raise ValueError("auth_abuse.policies must be an object")
    unknown_policies = set(overrides) - set(_POLICY_DEFAULTS)
    if unknown_policies:
        raise ValueError(
            f"Unknown auth_abuse policies: {', '.join(sorted(unknown_policies))}"
        )
    policies = {}
    for name, values in _POLICY_DEFAULTS.items():
        action, dimension, channel, default_limit, default_window = values
        override = overrides.get(name, {})
        if not isinstance(override, Mapping) or set(override) - {"limit", "window_seconds"}:
            raise ValueError(f"Invalid auth_abuse policy {name}")
        limit = override.get("limit", default_limit)
        window = override.get("window_seconds", default_window)
        if type(limit) is not int or type(window) is not int or limit <= 0 or window <= 0:
            raise ValueError(f"Policy {name} limit and window_seconds must be positive integers")
        policies[name] = Policy(action, dimension, channel, limit, window)
    retention = raw.get("retention_seconds", 604800)
    batch = raw.get("cleanup_batch", 200)
    interval = raw.get("cleanup_interval_seconds", 300)
    hops = raw.get("trusted_hops", 0)
    proxy_error = raw.get("proxy_error", "fallback")
    for key, value in (("retention_seconds", retention), ("cleanup_batch", batch),
                       ("cleanup_interval_seconds", interval)):
        if type(value) is not int or value <= 0:
            raise ValueError(f"auth_abuse.{key} must be a positive integer")
    if retention < max(policy.window_seconds for policy in policies.values()):
        raise ValueError("auth_abuse.retention_seconds must cover the longest policy window")
    if type(hops) is not int or hops < 0:
        raise ValueError("auth_abuse.trusted_hops must be a non-negative integer")
    if proxy_error not in {"fallback", "refuse"}:
        raise ValueError("auth_abuse.proxy_error must be fallback or refuse")
    networks = []
    raw_proxies = raw.get("trusted_proxies", [])
    if not isinstance(raw_proxies, list):
        raise ValueError("auth_abuse.trusted_proxies must be an array")
    for value in raw_proxies:
        try:
            networks.append(ipaddress.ip_network(value, strict=False))
        except (TypeError, ValueError) as exc:
            raise ValueError(f"Invalid trusted proxy address/CIDR: {value!r}") from exc
    if bool(networks) != bool(hops):
        raise ValueError("trusted_proxies and trusted_hops must be enabled together")
    return Config(policies, retention, batch, interval, tuple(networks), hops, proxy_error)


def canonical_ip(value: str) -> str:
    return ipaddress.ip_address(value).compressed


def subject_digest(secret: str, dimension: str, channel: str, subject: str) -> str:
    if dimension not in {"client", "contact", "person", "global"}:
        raise ValueError("Unknown limiter dimension")
    domain = f"nuligahelper-auth-abuse:v1:{dimension}:{channel}\0".encode()
    return hmac.new(secret.encode(), domain + subject.encode(), hashlib.sha256).hexdigest()


def attributed_client(peer: str | None, x_forwarded_for: str | None, config: Config) -> str | None:
    """Return a canonical client, trusting an exact configured proxy chain only."""
    try:
        direct = canonical_ip(peer or "")
    except ValueError:
        return None
    if not config.trusted_proxies:
        return direct
    direct_ip = ipaddress.ip_address(direct)
    if not any(direct_ip in network for network in config.trusted_proxies):
        return direct
    try:
        values = [canonical_ip(value.strip()) for value in (x_forwarded_for or "").split(",")]
        if len(values) != config.trusted_hops or not all(values):
            raise ValueError
        for proxy in values[1:]:
            proxy_ip = ipaddress.ip_address(proxy)
            if not any(proxy_ip in network for network in config.trusted_proxies):
                raise ValueError
    except ValueError:
        return direct if config.proxy_error == "fallback" else None
    return values[0]


class Service:
    def __init__(self, engine, config: Config, secret: str):
        self.engine = engine
        self.config = config
        self.secret = secret
        self._last_cleanup = 0.0

    def digest(self, dimension: str, channel: str, subject: str) -> str:
        if dimension == "client":
            subject = canonical_ip(subject)
        elif dimension == "global":
            subject = "application"
        return subject_digest(self.secret, dimension, channel, subject)

    def reserve(self, rules: Iterable[Rule], now: datetime | None = None) -> Decision:
        now = (now or datetime.now(timezone.utc)).replace(tzinfo=None)
        prepared = []
        for rule in rules:
            policy = self.config.policies[rule.policy]
            digest = rule.subject if rule.prehashed else self.digest(
                policy.dimension, policy.channel, rule.subject
            )
            start_epoch = int(now.replace(tzinfo=timezone.utc).timestamp())
            start_epoch -= start_epoch % policy.window_seconds
            started = datetime.fromtimestamp(start_epoch, timezone.utc).replace(tzinfo=None)
            prepared.append((policy, digest, started))
        # Duplicate semantic rules must consume only one allowance per action.
        prepared = list(dict.fromkeys(prepared))
        table = db.AuthAbuseCounter.__table__
        try:
            with self.engine.connect() as connection:
                connection.exec_driver_sql("BEGIN IMMEDIATE")
                rows = []
                exhausted = set()
                for policy, digest, started in prepared:
                    row = connection.execute(select(table).where(
                        table.c.action == policy.action,
                        table.c.dimension == policy.dimension,
                        table.c.subject_digest == digest,
                        table.c.channel == policy.channel,
                        table.c.window_started_at == started,
                    )).mappings().first()
                    if row is not None and row["count"] >= policy.limit:
                        exhausted.add(policy.dimension)
                    rows.append((policy, digest, started, row))
                if exhausted:
                    connection.rollback()
                    dimensions = tuple(sorted(exhausted))
                    event = (
                        "auth_abuse_global_sms_cap"
                        if "global" in exhausted
                        else "auth_abuse_throttled"
                    )
                    channel = "sms" if any(item[0].channel == "sms" for item in prepared) else (
                        "email" if any(item[0].channel == "email" for item in prepared) else "any"
                    )
                    LOGGER.warning(
                        f"{event} action=%s channel=%s dimensions=%s",
                        prepared[0][0].action if prepared else "unknown",
                        channel,
                        ",".join(dimensions),
                    )
                    return Decision(False, dimensions)
                for policy, digest, started, row in rows:
                    if row is None:
                        connection.execute(table.insert().values(
                            action=policy.action, dimension=policy.dimension,
                            subject_digest=digest, channel=policy.channel,
                            window_started_at=started, count=1,
                            expires_at=started + timedelta(
                                seconds=policy.window_seconds + self.config.retention_seconds
                            ),
                        ))
                    else:
                        connection.execute(table.update().where(
                            table.c.id == row["id"]
                        ).values(count=row["count"] + 1))
                connection.commit()
        except (SQLAlchemyError, OSError) as exc:
            LOGGER.error(
                "auth_abuse_storage_error action=%s channel=%s reason=%s",
                prepared[0][0].action if prepared else "unknown",
                prepared[0][0].channel if prepared else "any",
                type(exc).__name__,
            )
            return Decision(False, (), True)
        channel = "sms" if any(item[0].channel == "sms" for item in prepared) else (
            "email" if any(item[0].channel == "email" for item in prepared) else "any"
        )
        LOGGER.info(
            "auth_abuse_allowed action=%s channel=%s dimensions=%s",
            prepared[0][0].action if prepared else "unknown",
            channel,
            ",".join(sorted({item[0].dimension for item in prepared})),
        )
        self.maybe_cleanup(now)
        return Decision(True)

    def cleanup(self, now: datetime | None = None) -> int:
        now = (now or datetime.now(timezone.utc)).replace(tzinfo=None)
        table = db.AuthAbuseCounter.__table__
        try:
            with self.engine.begin() as connection:
                ids = list(connection.scalars(
                    select(table.c.id).where(table.c.expires_at < now)
                    .order_by(table.c.expires_at).limit(self.config.cleanup_batch)
                ))
                if ids:
                    connection.execute(delete(table).where(table.c.id.in_(ids)))
        except (SQLAlchemyError, OSError) as exc:
            LOGGER.warning("auth_abuse_cleanup status=skipped reason=%s", type(exc).__name__)
            return 0
        LOGGER.info("auth_abuse_cleanup status=complete count=%s", len(ids))
        return len(ids)

    def maybe_cleanup(self, now: datetime | None = None) -> int:
        monotonic = time.monotonic()
        if monotonic - self._last_cleanup < self.config.cleanup_interval_seconds:
            return 0
        self._last_cleanup = monotonic
        return self.cleanup(now)
