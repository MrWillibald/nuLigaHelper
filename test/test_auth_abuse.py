"""Persistent authentication-abuse policy, storage, and privacy tests."""

import json
import logging
import os
import re
import tempfile
import threading
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timedelta, timezone

from sqlalchemy import inspect, select
from sqlalchemy.exc import OperationalError

import helpers as h
import auth_abuse
import db
import notifier
import webapp


def _database():
    path = os.path.join(h._TEST_DIR, f"abuse-{next(tempfile._get_candidate_names())}.db")
    engine = db.make_engine(path)
    db.init_db(engine)
    return path, engine


def _app(path, section):
    old_path = os.environ["NULIGAHELPER_DB"]
    old_config = os.environ.get("NULIGAHELPER_AUTH_ABUSE_CONFIG")
    os.environ["NULIGAHELPER_DB"] = path
    os.environ["NULIGAHELPER_AUTH_ABUSE_CONFIG"] = json.dumps(section)
    try:
        app = webapp.create_app()
    finally:
        os.environ["NULIGAHELPER_DB"] = old_path
        if old_config is None:
            os.environ.pop("NULIGAHELPER_AUTH_ABUSE_CONFIG", None)
        else:
            os.environ["NULIGAHELPER_AUTH_ABUSE_CONFIG"] = old_config
    app.config["TESTING"] = True
    return app


def _csrf(client, path="/login"):
    body = client.get(path).get_data(as_text=True)
    return re.search(r'name="csrf-token" content="([^"]+)"', body).group(1)


def test_policy_defaults_overrides_and_validation():
    config = auth_abuse.load_config()
    assert config.policies["login_client"].limit == 10
    assert config.policies["sms_global_cap"].window_seconds == 86400
    changed = auth_abuse.load_config({
        "policies": {"login_client": {"limit": 7, "window_seconds": 60}},
        "retention_seconds": 86400,
        "cleanup_batch": 12,
    })
    assert (changed.policies["login_client"].limit, changed.cleanup_batch) == (7, 12)
    invalid = [
        {"policies": {"unknown": {"limit": 1}}},
        {"policies": {"login_client": {"limit": 0}}},
        {"policies": {"login_client": {"window_seconds": -1}}},
        {"retention_seconds": 60},
        {"trusted_proxies": ["127.0.0.1/32"], "trusted_hops": 0},
        {"trusted_proxies": ["bad"], "trusted_hops": 1},
    ]
    for value in invalid:
        try:
            auth_abuse.load_config(value)
        except ValueError:
            pass
        else:
            raise AssertionError(f"invalid configuration was accepted: {value}")


def test_additive_table_initialization_and_opaque_domain_separated_subjects():
    path = os.path.join(h._TEST_DIR, f"existing-{next(tempfile._get_candidate_names())}.db")
    engine = db.make_engine(path)
    with engine.begin() as connection:
        connection.exec_driver_sql("CREATE TABLE legacy_data (id INTEGER PRIMARY KEY)")
        connection.exec_driver_sql("INSERT INTO legacy_data VALUES (1)")
    db.init_db(engine)
    names = set(inspect(engine).get_table_names())
    assert {"legacy_data", "auth_abuse_counters"} <= names
    with engine.connect() as connection:
        assert connection.exec_driver_sql("SELECT COUNT(*) FROM legacy_data").scalar_one() == 1
    service = auth_abuse.Service(engine, auth_abuse.load_config(), "test-secret")
    values = {
        service.digest("client", "any", "2001:0db8::1"),
        service.digest("contact", "email", "sentinel@example.test"),
        service.digest("contact", "sms", "+4915112345678"),
        service.digest("person", "email", "42"),
        service.digest("global", "sms", "ignored"),
    }
    assert len(values) == 5 and all(len(value) == 64 for value in values)
    service.reserve([
        auth_abuse.Rule("login_client", "2001:db8::1"),
        auth_abuse.Rule("login_contact_email", "sentinel@example.test"),
    ])
    with engine.connect() as connection:
        dump = " ".join(str(value) for row in connection.execute(
            select(db.AuthAbuseCounter.__table__)
        ) for value in row)
    for raw in ("2001:db8::1", "sentinel@example.test", "+4915112345678", "Sentinel Person"):
        assert raw not in dump


def test_atomic_limit_rollover_no_partial_increment_and_caller_isolation():
    _path, engine = _database()
    config = auth_abuse.load_config({
        "policies": {
            "login_client": {"limit": 2, "window_seconds": 60},
            "login_contact_email": {"limit": 1, "window_seconds": 60},
        }
    })
    service = auth_abuse.Service(engine, config, "secret")
    rules = [auth_abuse.Rule("login_client", "192.0.2.1"),
             auth_abuse.Rule("login_contact_email", "one@example.test")]
    now = datetime(2026, 1, 1, tzinfo=timezone.utc)
    assert service.reserve(rules, now).allowed
    refused = service.reserve(rules, now + timedelta(seconds=1))
    assert not refused.allowed and refused.reason_dimensions == ("contact",)
    with h.Session(engine) as session:
        client_row = session.query(db.AuthAbuseCounter).filter_by(dimension="client").one()
        assert client_row.count == 1, "a refused multi-rule reservation must increment nothing"
        session.add(db.Person(name="Pending Caller Work"))
        assert service.reserve(
            [auth_abuse.Rule("login_client", "192.0.2.2")], now
        ).allowed
        session.rollback()
    with h.Session(engine) as session:
        assert session.query(db.Person).filter_by(name="Pending Caller Work").count() == 0
    assert service.reserve(rules, now + timedelta(seconds=60)).allowed


def test_concurrent_final_allowance_and_restart_shared_state():
    path, engine = _database()
    config = auth_abuse.load_config({
        "policies": {
            "login_client": {"limit": 3, "window_seconds": 86400},
            "login_contact_email": {"limit": 3, "window_seconds": 86400},
            "sms_global_cap": {"limit": 3, "window_seconds": 86400},
        }
    })
    now = datetime(2026, 1, 1, tzinfo=timezone.utc)

    def contend(_):
        local = db.make_engine(path)
        try:
            return auth_abuse.Service(local, config, "secret").reserve(
                [
                    auth_abuse.Rule("login_client", "192.0.2.10"),
                    auth_abuse.Rule("login_contact_email", "race@example.test"),
                    auth_abuse.Rule("sms_global_cap", "application"),
                ], now
            ).allowed
        finally:
            local.dispose()

    with ThreadPoolExecutor(max_workers=8) as pool:
        outcomes = list(pool.map(contend, range(8)))
    assert sum(outcomes) == 3
    restarted = auth_abuse.Service(db.make_engine(path), config, "secret")
    assert not restarted.reserve(
        [
            auth_abuse.Rule("login_client", "192.0.2.10"),
            auth_abuse.Rule("login_contact_email", "race@example.test"),
            auth_abuse.Rule("sms_global_cap", "application"),
        ], now
    ).allowed
    with h.Session(engine) as session:
        counts = {
            row.dimension: row.count
            for row in session.query(db.AuthAbuseCounter).all()
        }
        assert counts == {"client": 3, "contact": 3, "global": 3}
    engine.dispose()


def test_separate_app_instances_share_client_cooling_off_and_ignore_spoofed_headers():
    path, engine = _database()
    section = {"policies": {"login_client": {"limit": 1, "window_seconds": 900}}}
    first = _app(path, section)
    second = _app(path, section)
    first_client = first.test_client()
    second_client = second.test_client()
    first_client.post("/login", data={
        "csrf_token": _csrf(first_client), "channel": "email",
        "email": "one@example.test", "action": "request_code",
    }, headers={"X-Forwarded-For": "192.0.2.1", "Forwarded": "for=192.0.2.1"})
    response = second_client.post("/login", data={
        "csrf_token": _csrf(second_client), "channel": "email",
        "email": "two@example.test", "action": "request_code",
    }, headers={"X-Forwarded-For": "192.0.2.2", "Forwarded": "for=192.0.2.2"})
    assert response.status_code == 200
    assert "Falls die Angaben bekannt sind" in response.get_data(as_text=True)
    with h.Session(engine) as session:
        rows = session.query(db.AuthAbuseCounter).filter_by(
            action="login_request", dimension="client"
        ).all()
        assert len(rows) == 1 and rows[0].count == 1


def test_bounded_cleanup_preserves_live_rows_and_drains_after_restart():
    path, engine = _database()
    config = auth_abuse.load_config({"cleanup_batch": 2})
    table = db.AuthAbuseCounter.__table__
    now = datetime(2026, 1, 10)
    with engine.begin() as connection:
        for index in range(5):
            connection.execute(table.insert().values(
                action="login_request", dimension="client",
                subject_digest=f"{index:064x}", channel="any",
                window_started_at=now - timedelta(days=10), count=1,
                expires_at=now - timedelta(days=1),
            ))
        connection.execute(table.insert().values(
            action="login_request", dimension="client",
            subject_digest="f" * 64, channel="any",
            window_started_at=now, count=1, expires_at=now + timedelta(days=1),
        ))
    service = auth_abuse.Service(engine, config, "secret")
    assert service.cleanup(now) == 2
    assert service.cleanup(now) == 2
    restarted = auth_abuse.Service(db.make_engine(path), config, "secret")
    assert restarted.cleanup(now) == 1
    with h.Session(engine) as session:
        rows = session.query(db.AuthAbuseCounter).all()
        assert len(rows) == 1 and rows[0].subject_digest == "f" * 64


def test_storage_failure_is_structured_and_client_attribution_is_strict():
    class BrokenEngine:
        def connect(self):
            raise OperationalError("BEGIN", {}, Exception("private provider payload"))

        def begin(self):
            raise OperationalError("DELETE", {}, Exception("private cleanup payload"))

    decision = auth_abuse.Service(
        BrokenEngine(), auth_abuse.load_config(), "secret"
    ).reserve([auth_abuse.Rule("login_client", "192.0.2.1")])
    assert not decision.allowed and decision.storage_error
    assert auth_abuse.Service(
        BrokenEngine(), auth_abuse.load_config(), "secret"
    ).cleanup() == 0

    direct = auth_abuse.load_config()
    assert auth_abuse.attributed_client(
        "2001:0db8::1", "192.0.2.99", direct
    ) == "2001:db8::1"
    trusted = auth_abuse.load_config({
        "trusted_proxies": ["127.0.0.0/8", "10.0.0.0/8"], "trusted_hops": 2,
    })
    assert auth_abuse.attributed_client(
        "127.0.0.1", "198.51.100.4, 10.0.0.2", trusted
    ) == "198.51.100.4"
    for peer, header in (
        ("198.51.100.8", "203.0.113.1"),
        ("127.0.0.1", "bad"),
        ("127.0.0.1", "203.0.113.1"),
        ("127.0.0.1", "203.0.113.1, 192.0.2.2"),
    ):
        assert auth_abuse.attributed_client(peer, header, trusted) == auth_abuse.canonical_ip(peer)


def test_sms_global_cap_is_shared_with_registration_email_is_independent_and_logs_are_redacted():
    path, engine = _database()
    section = {"policies": {
        "login_client": {"limit": 20, "window_seconds": 900},
        "login_contact_sms": {"limit": 20, "window_seconds": 900},
        "login_person_sms": {"limit": 20, "window_seconds": 900},
        "registration_client": {"limit": 20, "window_seconds": 900},
        "registration_contact_sms": {"limit": 20, "window_seconds": 900},
        "registration_person_sms": {"limit": 20, "window_seconds": 900},
        "sms_contact_cap": {"limit": 20, "window_seconds": 86400},
        "sms_person_cap": {"limit": 20, "window_seconds": 86400},
        "sms_global_cap": {"limit": 1, "window_seconds": 86400},
    }}
    with h.Session(engine) as session:
        team = db.get_support_team(session)
        people = [
            db.Person(name="Sentinel One", email="one@example.test", phone="+4915111111111",
                      team=team, account_status=db.ACCOUNT_ACTIVE),
            db.Person(name="Sentinel Two", email="two@example.test", phone="+4915222222222",
                      team=team, account_status=db.ACCOUNT_ACTIVE),
        ]
        session.add_all(people)
        session.commit()
        team_id = team.id
    app = _app(path, section)
    sent = []
    original = notifier.Notifier.send_account_message_via
    notifier.Notifier.send_account_message_via = (
        lambda self, person, channel, subject, body: sent.append(channel) or 1
    )
    records = []
    handler = logging.Handler()
    handler.emit = lambda record: records.append(record.getMessage())
    auth_abuse.LOGGER.addHandler(handler)
    auth_abuse.LOGGER.setLevel(logging.INFO)
    try:
        client = app.test_client()
        first = client.post("/login", data={
            "csrf_token": _csrf(client), "action": "request_code", "channel": "sms",
            "country_code": "+49", "phone": "15111111111",
        }, environ_overrides={"REMOTE_ADDR": "198.51.100.77"})
        registration = client.post("/registrieren", data={
            "csrf_token": _csrf(client, "/registrieren"), "action": "request_code",
            "name": "Sentinel Two", "team_id": str(team_id), "consent": "yes",
            "channel": "sms", "country_code": "+49", "phone": "15222222222",
        }, environ_overrides={"REMOTE_ADDR": "198.51.100.77"})
        email = client.post("/login", data={
            "csrf_token": _csrf(client), "action": "request_code", "channel": "email",
            "email": "two@example.test",
        }, environ_overrides={"REMOTE_ADDR": "198.51.100.77"})
        assert all(response.status_code == 200 for response in (first, registration, email))
        assert sent == ["sms", "email"]
        assert "Falls die Angaben verwendet werden können" in registration.get_data(as_text=True)
    finally:
        notifier.Notifier.send_account_message_via = original
        auth_abuse.LOGGER.removeHandler(handler)
    joined = " ".join(records)
    assert "auth_abuse_global_sms_cap" in joined
    for sentinel in (
        "Sentinel", "198.51.100.77", "one@example.test", "+4915111111111",
        "private provider payload",
    ):
        assert sentinel not in joined


def test_each_sms_cost_cap_rolls_over_and_concurrent_dispatch_never_exceeds_global_limit():
    now = datetime(2026, 1, 1, tzinfo=timezone.utc)
    for policy, subject in (
        ("sms_contact_cap", "+4915111111111"),
        ("sms_person_cap", "42"),
        ("sms_global_cap", "application"),
    ):
        _path, engine = _database()
        config = auth_abuse.load_config({
            "policies": {policy: {"limit": 1, "window_seconds": 86400}}
        })
        service = auth_abuse.Service(engine, config, "secret")
        rule = [auth_abuse.Rule(policy, subject)]
        assert service.reserve(rule, now).allowed
        assert not service.reserve(rule, now + timedelta(hours=1)).allowed
        assert service.reserve(rule, now + timedelta(days=1)).allowed
        engine.dispose()

    path, engine = _database()
    section = {"policies": {
        "login_client": {"limit": 50, "window_seconds": 900},
        "login_contact_sms": {"limit": 50, "window_seconds": 900},
        "login_person_sms": {"limit": 50, "window_seconds": 900},
        "sms_contact_cap": {"limit": 50, "window_seconds": 86400},
        "sms_person_cap": {"limit": 50, "window_seconds": 86400},
        "sms_global_cap": {"limit": 3, "window_seconds": 86400},
    }}
    phones = [f"+4915111111{index:03d}" for index in range(8)]
    with h.Session(engine) as session:
        session.add_all([
            db.Person(name=f"Concurrent {index}", phone=phone,
                      account_status=db.ACCOUNT_ACTIVE)
            for index, phone in enumerate(phones)
        ])
        session.commit()
    app = _app(path, section)
    sent = []
    sent_lock = threading.Lock()
    original = notifier.Notifier.send_account_message_via

    def fake_send(self, person, channel, subject, body):
        with sent_lock:
            sent.append(person.id)
        return 1

    notifier.Notifier.send_account_message_via = fake_send

    def request_code(item):
        index, phone = item
        client = app.test_client()
        response = client.post("/login", data={
            "csrf_token": _csrf(client), "action": "request_code", "channel": "sms",
            "country_code": "+49", "phone": phone[3:],
        }, environ_overrides={"REMOTE_ADDR": f"198.51.100.{index + 1}"})
        return response.status_code

    try:
        with ThreadPoolExecutor(max_workers=8) as pool:
            statuses = list(pool.map(request_code, enumerate(phones)))
    finally:
        notifier.Notifier.send_account_message_via = original
    assert statuses == [200] * 8
    assert len(sent) == 3
    with h.Session(engine) as session:
        assert session.query(db.AuthToken).count() == 3


def test_route_storage_failure_keeps_success_shape_and_has_zero_side_effects():
    path, engine = _database()
    with h.Session(engine) as session:
        session.add(db.Person(
            name="Known", email="known@example.test", account_status=db.ACCOUNT_ACTIVE
        ))
        session.commit()
    app = _app(path, {})
    service = app.extensions["nuligahelper_auth_abuse"]
    service.reserve = lambda rules, now=None: auth_abuse.Decision(False, (), True)
    sent = []
    original = notifier.Notifier.send_account_message_via
    notifier.Notifier.send_account_message_via = lambda *args: sent.append(True) or 1
    try:
        client = app.test_client()
        response = client.post("/login", data={
            "csrf_token": _csrf(client), "action": "request_code", "channel": "email",
            "email": "known@example.test",
        })
    finally:
        notifier.Notifier.send_account_message_via = original
    body = response.get_data(as_text=True)
    assert response.status_code == 200
    assert 'name="challenge"' in body
    assert "Falls die Angaben bekannt sind" in body
    assert not sent
    with h.Session(engine) as session:
        assert session.query(db.AuthToken).count() == 0


if __name__ == "__main__":
    h.run_all(globals())
