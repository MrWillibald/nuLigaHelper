"""Offline production-runtime and versioned deployment-asset checks."""

import logging
import os
import tempfile
from pathlib import Path

from flask import jsonify, redirect, request, session

import helpers as h
import common
import db
import webapp

HOST = "nuliga.example.invalid"
PROJECT = Path(h.PROJECT_DIR)


def _production_app(*, load_config=None):
    path = os.path.join(h._TEST_DIR, f"production-{next(tempfile._get_candidate_names())}.db")
    values = {
        "NULIGAHELPER_ENV": "production",
        "NULIGAHELPER_SECRET": "synthetic-stable-production-secret",
        "NULIGAHELPER_DB": path,
        "NULIGAHELPER_TRUSTED_HOSTS": HOST,
    }
    previous = {name: os.environ.get(name) for name in values}
    original_loader = common.load_config
    os.environ.update(values)
    if load_config is not None:
        common.load_config = load_config
    try:
        app = webapp.create_app()
    finally:
        common.load_config = original_loader
        for name, old_value in previous.items():
            if old_value is None:
                os.environ.pop(name, None)
            else:
                os.environ[name] = old_value
    app.config["TESTING"] = True
    return app, path


def _proxy(client, path="/", *, method="GET", client_ip="203.0.113.10",
           host=HOST, scheme="https", headers=None, **kwargs):
    canonical = {
        "X-Forwarded-For": client_ip,
        "X-Forwarded-Proto": scheme,
        "X-Forwarded-Host": host,
    }
    canonical.update(headers or {})
    return client.open(
        path,
        method=method,
        base_url=f"https://{HOST}",
        headers=canonical,
        environ_overrides={"REMOTE_ADDR": "127.0.0.1", "wsgi.url_scheme": "http"},
        **kwargs,
    )


def _with_environment(values, callback):
    names = {
        "NULIGAHELPER_ENV", "NULIGAHELPER_SECRET", "NULIGAHELPER_DB",
        "NULIGAHELPER_TRUSTED_HOSTS",
    }
    previous = {name: os.environ.get(name) for name in names}
    for name in names:
        os.environ.pop(name, None)
    os.environ.update(values)
    try:
        return callback()
    finally:
        for name, old_value in previous.items():
            if old_value is not None:
                os.environ[name] = old_value
            else:
                os.environ.pop(name, None)


def _add_active_person(path):
    with h.Session(db.make_engine(path)) as database_session:
        person = db.Person(
            name="Production Tester", email="prod@example.test",
            account_status=db.ACCOUNT_ACTIVE,
        )
        database_session.add(person)
        database_session.commit()
        return person.id


def _sign_in_production(client, person_id):
    with client.session_transaction(
        base_url=f"https://{HOST}"
    ) as browser_session:
        browser_session["person_id"] = person_id
        browser_session["csrf_token"] = "test-csrf-token"
        browser_session.permanent = True


def test_production_startup_validation_is_fail_closed_and_config_free():
    app, path = _production_app(
        load_config=lambda: (_ for _ in ()).throw(
            AssertionError("production app creation must not read config.json")
        )
    )
    assert app.config["NULIGAHELPER_PRODUCTION"] is True
    assert app.config["TRUSTED_HOSTS"] == [HOST]
    assert os.path.exists(path)

    valid = {
        "NULIGAHELPER_ENV": "production",
        "NULIGAHELPER_SECRET": "stable-secret",
        "NULIGAHELPER_DB": "/tmp/nuliga-production-validation.db",
        "NULIGAHELPER_TRUSTED_HOSTS": "one.example.invalid,two.example.invalid",
    }
    refused = [
        ("NULIGAHELPER_SECRET", ""),
        ("NULIGAHELPER_DB", "relative.db"),
        ("NULIGAHELPER_TRUSTED_HOSTS", ""),
        ("NULIGAHELPER_TRUSTED_HOSTS", "*.example.invalid"),
        ("NULIGAHELPER_TRUSTED_HOSTS", "https://example.invalid"),
        ("NULIGAHELPER_TRUSTED_HOSTS", "example.invalid/path"),
        ("NULIGAHELPER_TRUSTED_HOSTS", "example.invalid:443"),
        ("NULIGAHELPER_TRUSTED_HOSTS", "Example.invalid"),
        ("NULIGAHELPER_TRUSTED_HOSTS", "example..invalid"),
        ("NULIGAHELPER_TRUSTED_HOSTS", "example.invalid,example.invalid"),
    ]
    for name, value in refused:
        environment = {**valid, name: value}
        try:
            _with_environment(environment, webapp.create_app)
        except RuntimeError as exc:
            assert name in str(exc), f"failure for {name} must be actionable"
        else:
            raise AssertionError(f"production startup accepted invalid {name}={value!r}")

    try:
        _with_environment({**valid, "NULIGAHELPER_ENV": "staging"}, webapp.create_app)
    except RuntimeError as exc:
        assert "NULIGAHELPER_ENV" in str(exc)
    else:
        raise AssertionError("unknown environment mode must fail")


def test_local_mode_retains_http_and_does_not_enable_production_implicitly():
    original_loader = common.load_config
    common.load_config = lambda: (_ for _ in ()).throw(
        AssertionError("an explicit synthetic test database must avoid config.json")
    )
    try:
        app = _with_environment({
            "NULIGAHELPER_SECRET": "synthetic-local-secret",
            "NULIGAHELPER_DB": os.path.join(h._TEST_DIR, "local-runtime.db"),
        }, webapp.create_app)
    finally:
        common.load_config = original_loader
    app.config["TESTING"] = True
    assert app.config["NULIGAHELPER_PRODUCTION"] is False
    assert app.config["SESSION_COOKIE_SECURE"] is False
    assert app.config["TRUSTED_HOSTS"] is None
    assert app.test_client().get("/", base_url="http://localhost").status_code == 200
    assert webapp.app.config["NULIGAHELPER_PRODUCTION"] is False


def test_proxy_boundary_accepts_one_hop_and_rejects_every_uncanonical_shape():
    app, path = _production_app()

    @app.get("/_metadata")
    def metadata():
        return jsonify(ip=request.remote_addr, scheme=request.scheme, host=request.host)

    client = app.test_client()
    _sign_in_production(client, _add_active_person(path))
    accepted = _proxy(client, "/_metadata")
    assert accepted.status_code == 200
    assert accepted.get_json() == {"ip": "203.0.113.10", "scheme": "https", "host": HOST}

    bad_requests = [
        {"environ_overrides": {"REMOTE_ADDR": "198.51.100.2"}},
        {"headers": {"X-Forwarded-For": ""}},
        {"headers": {"X-Forwarded-For": "not-an-ip"}},
        {"headers": {"X-Forwarded-For": "203.0.113.1, 127.0.0.1"}},
        {"headers": {"X-Forwarded-Proto": "http"}},
        {"headers": {"X-Forwarded-Proto": "https,http"}},
        {"headers": {"X-Forwarded-Host": ""}},
        {"headers": {"X-Forwarded-Host": "bad.example.invalid:443"}},
        {"headers": {"Forwarded": "for=198.51.100.1"}},
        {"headers": {"X-Forwarded-Port": "443"}},
        {"headers": {"X-Forwarded-Prefix": "/spoof"}},
    ]
    for case in bad_requests:
        headers = {
            "X-Forwarded-For": "203.0.113.10",
            "X-Forwarded-Proto": "https",
            "X-Forwarded-Host": HOST,
            **case.get("headers", {}),
        }
        response = client.get(
            "/_metadata", base_url=f"https://{HOST}", headers=headers,
            environ_overrides={
                "wsgi.url_scheme": "http",
                **case.get("environ_overrides", {"REMOTE_ADDR": "127.0.0.1"}),
            },
        )
        assert response.status_code == 400, case


def test_trusted_host_rejects_before_mutation_and_cookie_is_hardened():
    app, path = _production_app()
    mutations = []

    @app.get("/_session")
    def set_session():
        session["test"] = True
        session.permanent = True
        return "ok"

    @app.post("/_mutate")
    def mutate():
        mutations.append(True)
        return "changed"

    client = app.test_client()
    _sign_in_production(client, _add_active_person(path))
    cookie = _proxy(client, "/_session").headers.get("Set-Cookie", "")
    assert "Secure" in cookie
    assert "HttpOnly" in cookie
    assert "SameSite=Lax" in cookie
    assert "Path=/" in cookie
    assert "Domain=" not in cookie
    assert "Expires=" in cookie

    refused = _proxy(client, "/_mutate", method="POST", host="other.example.invalid")
    assert refused.status_code == 400
    assert not mutations


def test_request_limit_rejects_before_endpoint_mutation():
    app, path = _production_app()
    person_id = _add_active_person(path)
    mutations = []

    @app.post("/_bounded-mutation")
    def bounded_mutation():
        mutations.append(request.get_data())
        return "changed"

    client = app.test_client()
    _sign_in_production(client, person_id)
    for content_type, data in (
        ("application/json", b"x" * (webapp.MAX_PRODUCTION_BODY_BYTES + 1)),
        ("application/x-www-form-urlencoded",
         b"csrf_token=test-csrf-token&value=" + b"x" * webapp.MAX_PRODUCTION_BODY_BYTES),
    ):
        oversized = _proxy(
            client, "/_bounded-mutation", method="POST", data=data,
            content_type=content_type,
            headers={"X-CSRF-Token": "test-csrf-token"},
        )
        assert oversized.status_code == 413
        assert not mutations
    accepted = _proxy(
        client, "/_bounded-mutation", method="POST", data=b"{}",
        content_type="application/json",
        headers={"X-CSRF-Token": "test-csrf-token"},
    )
    assert accepted.status_code == 200
    assert mutations == [b"{}"]


def test_security_headers_cover_html_json_redirect_error_and_static(caplog=None):
    app, path = _production_app()

    @app.get("/_json")
    def json_response():
        return jsonify(ok=True)

    @app.get("/_redirect")
    def redirect_response():
        return redirect("/")

    expected = {**webapp.PRODUCTION_SECURITY_HEADERS,
                "Strict-Transport-Security": "max-age=31536000"}
    client = app.test_client()
    _sign_in_production(client, _add_active_person(path))
    for path in ("/", "/_json", "/_redirect", "/missing", "/static/app.js"):
        response = _proxy(client, path)
        expected_status = {"/_json": 200, "/_redirect": 302, "/missing": 404}.get(path)
        if expected_status is not None:
            assert response.status_code == expected_status
        for name, value in expected.items():
            assert response.headers.getlist(name) == [value], (path, name)
    assert "'unsafe-inline'" not in expected["Content-Security-Policy"].split("script-src", 1)[1].split(";", 1)[0]


def test_effective_client_identity_drives_logging_and_rate_limit_keys():
    app, path = _production_app()
    client = app.test_client()
    with app.test_request_context():
        pass
    logger = logging.getLogger()
    records = []

    class Capture(logging.Handler):
        def emit(self, record):
            records.append(record.getMessage())

    handler = Capture()
    logger.addHandler(handler)
    old_level = logger.level
    logger.setLevel(logging.INFO)
    try:
        for address in ("198.51.100.10", "198.51.100.11"):
            response = _proxy(client, "/login")
            csrf = response.get_data(as_text=True).split('name="csrf-token" content="', 1)[1].split('"', 1)[0]
            _proxy(
                client, "/login", method="POST", client_ip=address,
                data={"channel": "email", "email": "nobody@example.test", "csrf_token": csrf},
            )
        service = app.extensions["nuligahelper_auth_abuse"]
        expected = {
            service.digest("client", "any", "198.51.100.10"),
            service.digest("client", "any", "198.51.100.11"),
        }
        with h.Session(db.make_engine(path)) as database_session:
            stored = set(database_session.scalars(
                db.select(db.AuthAbuseCounter.subject_digest).where(
                    db.AuthAbuseCounter.dimension == "client"
                )
            ))
        assert expected <= stored
        assert not any("198.51.100" in message for message in records)
        before = set(stored)
        spoofed = _proxy(
            client, "/login", headers={"Forwarded": "for=192.0.2.99"}
        )
        assert spoofed.status_code == 400
        with h.Session(db.make_engine(path)) as database_session:
            after = set(database_session.scalars(
                db.select(db.AuthAbuseCounter.subject_digest).where(
                    db.AuthAbuseCounter.dimension == "client"
                )
            ))
        assert after == before
    finally:
        logger.removeHandler(handler)
        logger.setLevel(old_level)


def test_versioned_deployment_assets_are_bounded_and_placeholder_only():
    requirements = (PROJECT / "requirements-production.txt").read_text()
    gunicorn = (PROJECT / "deploy/gunicorn.conf.py").read_text()
    environment = (PROJECT / "deploy/nuligahelper.env.example").read_text()
    unit = (PROJECT / "deploy/nuligahelper-web.service").read_text()
    caddy = (PROJECT / "deploy/Caddyfile.example").read_text()
    guide = (PROJECT / "deploy/PRODUCTION.md").read_text()
    assert "-r requirements.txt" in requirements and "gunicorn>=23.0,<24" in requirements
    for setting in ('bind = "127.0.0.1:8080"', "workers = 1", 'worker_class = "sync"',
                    'accesslog = "-"', 'errorlog = "-"', "limit_request_line"):
        assert setting in gunicorn
    for key in ("NULIGAHELPER_ENV=production", "NULIGAHELPER_SECRET=REPLACE_",
                "NULIGAHELPER_DB=/var/lib/nuligahelper/", "NULIGAHELPER_TRUSTED_HOSTS="):
        assert key in environment
    assert "example.invalid" in environment
    assert "@" not in environment and "sk_" not in environment
    for directive in ("User=nuligahelper", "EnvironmentFile=/etc/nuligahelper/web.env",
                      "webapp:app", "Restart=on-failure", "UMask=0077",
                      "NoNewPrivileges=true", "ProtectSystem=strict",
                      "ReadWritePaths=/var/lib/nuligahelper"):
        assert directive in unit
    for directive in ("reverse_proxy 127.0.0.1:8080", "max_size 1MB",
                      "header_up -Forwarded", "header_up X-Forwarded-For {remote_host}",
                      "header_up X-Forwarded-Proto https", "header_up X-Forwarded-Host {host}",
                      "Strict-Transport-Security", "Content-Security-Policy",
                      "X-Content-Type-Options", "X-Frame-Options", "Referrer-Policy",
                      "Permissions-Policy"):
        assert directive in caddy
    for phrase in (
        "ports 80 and 443", "requirements-production.txt", "dedicated",
        "/etc/nuligahelper/web.env", "root-owned and mode 0600",
        "rotating it invalidates every browser session", "identical stable",
        "systemctl enable --now nuligahelper-web.service",
        "systemctl enable --now caddy.service", "127.0.0.1:8080",
        "curl --head", "--max-time", "SQLite-consistent backup",
        "stop public ingress first", "Never delete/recreate",
        "never a public rollback",
    ):
        assert phrase.lower() in guide.lower(), phrase


if __name__ == "__main__":
    h.run_all(globals())
