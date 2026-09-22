"""Private deployment asset checks."""
from pathlib import Path
import helpers as h

PROJECT = Path(h.PROJECT_DIR)

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
    assert "sk_" not in environment
    assert all(address in environment for address in (
        "sender@example.invalid", "paper@example.invalid", "admin@example.invalid",
        "sale@example.invalid"))
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
    for header in ("X-Forwarded-For", "X-Forwarded-Proto", "X-Forwarded-Host"):
        assert f"header_up -{header}" not in caddy, "deleting a forwarded header after setting it breaks the proxy boundary"
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
