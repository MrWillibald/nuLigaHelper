## Why

The Flask development server is intentionally unsuitable for public exposure, leaving the Raspberry Pi web interface without a supported production WSGI runtime, TLS termination, or a defined trust boundary for proxy-derived request data. A documented, reproducible production deployment is needed before the service can be exposed through a public hostname without weakening the existing local development workflow.

## What Changes

- Add Gunicorn as the Linux/Raspberry Pi production WSGI server for the existing module-level `webapp.app`, bound only to loopback and supervised by systemd.
- Add a Caddy reverse-proxy configuration that serves the configured public hostname, obtains and renews HTTPS certificates automatically, and proxies only to the loopback Gunicorn listener.
- Define production-only Flask hardening for secure session cookies, explicit trusted hosts, bounded request bodies, and security response headers while preserving useful direct local developer launching.
- Trust forwarded scheme, host, and client-address values only when requests come through the single configured Caddy hop; reject untrusted hosts and avoid accepting arbitrary proxy chains or internet-supplied forwarding headers.
- Document installation, persistent `NULIGAHELPER_SECRET` and database configuration, file ownership, startup, certificate prerequisites, deployment verification, logging, restart, update, backup, and rollback procedures for a Raspberry Pi/Linux host.
- Add offline tests for application configuration and proxy/host/header behavior where feasible, plus syntax or configuration checks that do not require certificate issuance or network access.
- Preserve the existing Flask/Jinja application, SQLite persistence model, and daily-job integration.
- Explicit non-goals: rewriting the application with FastAPI, moving from SQLite to PostgreSQL, or introducing containers/Kubernetes.

## Capabilities

### New Capabilities

- `production-web-runtime`: Defines the supported production serving topology, process supervision, HTTPS boundary, request trust and hardening behavior, deployable configuration, and offline verification expectations.

### Modified Capabilities

None.

## Impact

- Application configuration and WSGI startup in `webapp.py`, while retaining the existing module-level `app` import target.
- Python dependencies and launch scripts, including `requirements.txt` and `run_webapp.sh` or an additional production launcher.
- New deployment assets for Gunicorn, systemd, and Caddy, plus production documentation in `README.MD` and/or a dedicated deployment guide.
- Flask request handling for proxy metadata, host validation, cookie flags, request-size enforcement, and security headers.
- Offline tests under `test/` and existing test launch commands.
- Raspberry Pi operations: a dedicated service account/permissions, stable `NULIGAHELPER_SECRET`, writable SQLite/log locations, public DNS and inbound ports 80/443, and Caddy-managed certificates.
