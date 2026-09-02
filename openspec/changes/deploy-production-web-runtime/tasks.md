## 1. Production Configuration and Dependencies

- [ ] 1.1 Add a production requirements file that includes the base requirements and a bounded Gunicorn dependency, adjust the minimum Flask/Werkzeug compatibility only if required for trusted-host support, and verify a clean virtual-environment install resolves the declared files without changing development-only launch requirements.
- [ ] 1.2 Add explicit `NULIGAHELPER_ENV=production` configuration handling while preserving the existing default local mode, and verify focused tests show local HTTP remains usable and production settings are not enabled implicitly.
- [ ] 1.3 Validate production startup inputs for a non-empty stable `NULIGAHELPER_SECRET`, absolute `NULIGAHELPER_DB`, and one or more exact `NULIGAHELPER_TRUSTED_HOSTS` entries, rejecting wildcards, schemes, paths, malformed entries, and missing values; verify synthetic startup tests cover every accepted and refused case without reading `config.json`.

## 2. Flask Production Request Hardening

- [ ] 2.1 Implement the production-only raw WSGI proxy-boundary validator and exactly-one-hop proxy correction for client address, HTTPS scheme, and host, leaving proxy correction disabled locally; verify focused offline tests accept one canonical loopback-proxy request and reject non-loopback peers plus missing, spoofed, comma-separated, malformed, or non-HTTPS metadata.
- [ ] 2.2 Apply exact trusted-host enforcement after proxy correction and verify requests for configured hosts route normally while absent, malformed, and unlisted hosts return a client error before a state-changing endpoint runs.
- [ ] 2.3 Configure production sessions as Secure, HttpOnly, SameSite=Lax, Path=/, and host-only while retaining the one-hour sliding lifetime, and verify login/session tests inspect all cookie attributes and confirm the local HTTP cookie remains usable.
- [ ] 2.4 Add the 1 MiB application request-body limit and an appropriate HTTP 413 response path, and verify oversized form and JSON requests cannot invoke endpoint mutation while requests under the limit retain current behavior.
- [ ] 2.5 Add the production response-header baseline from `design.md`, including HSTS only for the trusted HTTPS production path and the enforcing CSP with no inline-script allowance, and verify representative HTML, JSON, redirect, error, and static responses contain the intended non-duplicated headers without breaking current page resources.
- [ ] 2.6 Verify effective proxy metadata reaches all existing client-IP consumers, especially authentication rate limiting and request logging, by adding focused tests that distinguish separate rewritten client addresses and demonstrate that supplied internet forwarding headers cannot choose the effective identity.

## 3. Gunicorn and systemd Runtime Assets

- [ ] 3.1 Add versioned Gunicorn configuration for `webapp:app` with one synchronous worker, `127.0.0.1:8080` binding, bounded timeouts and request/header limits, and journald-compatible stdout/stderr logging; verify a bounded local startup/import smoke test uses a synthetic secret and temporary SQLite path and exposes no non-loopback listener.
- [ ] 3.2 Add a placeholder-only production environment-file example with production mode, absolute shared database path, exact trusted host, and secret placeholder, and verify an automated static check confirms required keys are present and no real secret, hostname, credential, or personal data is committed.
- [ ] 3.3 Add a systemd service unit example using the dedicated unprivileged account, project virtual environment, environment file, working directory, restrictive umask, bounded restart policy, journald, and practical filesystem/process hardening with only the SQLite state directory writable; verify `systemd-analyze verify` succeeds where systemd tooling is installed and a static offline test asserts the critical supervision, loopback-runtime, and hardening directives.

## 4. Caddy HTTPS Boundary

- [ ] 4.1 Add a placeholder Caddyfile for one explicit public site that automatically manages HTTPS, redirects HTTP, proxies only to `127.0.0.1:8080`, strips client-supplied forwarding headers, and writes one canonical client-address/scheme/host hop; verify `caddy validate` succeeds where Caddy is installed and static tests assert the backend and header-rewrite trust boundary.
- [ ] 4.2 Configure Caddy’s matching 1 MiB edge body limit and security headers for proxy-generated redirects/errors without appending duplicates to backend responses, and verify static tests cover the limit plus HSTS, CSP, content-type, frame, referrer, and permissions directives.
- [ ] 4.3 Add documented local smoke-test commands for host matching, HTTP-to-HTTPS redirect, loopback-only backend access, forwarding-header overwrite, oversized HTTP 413, and edge error headers, and verify each command is bounded and clearly distinguishes offline/local checks from certificate-authority or public-network checks.

## 5. Raspberry Pi Deployment and Operations Documentation

- [ ] 5.1 Add a production deployment guide covering supported Linux/Raspberry Pi prerequisites, DNS and ports 80/443, Caddy and Python/Gunicorn installation, dedicated account/directories, checkout/venv permissions, and rendering site-local config from the examples; verify a clean-document review finds no undocumented prerequisite between checkout and first service start.
- [ ] 5.2 Document creation and mode/ownership of the external persistent environment file, stable `NULIGAHELPER_SECRET`, absolute `NULIGAHELPER_DB`, read-only `config.json` needs, and sharing the same secret/database with the daily job; verify the guide explicitly warns that secret rotation invalidates sessions and prevents divergent SQLite files.
- [ ] 5.3 Document configuration validation, service enable/start/reload ordering, loopback-listener inspection, journald diagnostics, certificate readiness, HTTPS/login/assignment smoke tests, and ongoing certificate supervision; verify commands name the correct Gunicorn and Caddy units and do not expose secret values in command output.
- [ ] 5.4 Document update, SQLite-consistent backup, ownership checks, restart, health verification, and rollback procedures that preserve data by default and stop public ingress first on failure; verify the rollback procedure never recommends exposing Werkzeug to the public internet or deleting/recreating the database.
- [ ] 5.5 Update `README.MD` and `run_webapp.sh` messaging to label the existing direct Flask launch as local/development-only, link the production guide, and describe Gunicorn/Caddy/systemd as the supported production path; verify local launch behavior still works and the obsolete “production serving remains required” statement is replaced accurately.

## 6. Offline Regression and Deployment Validation

- [ ] 6.1 Add or reorganize synthetic offline tests so production app creation does not leak environment/module-level `app` state between tests and never reads real `config.json`, then verify the focused production-runtime test module passes independently and in the existing order-dependent web scenario.
- [ ] 6.2 Run all focused authentication, refusal, concurrency, and webapp tests and verify proxy hardening, secure cookies, host checks, request limits, and headers introduce no regressions in passwordless authentication, CSRF protection, access tiers, or slot compare-and-swap behavior.
- [ ] 6.3 Run `test/run_tests.sh` offline and verify the complete suite passes with `DEBUG_FLAG` and `CHANGE_DAY` both left `False` and no test contacting DNS, certificate authorities, package registries, or public services.
- [ ] 6.4 Run `openspec validate deploy-production-web-runtime --strict` and review every delivered deployment example against the delta spec and design, verifying all requirements and scenarios have corresponding implementation, test, documentation, or explicitly deployment-time validation coverage.
