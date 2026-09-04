## Context

See `proposal.md` for motivation. The current `webapp.py` has a `create_app()` factory and eagerly creates a module-level `app`, which is already a valid WSGI import target. Its `__main__` block binds Werkzeug to `0.0.0.0:8080`, and `run_webapp.sh` invokes that path after loading `.nuligahelper_secret`. Sessions already use a one-hour sliding lifetime and SameSite=Lax, but production secure-cookie behavior, host validation, request limits, proxy trust, and security headers are not configured.

The application is a low-traffic Flask/Jinja service on a Raspberry Pi. It stores state in SQLite and also maintains an in-process authentication rate-limit deque. The templates load script and stylesheet assets from the same origin; several templates generate controlled inline `style` attributes. The daily job and web process must continue to share the same SQLite file and stable `NULIGAHELPER_SECRET`. There is deliberately no database migration layer.

## Goals / Non-Goals

**Goals:**

- Establish one documented production path with a small operational footprint suitable for Raspberry Pi OS or another systemd-based Linux distribution.
- Make the trust boundary explicit: the internet reaches Caddy, Caddy reaches one loopback Gunicorn listener, and only metadata rewritten by that one proxy hop is trusted.
- Fail closed when production-only secret, host, proxy, or database configuration is incomplete.
- Keep deployment configuration reviewable in the repository while keeping secrets and site-specific values outside version control.
- Preserve a direct HTTP development launch for local work and the existing WSGI import target.
- Provide defense in depth for host validation, body limits, response headers, least-privilege service execution, and restart behavior.

**Non-Goals:**

- Multi-host high availability, horizontal scaling, zero-downtime rolling deploys, or a general-purpose deployment framework.
- Changing application routes, authentication semantics, templates, database schema, or the daily notification/backup workflow except where deployment configuration must be shared.
- FastAPI or other framework rewrites, PostgreSQL or another database replacement, and containers/Kubernetes.
- Automating DNS registration, router port forwarding, firewall policy, OS package installation, or certificate-authority accounts; these remain documented operator prerequisites.
- Supporting arbitrary proxy chains, CDN proxying, or a second load balancer in front of Caddy. Such a topology requires a separate trust-model change.

## Decisions

### 1. Use Caddy → loopback Gunicorn → existing `webapp.app`

Production uses Caddy on ports 80/443 and Gunicorn on a fixed loopback endpoint such as `127.0.0.1:8080`. Gunicorn imports `webapp:app`, preserving the existing global app and avoiding application restructuring. The systemd service starts Gunicorn from the project virtual environment and never invokes `app.run()`.

Caddy is selected over embedding TLS in Gunicorn because it provides automatic certificate issuance/renewal, HTTP-to-HTTPS redirect behavior, modern TLS defaults, and concise configuration. Gunicorn is selected over Werkzeug because it is a production WSGI server designed for supervision and timeout handling. Binding Gunicorn to loopback rather than all interfaces makes the reverse proxy the only network ingress even if firewall rules drift.

Alternatives considered:

- Werkzeug directly: rejected because Flask explicitly treats it as a development server.
- Nginx plus Certbot: viable, but adds a separate certificate-renewal component and more operational configuration for this single-site Pi.
- Unix socket: offers a similarly narrow boundary, but loopback is easier to inspect and troubleshoot across Caddy/systemd permissions. A future switch to a Unix socket would not change the capability contract.

### 2. Run one synchronous Gunicorn worker

Use one Gunicorn worker with conservative timeout, graceful-timeout, keep-alive, and request-line/header bounds, and let systemd restart failed processes. Access and error logs go to stdout/stderr so journald is the operational log source. The production dependency is isolated in a production requirements file that includes the base requirements and pins a compatible bounded Gunicorn version.

One worker avoids making the current in-memory authentication rate limiter inconsistent across processes and limits SQLite write contention. It also matches the expected low traffic and memory budget of a Raspberry Pi. The trade-off is serialized request handling; this is acceptable because endpoints are short-lived and no streaming or long-running web requests are expected. Increasing workers is not a supported tuning knob until rate limiting is shared or redesigned and SQLite concurrency is re-evaluated.

Alternatives considered:

- Multiple workers: rejected for now because rate limits would be per worker and concurrent SQLite writes would increase.
- Threaded workers: rejected because the current rate-limit state has no explicit synchronization and expected traffic does not require it.
- A separate WSGI module: unnecessary while `webapp.app` remains the stable import target.

### 3. Add an explicit production mode while preserving local launch behavior

Introduce a documented environment selector, `NULIGAHELPER_ENV=production`; absence means the current local/development behavior. Production startup validates `NULIGAHELPER_SECRET`, `NULIGAHELPER_DB`, and a comma-separated `NULIGAHELPER_TRUSTED_HOSTS` containing exact hostnames. Empty values, wildcard hosts, URL schemes, paths, and malformed host entries fail startup. The environment file supplies an absolute database path under a service-writable state directory such as `/var/lib/nuligahelper`.

`run_webapp.sh` remains a developer launcher and may continue to load `.nuligahelper_secret` and run `python webapp.py` over local HTTP. Its documentation must clearly label it non-production. The `__main__` development server can retain network behavior useful on a trusted LAN, but production documentation must never invoke it.

Production configuration sets:

- `SESSION_COOKIE_SECURE=True`
- `SESSION_COOKIE_HTTPONLY=True`
- `SESSION_COOKIE_SAMESITE="Lax"`
- `SESSION_COOKIE_PATH="/"`
- no `SESSION_COOKIE_DOMAIN` by default, keeping the cookie host-only
- the existing one-hour `PERMANENT_SESSION_LIFETIME` and refresh-on-request behavior
- a 1 MiB `MAX_CONTENT_LENGTH`, matching the edge limit
- exact trusted-host enforcement, preferably through the Flask/Werkzeug-supported host-validation facility with a compatible minimum Flask version

Alternatives considered:

- Always-secure cookies: rejected because browsers would not return them to the preserved local HTTP launcher.
- Inferring production only from forwarded HTTPS: rejected because security configuration must not be selected by attacker-controlled request data.
- Putting the secret in the unit file: rejected because repository examples and world-readable unit metadata must not contain credentials.

### 4. Trust exactly one local proxy hop and canonicalize forwarding headers at Caddy

Caddy is configured to discard inbound `Forwarded` and `X-Forwarded-*` values relevant to scheme, host, and client identity, then set a single canonical client address, HTTPS scheme, and original validated host for the backend. It proxies to the fixed loopback endpoint and does not enable trust for upstream CDN/private proxy ranges.

In production, a small WSGI boundary validates the raw socket peer before proxy correction: the peer must be the configured loopback address, and each required forwarded value must be present, single-valued, and syntactically valid. It rejects comma-separated/multi-hop client chains and non-HTTPS schemes. After validation, Werkzeug `ProxyFix` is configured for exactly one `X-Forwarded-For`, `X-Forwarded-Proto`, and `X-Forwarded-Host` hop, with no prefix trust and no broader counts. Host allowlisting then applies to the corrected host. Middleware ordering must preserve access to the raw peer and headers for validation before `ProxyFix` mutates the WSGI environment.

This yields a simple threat model: host-local privileged processes and Caddy are trusted; arbitrary internet headers and direct non-loopback peers are not. A host-local process with permission to connect to the backend is outside the HTTP proxy threat boundary, but still must supply the exact canonical metadata and an allowed host.

Alternatives considered:

- Applying `ProxyFix` unconditionally: rejected because hop counts alone do not verify the immediate peer or header shape.
- Trusting RFC `Forwarded` plus all `X-Forwarded-*` variants: rejected to avoid ambiguous precedence.
- Parsing the left-most address in a chain: rejected because the supported topology has exactly one proxy and therefore needs no chain.

### 5. Enforce host and body limits at both relevant layers

Caddy only serves the configured site address and rejects request bodies over 1 MiB before proxying. Flask independently applies exact trusted-host validation and the same 1 MiB maximum. A 413 handler returns a normal application-shaped response where appropriate, but endpoint code must not execute for oversized bodies.

One MiB is deliberately generous for the current small forms and JSON assignment operations while preventing unbounded buffering. The value is documented as a single production constant and kept equal in Caddy and Flask examples. If uploads are introduced later, their requirement must revisit both limits rather than silently raising only one.

Host validation is duplicated because Caddy site matching protects the edge while application validation protects URL generation and accidental local/bypass requests. Only the canonical public hostname is required; aliases such as `www` are separate explicit allowlist entries and certificate names.

### 6. Apply a compatible security-header baseline in production

Production application responses receive:

- `Strict-Transport-Security: max-age=31536000` after HTTPS is established; do not initially include `includeSubDomains` or `preload` because the deployment does not control or attest every subdomain.
- `Content-Security-Policy: default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline'; img-src 'self' data:; object-src 'none'; base-uri 'self'; frame-ancestors 'none'; form-action 'self'`
- `X-Content-Type-Options: nosniff`
- `X-Frame-Options: DENY` as defense for older clients alongside CSP `frame-ancestors`
- `Referrer-Policy: strict-origin-when-cross-origin`
- `Permissions-Policy` disabling unused capabilities such as camera, microphone, and geolocation

Inline scripts remain forbidden. `'unsafe-inline'` is limited to styles because existing templates use dynamic style attributes for team colors and statistics bars. Removing those attributes with classes or nonces is optional future hardening, not part of this change.

Caddy also sets the applicable header baseline on responses it generates itself (redirects, body-limit errors, and proxy failures), replacing rather than appending duplicate values. Application tests cover app-produced headers; static/config checks and deployment smoke tests cover edge-produced responses.

Alternatives considered:

- Headers only in Caddy: rejected because bypass/offline application tests would not exercise the baseline and app responses would rely entirely on external configuration.
- CSP report-only: rejected because current resource use can support an enforcing policy with only the documented style concession.
- HSTS preload: rejected because it is difficult to reverse and requires broader domain-level commitments.

### 7. Provide versioned examples with site-local configuration outside Git

Add deployment assets under a dedicated `deploy/` directory: a Gunicorn configuration or documented arguments, a systemd unit example, a Caddyfile example, and an environment-file template containing placeholders only. Add a dedicated production deployment guide and link it from `README.MD`, replacing the current statement that production serving remains future work while retaining the local-launch section.

The service runs as a dedicated unprivileged `nuligahelper` account with a restrictive umask. The checkout and virtual environment are readable/executable by that account; the SQLite directory is the narrowly writable state path. The environment file is root-owned and mode 0600 (or equivalently restricted), and Caddy keeps its own certificate state under its packaged service account. The unit uses practical systemd hardening such as `NoNewPrivileges`, private temporary storage, protected home/system paths, a narrow `ReadWritePaths`, and bounded restart timing, provided each directive is supported on the documented Raspberry Pi OS baseline.

The daily job must receive the same `NULIGAHELPER_SECRET` and `NULIGAHELPER_DB`. Deployment instructions explicitly account for SQLite backup consistency, file ownership, stopping/restarting the web service around rollback when necessary, and avoiding two divergent database files.

### 8. Validate behavior offline and separate it from deployment-time network checks

Add synthetic tests that instantiate the app in both local and production modes without reading real `config.json`. Tests cover startup validation, cookie attributes, exact host acceptance/refusal, raw-peer and forwarding-header cases, effective client IP/scheme, 413 behavior, and application security headers. Environment changes are isolated per test because the module-level `app` is eager; tests should exercise the factory with monkeypatched synthetic settings and avoid production settings leaking between imports/tests.

Static tests can inspect deployment examples for required directives and absence of known secret values. When Caddy and systemd binaries are already installed, the guide provides `caddy validate` and `systemd-analyze verify` commands; these are optional host-level checks rather than assumptions of the offline Python suite. Gunicorn import/startup can be checked locally with a bounded command and synthetic temporary database, without binding a public port.

Certificate issuance, public DNS, firewall/router behavior, and external HTTPS reachability are explicit deployment smoke tests and are not run by the offline suite.

## Risks / Trade-offs

- [Single worker limits throughput and a slow request can delay others] → Keep web handlers short, configure bounded Gunicorn timeouts, observe journald, and revisit concurrency only together with shared rate limiting and SQLite analysis.
- [SQLite is accessed by both the daily job and web service] → Use one absolute shared database path, preserve existing short transactions, document ownership/backup coordination, and do not add worker concurrency in this change.
- [Incorrect proxy-header ordering can reintroduce spoofing] → Validate the raw peer and exact header shape before one-hop correction; test spoofed, missing, comma-separated, and direct-request cases offline.
- [A local process can connect to the loopback backend] → Treat the host boundary as trusted, run only necessary services, use systemd hardening, and consider a permissioned Unix socket in a future hardening change if local multi-user isolation becomes necessary.
- [CSP may block a resource introduced later] → Keep the policy explicit and covered by page smoke tests; update it narrowly with any intentional new resource rather than adding broad wildcards or inline-script permission.
- [HSTS can make an HTTPS outage more visible] → Enable it only on the production HTTPS path after certificate readiness; avoid preload/includeSubDomains so recovery remains scoped.
- [Automatic certificates require working public DNS and ports 80/443] → Document prerequisites and staging checks; keep certificate issuance out of offline tests.
- [systemd hardening differs across Raspberry Pi OS releases] → Target and document a baseline, verify the unit on-host, and prefer supported directives over an untested maximal sandbox.
- [Environment-file secrets may be exposed by weak permissions or operational copying] → Require restrictive ownership/mode, placeholders in Git, and redaction in support output; never log the secret.
- [The eager module-level app validates configuration at import time] → Keep production environment complete before Gunicorn imports `webapp:app`, and structure tests around the factory rather than introducing an unnecessary WSGI refactor.

## Migration Plan

1. Back up the current SQLite database and existing secret/configuration; record the currently deployed revision and local launch command.
2. Install/upgrade the project virtual environment from the base and production requirements without changing the database schema.
3. Create the dedicated service account, state directory, production environment file, and shared absolute database path with least-privilege ownership. Confirm the daily job uses the same path and secret.
4. Render site-local copies of the example systemd and Caddy configurations with the real checkout path and public hostname. Keep those site-local values and secrets out of Git.
5. Run the full offline project suite plus available Gunicorn import, systemd unit, and Caddy syntax validation.
6. Stop any manually launched Werkzeug process. Start and enable the Gunicorn systemd service, confirm only the loopback listener exists, and test it through a local Caddy request with the configured host.
7. Confirm DNS and inbound ports, then enable/reload Caddy and verify HTTP redirect, certificate validity, trusted-host refusal, cookie flags, response headers, client-address logging/rate limiting, normal login, and a representative assignment mutation.
8. Monitor both units and application behavior in journald, then update the operational record with backup and recovery locations.

Rollback:

1. Disable or stop Caddy exposure first if the production boundary is malfunctioning.
2. Stop the Gunicorn service and restore the previous application revision/configuration; restore the database only when data rollback is explicitly required, using the pre-deploy backup and with writers stopped.
3. For trusted-LAN emergency access only, use the documented development launcher; do not expose Werkzeug publicly and do not present this as a production rollback state.
4. Re-enable the prior known-good Gunicorn/Caddy configuration once syntax and local health checks pass. Preserve the stable secret unless deliberate session invalidation is intended.
