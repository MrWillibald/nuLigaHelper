# Production deployment (Raspberry Pi OS / Linux)

This is the supported public-serving topology:

```text
Internet :80/:443 -> Caddy -> 127.0.0.1:8080 -> one Gunicorn worker -> webapp:app
```

The instructions target a current, 64-bit Raspberry Pi OS or Debian-family Linux
release with systemd. Run administrative commands as root (or through `sudo`).
Commands containing `example.invalid`, `/opt/nuligahelper`, or another example path
must be adapted to the host. Do not expose `run_webapp.sh` or Werkzeug publicly.

## 1. Prerequisites

Before installation, arrange all of the following:

- a public DNS A/AAAA record for one exact hostname pointing to this host;
- inbound TCP ports 80 and 443 forwarded and permitted by the host/router firewall;
- correct system time and outbound HTTPS/DNS access for certificate issuance;
- Python 3.12 or a project-supported Python 3 version, `venv`, Git, SQLite, and Caddy
  from the distribution or Caddy's official Debian repository;
- a local filesystem for SQLite (not NFS or another network filesystem);
- the existing read-only `config.json` containing the club's mail/SMS/Dropbox
  settings. It remains outside Git and is required by notification operations.

Install the OS packages using the distribution's documented package workflow. The
Python production dependency is Gunicorn and is installed from the repository's
`requirements-production.txt`; it is not necessary to install a separate OS
Gunicorn package.

Check DNS before claiming public readiness:

```bash
getent ahosts nuliga.example.invalid
sudo ss -ltnp '( sport = :80 or sport = :443 )'
timedatectl status
```

If DNS does not resolve to the public address, port forwarding is absent, or ports
80/443 are blocked, Caddy cannot complete public certificate issuance. Diagnose
with `journalctl -u caddy` and the DNS/router provider before continuing.

## 2. Dedicated account, checkout, state, and virtual environment

Create one non-login account and narrowly scoped directories:

```bash
sudo useradd --system --home-dir /nonexistent --shell /usr/sbin/nologin nuligahelper
sudo install -d -o root -g nuligahelper -m 0750 /opt/nuligahelper
sudo install -d -o nuligahelper -g nuligahelper -m 0700 /var/lib/nuligahelper
sudo install -d -o root -g root -m 0700 /etc/nuligahelper
```

Clone or copy the reviewed checkout into `/opt/nuligahelper`. Keep the checkout and
virtual environment owned by root and readable/executable by the service group;
only `/var/lib/nuligahelper` is writable by the service account. For example:

```bash
sudo git clone REPLACE_WITH_REPOSITORY_URL /opt/nuligahelper
sudo chown -R root:nuligahelper /opt/nuligahelper
sudo chmod -R g+rX,o-rwx /opt/nuligahelper
sudo python3 -m venv /opt/nuligahelper/venv
sudo /opt/nuligahelper/venv/bin/pip install -r /opt/nuligahelper/requirements-production.txt
sudo chown -R root:nuligahelper /opt/nuligahelper/venv
sudo chmod -R g+rX,o-rwx /opt/nuligahelper/venv
```

Copy the site-specific, gitignored `config.json` to `/opt/nuligahelper/config.json`,
then set it to `root:nuligahelper` mode 0640. Do not put credentials in a unit or
Caddyfile.

## 3. Persistent environment and shared SQLite identity

Install the example and edit every placeholder in the external file:

```bash
sudo install -o root -g root -m 0600 \
  /opt/nuligahelper/deploy/nuligahelper.env.example \
  /etc/nuligahelper/web.env
sudoedit /etc/nuligahelper/web.env
```

The finished file must contain the intended production values:

```text
NULIGAHELPER_ENV=production
NULIGAHELPER_SECRET=REPLACE_WITH_A_STABLE_RANDOM_SECRET
NULIGAHELPER_DB=/var/lib/nuligahelper/nuliga_helper.db
NULIGAHELPER_TRUSTED_HOSTS=nuliga.example.invalid
```

Generate the secret once with a cryptographically secure generator and paste it
using `sudoedit`; do not print it in diagnostics, shell tracing, process arguments,
or support logs. Keep `/etc/nuligahelper/web.env` root-owned and mode 0600. The
secret must remain stable: rotating it invalidates every browser session. It also
changes every keyed authentication-abuse subject, effectively starting new limits
while old opaque counter rows age out. Treat rotation as a security event, not as a
way to clear a throttle.

Authentication abuse controls use safe built-in defaults. To tune them, add one
single-line `NULIGAHELPER_AUTH_ABUSE_CONFIG` JSON object to this root-owned file;
the complete sample and every accepted key are in `config_template.json` under
`club.auth_abuse`. Never put contacts, names, IP addresses, credentials, or the
HMAC secret in this JSON. Startup refuses unknown policies, wrong types,
non-positive limits/windows, inconsistent proxy settings, or retention shorter
than the longest configured window.

Each `policies` entry has a positive `limit` and `window_seconds`. Login and
registration have independent client, canonical-contact, and resolved-person
policies; confirmation has client/contact/person policies. `sms_contact_cap`,
`sms_person_cap`, and `sms_global_cap` are additional SMS-only cost bounds shared
by login and registration. E-mail never consumes SMS-cap rows. Windows are aligned
fixed UTC epoch buckets, including the default 24-hour SMS operational period; they
are not rolling intervals or local-midnight days. A fixed window can therefore
admit a boundary burst across adjacent buckets.

`retention_seconds` is padding after a bucket ends and must cover the longest
policy window. `cleanup_batch` bounds deletions in one startup/opportunistic run,
and `cleanup_interval_seconds` limits how often request traffic tries cleanup.
Cleanup is indexed, preserves live rows, skips lock contention, and resumes on a
later request or restart. Active counters persist in the shared SQLite database
across process restarts and workers. Evaluation uses a dedicated, short SQLite
writer transaction and fails closed on busy/storage errors before token, account,
session, or delivery side effects. An allowance stays consumed if a provider later
fails; this conservative behavior bounds sends and spending.

`trusted_proxies`, `trusted_hops`, and `proxy_error` control optional direct WSGI
forwarded-client parsing (`fallback` uses the trusted direct peer; `refuse` rejects
attribution). Keep `trusted_proxies=[]` and `trusted_hops=0` in the supported Caddy
topology: `ProductionProxyBoundary` already accepts only loopback, requires exactly
one replacement `X-Forwarded-For` value, and passes the verified address through
`ProxyFix` before the limiter sees it. Do not enable a second attribution layer.
For another topology, first document and test the exact Flask bind address, every
trusted peer CIDR, replacement—not append—header behavior, and exact hop count.
Keep proxy-derived attribution disabled and public exposure blocked until network
reachability and application trust configuration agree.

Monitor `nuligahelper.security` events for `auth_abuse_throttled`,
`auth_abuse_global_sms_cap`, `auth_abuse_storage_error`, `auth_abuse_cleanup`, and
`auth_delivery_failed`. They contain action, channel, coarse dimensions, counts or
stable exception classes, never limiter digests or personal/authentication data.
Choose caps from expected member/game-day traffic and current provider pricing;
raise them deliberately only after reviewing these events. A normal reset is to
wait for the fixed bucket to roll over. For an emergency manual global-cap reset,
stop public ingress and the web service, make a SQLite-consistent backup, delete
only the current `action='sms_delivery' AND dimension='global'` row in a maintenance
transaction, restart, and repeat health checks. Never delete arbitrary auth rows
or recreate the database to clear a cap.

Before enabling SMS publicly, use the actual Twilio Console/account documentation
to configure supported Usage Triggers or alerts, restrict geographic permissions
to required destinations, and verify an account/project spending limit or prepaid
balance protection where that account and region provide one. Console capabilities
and names vary by account type and region, so confirm controls by observation and
alert testing rather than assuming a feature exists. Store no Twilio credential in
tracked files. Application reservations bound calls initiated by this app, but
Twilio-side alerts/spending controls are the independent final cost backstop.

The daily job and web service **must load the same environment file**, giving them
the identical stable `NULIGAHELPER_SECRET` and absolute `NULIGAHELPER_DB`. A second,
relative, or default database path would create divergent SQLite files. A cron
wrapper may source `/etc/nuligahelper/web.env` with automatic export before it
executes `/opt/nuligahelper/run.sh`; keep that wrapper root-owned and unreadable to
other users. Confirm paths and ownership without displaying values:

```bash
sudo test -s /etc/nuligahelper/web.env
sudo test "$(stat -c '%U:%G %a' /etc/nuligahelper/web.env)" = "root:root 600"
sudo -u nuligahelper test -r /opt/nuligahelper/config.json
sudo -u nuligahelper test -w /var/lib/nuligahelper
```

## 4. Render and validate service configuration

Review `/opt/nuligahelper/deploy/nuligahelper-web.service`. If the checkout or state
paths differ, copy it and replace every path before installation:

```bash
sudo install -o root -g root -m 0644 \
  /opt/nuligahelper/deploy/nuligahelper-web.service \
  /etc/systemd/system/nuligahelper-web.service
sudo systemd-analyze verify /etc/systemd/system/nuligahelper-web.service
```

Copy the Caddy example, replace every `nuliga.example.invalid` occurrence with the
same exact lowercase hostname used by `NULIGAHELPER_TRUSTED_HOSTS`, then validate it:

```bash
sudo install -o root -g root -m 0644 \
  /opt/nuligahelper/deploy/Caddyfile.example /etc/caddy/Caddyfile
sudoedit /etc/caddy/Caddyfile
sudo caddy validate --config /etc/caddy/Caddyfile --adapter caddyfile
```

These syntax checks are offline when the binaries are already installed. They do
not query DNS, request a certificate, or prove public reachability. A bounded WSGI
import/start smoke test can run without a public listener:

```bash
sudo -u nuligahelper timeout 10s sh -c \
  'set -a; . /etc/nuligahelper/web.env; set +a; cd /opt/nuligahelper; exec venv/bin/gunicorn --check-config --config deploy/gunicorn.conf.py webapp:app'
```

## 5. Start order and deployment-time verification

Load and start the loopback backend first, then the public edge:

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now nuligahelper-web.service
sudo systemctl status nuligahelper-web.service --no-pager
sudo ss -ltnp '( sport = :8080 )'
sudo systemctl enable --now caddy.service
sudo systemctl reload caddy.service
sudo systemctl status caddy.service --no-pager
```

The `ss` output must show only `127.0.0.1:8080`, never `0.0.0.0:8080`, `[::]:8080`,
or a LAN/public address. Use bounded local checks after substituting the hostname:

```bash
# Backend accepts exactly one canonical loopback-proxy hop.
curl --fail --silent --show-error --max-time 5 \
  -H 'X-Forwarded-For: 203.0.113.10' -H 'X-Forwarded-Proto: https' \
  -H 'X-Forwarded-Host: nuliga.example.invalid' \
  http://127.0.0.1:8080/ -o /dev/null

# Wrong corrected host is refused by the application.
test "$(curl --silent --output /dev/null --write-out '%{http_code}' --max-time 5 \
  -H 'X-Forwarded-For: 203.0.113.10' -H 'X-Forwarded-Proto: https' \
  -H 'X-Forwarded-Host: other.example.invalid' \
  http://127.0.0.1:8080/)" = 400

# Missing canonical forwarding headers cannot bypass the proxy boundary.
test "$(curl --silent --output /dev/null --write-out '%{http_code}' --max-time 5 \
  http://127.0.0.1:8080/)" = 400
```

The following checks exercise Caddy and certificate/public-network state; they are
deployment-time checks, not part of the offline suite:

```bash
# HTTP redirects to HTTPS. --resolve allows a same-host edge check once Caddy runs.
curl --head --silent --show-error --max-time 10 \
  --resolve nuliga.example.invalid:80:127.0.0.1 \
  http://nuliga.example.invalid/

# Certificate, application, and response headers through the public boundary.
curl --head --fail --show-error --max-time 15 https://nuliga.example.invalid/

# Oversized ingress is rejected with 413.
head -c 1048577 /dev/zero | curl --silent --output /dev/null \
  --write-out '%{http_code}\n' --max-time 15 --data-binary @- \
  https://nuliga.example.invalid/login

# Edge-generated missing route/error responses retain the security baseline.
curl --head --silent --show-error --max-time 10 \
  https://nuliga.example.invalid/__missing_health_probe__
```

Inspect these responses for HSTS, CSP, content-type, frame, referrer, and permissions
headers. To verify forwarding-header overwrite, send a harmless spoofed
`X-Forwarded-For` value through Caddy and confirm the request succeeds without a
proxy-boundary error. Application security logs intentionally omit client addresses:

```bash
curl --silent --output /dev/null --max-time 10 \
  -H 'X-Forwarded-For: 192.0.2.99' https://nuliga.example.invalid/
sudo journalctl -u nuligahelper-web.service --since '-2 minutes' --no-pager -n 50
```

Complete a browser login and one representative assignment claim/release through
HTTPS. Confirm the session cookie is Secure, HttpOnly, SameSite=Lax, Path=/, and has
no Domain attribute. Certificate issuance and public reachability are successful
only after a normal browser/curl trusts the certificate and external clients can
reach the hostname. Continue supervising renewal and failures with:

```bash
sudo journalctl -u caddy.service -u nuligahelper-web.service --since today --no-pager
sudo systemctl is-active caddy.service nuligahelper-web.service
```

## 6. Updates and SQLite-consistent backup

Record the current revision and create a consistent snapshot before changing code.
Do not copy only the live WAL-mode main file. Either use the successful dated/latest
Dropbox snapshot from the daily job or use SQLite's online backup mechanism while
writers are running. For a maintenance-window filesystem copy, stop both writers
and preserve the complete `.db`, `-wal`, and `-shm` set.

An ordered update is:

1. Confirm the latest validated snapshot and record `git rev-parse HEAD` without
   printing environment values.
2. Fetch and check out the reviewed revision as the checkout owner.
3. Reapply `root:nuligahelper` ownership and read/execute permissions; verify the
   state directory remains the only service-writable path.
4. Run `venv/bin/pip install -r requirements-production.txt`.
5. Run `test/run_tests.sh`, the Gunicorn check, `systemd-analyze verify`, and
   `caddy validate` before restart.
6. Restart `nuligahelper-web.service`, reload Caddy only if its configuration
   changed, and repeat listener, HTTPS, login, and assignment health checks.

Useful diagnostics are bounded journal queries, not environment dumps:

```bash
sudo systemctl restart nuligahelper-web.service
sudo journalctl -u nuligahelper-web.service --since '-10 minutes' --no-pager -n 200
sudo systemctl reload caddy.service
sudo journalctl -u caddy.service --since '-10 minutes' --no-pager -n 200
```

## 7. Rollback

On a broken public boundary, stop public ingress first:

```bash
sudo systemctl stop caddy.service
sudo systemctl stop nuligahelper-web.service
```

Restore the prior application revision and site-local configuration, reinstall its
production requirements, validate all assets, and restart Gunicorn before Caddy.
Preserve `/var/lib/nuligahelper/nuliga_helper.db` by default. Restore database data
only when a data rollback was explicitly chosen, with every database writer stopped,
using `manage_db.py restore-snapshot` and the documented guarded restore workflow in
`README.MD`. Never delete/recreate the production database as a deployment rollback.

The additive `auth_abuse_counters` table may remain unused after a code rollback;
do not remove it as part of rollback. The old version restores process-local limits
that reset on restart, so it is never an acceptable long-term publicly exposed
state. Diagnose and redeploy the fixed version before restoring ingress.

After local loopback health passes, start Caddy and repeat the public checks. Keep
the stable secret unless deliberate session invalidation is intended. Werkzeug may
be used only for emergency access on a trusted LAN; it is never a public rollback
runtime.
