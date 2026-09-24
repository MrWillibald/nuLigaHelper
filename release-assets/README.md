# Reviewed production release assets

This directory contains generic, non-secret files required by a clean checkout.
The installed systemd services use `/opt/nuligahelper/current` for all code,
virtualenv, and Gunicorn paths. The application root `/opt/nuligahelper` remains
a directory; `current` is a root-owned symlink to one complete release beneath
that directory.

The tracked `nuligahelper-deploy.py` currently implements **prepare and inspect
only**. It can pin a commit reachable from fetched `master`, export a clean
source archive into `releases/<sha>`, build an isolated virtual environment,
run the offline suite and syntax checks, and record the previous release. It
cannot activate a release. Do not install it as the production cutover tool
until the activation, snapshot, timer, rollback, and rehearsal tasks in the
OpenSpec change are complete. Its root-only configuration template is
`deployment.json.example`; the real `/etc/nuligahelper/deployment.json` must
be root-owned mode 0600 and must not enter Git. The Git source cache is outside
the service-readable application tree.

`recovery_check.py` provides the candidate-runtime SQLite snapshot and
read-only schema classification primitives for the future activation step. It
does not quiesce writers, close ingress, or authorize a migration. Running it
alone is not a safe cutover procedure.

The files in `systemd/` and `nuligahelper-watchdog.cron` are installation
templates. Install reviewed copies as root, verify them with
`systemd-analyze verify`, and run `systemctl daemon-reload` before starting a
changed service. Do not switch the link while the web, daily, or cleanup
services are using the database.
The first installed web unit still points to `current/deploy/gunicorn.conf.py`
inside the legacy release. Installing the new web unit, which points to
`current/release-assets/gunicorn.conf.py`, before a compatible release is active
would stop the web service. The cutover runbook must handle that transition
explicitly.

Site-specific Caddy configuration, `/etc/nuligahelper/web.env`, club texts,
contacts, credentials, the SQLite database in `/var/lib/nuligahelper`, snapshots, and operator approval
records are **not** release assets. Keep them outside Git and outside every
release directory. The `.example` files contain placeholders only; never
install them without supplying the private, approved site values.

The daily and cleanup timers are persistent. Starting either after a missed
09:00 Europe/Berlin firing can immediately run its service, including real
notifications or cleanup. Review the last trigger and job result before
resuming them. Never run the daily service merely as a deployment smoke test.
