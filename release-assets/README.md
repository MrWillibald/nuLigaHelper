# Reviewed production release assets

This directory contains generic, non-secret files required by a clean checkout.
The installed systemd services use `/opt/nuligahelper/current` for all code,
virtualenv, and Gunicorn paths. The application root `/opt/nuligahelper` remains
a directory; `current` is a root-owned symlink to one complete release beneath
that directory.

The tracked `nuligahelper-deploy.py` implements separate `inspect`, `prepare`,
`activate`, `continue`, `resume-timers`, and `rollback-code` commands. It pins a
commit reachable from fetched `master`, builds an isolated release, and checks
it before interruption. Activation then stops known writers and ingress,
creates a validated SQLite snapshot, gates the schema, switches `current`, and
checks local and public readiness. **The command is still a draft:** do not
install or run its activation commands on production until the representative
host rehearsal, runbook review, and operator approval in the OpenSpec change
are complete. Its root-only configuration template is
`deployment.json.example`; the real `/etc/nuligahelper/deployment.json` must
be root-owned mode 0600 and must not enter Git. The Git source cache is outside
the service-readable application tree. Its `database` must name the existing
`/var/lib/nuligahelper/nuliga_helper.db`; `recovery_dir` is a separate
root-owned mode 0700 directory for future cutover records and snapshots, and
`public_health_url` is the approved HTTPS `/healthz` URL. These additional
settings do not enable activation in the current command.

`recovery_check.py` provides the candidate-runtime SQLite snapshot, validation,
and read-only schema classification primitives. It does not quiesce writers,
close ingress, or authorize a migration. Running it alone is not a safe cutover.

The commands below show the intended order after those outstanding approvals.
Install the reviewed deploy script outside `/opt/nuligahelper` as root, and
replace `<full-sha>` and `<deployment-id>` with the pinned identifiers printed
by the command. Do not paste environment values into commands or records.
The host needs `git`, `python3-venv`, `lsof`, `iproute2` (`ss`),
`systemd-analyze`, and `sqlite3` installed before rehearsal.

```bash
sudo /usr/local/sbin/nuligahelper-deploy inspect
sudo /usr/local/sbin/nuligahelper-deploy prepare --sha <full-sha>
sudo /usr/local/sbin/nuligahelper-deploy inspect --sha <full-sha>
sudo /usr/local/sbin/nuligahelper-deploy activate --sha <full-sha>
```

`activate` deliberately returns exit 2 and leaves writers/ingress stopped if
the existing database needs a recognized schema upgrade. The operator must
review the candidate migration, run that release's existing guarded
`manage_db.py migrate-schema --confirm-stopped` against the approved database,
verify its postflight, then use `continue --deployment-id <deployment-id>`.
For example, while all writers remain stopped:

```bash
sudo -u nuligahelper /opt/nuligahelper/releases/<full-sha>/venv/bin/python -B \
  /opt/nuligahelper/releases/<full-sha>/manage_db.py \
  --db /var/lib/nuligahelper/nuliga_helper.db migrate-schema --confirm-stopped
sudo /usr/local/sbin/nuligahelper-deploy continue --deployment-id <deployment-id>
```

An unknown, corrupt, divergent, or newer schema is a refusal,
not an invitation to migrate. Do not initialize a replacement database.

On successful public readiness, the daily and cleanup timers remain disabled.
Inspect the printed `pending_catchup` values and effects of a missed 09:00
run. `resume-timers --deployment-id <deployment-id> --catchup hold` records a
decision to leave timers paused; `--catchup run` records acceptance *before*
restoring their prior state and may immediately send messages or perform
cleanup. Neither command invokes the daily service as a smoke test.

`rollback-code --deployment-id <deployment-id>` is permitted only before any
possible public write and only when the previous release accepts the unchanged
database schema. It never restores a database. A migrated database or one that
may have accepted public writes requires a separate, stopped-writer recovery
decision and the existing validated restore procedure; never blindly replace
the live database with an older snapshot.

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
