# Deploy a reviewed release to an existing installation

This is the repeatable update procedure for the existing production layout:
`/opt/nuligahelper/current` selects a complete release, while the live SQLite
database remains at `/var/lib/nuligahelper/nuliga_helper.db`. The root-managed
`/usr/local/sbin/nuligahelper-deploy` command prepares and activates releases.
Run these commands on the server over SSH as an operator with `sudo` access.

## Before the first use

This guide assumes the deploy command, its root-owned
`/etc/nuligahelper/deployment.json`, the `current` link, and the compatible
systemd units are installed and have passed the host preflight. The first
production cutover also needs the operator acceptance described in the private
operations runbook. The tracked [release asset guide](release-assets/README.md)
explains the setup and the legacy-release Gunicorn compatibility requirement.
Complete those gates before using `activate` on production.

Keep `/etc/nuligahelper/web.env`, site configuration, secrets, and the database
outside every release. The deploy command reads the approved paths from its
private configuration. Never use an in-place `git pull`, run `manage_db.py init`
against the existing database, or copy only a live SQLite `.db` file while WAL
may contain committed data.

## 1. Select and prepare the release

Choose the **full 40-character SHA** of a reviewed commit on protected `master`.
Record the PR and its test result. The deploy command fetches `master` and
rejects a SHA that is not reachable from it. A merge SHA can differ from the
feature branch's last commit.

Check the current link, service/timer states, database revision, recent backup,
available disk space, and the changes in the selected release. Review any
database migration, dependency, systemd, Caddy, privacy, or notification change
before choosing the maintenance window. In particular, the daily and cleanup
timers fire at 09:00 Europe/Berlin and are persistent.

```bash
RELEASE_SHA='paste-full-40-character-master-commit-sha'
sudo /usr/local/sbin/nuligahelper-deploy inspect
sudo /usr/local/sbin/nuligahelper-deploy prepare --sha "$RELEASE_SHA"
sudo /usr/local/sbin/nuligahelper-deploy inspect --sha "$RELEASE_SHA"
```

`prepare` exports the pinned commit to `/opt/nuligahelper/releases/<sha>`,
builds its own virtual environment, checks dependencies and service syntax,
and runs `test/run_tests.sh`. It does not change `current`, the live database,
or running services. Check its result and the second `inspect` output before
activation. If the candidate already exists, inspect that prepared candidate;
`prepare` deliberately refuses to overwrite it.

When a release changes the deploy command or installed service or proxy assets,
review and install those host changes separately using the private operations
runbook. Verify service syntax and compatibility with the retained prior
release before activation. Application releases do not automatically replace
host files.

When deploying the message-catalog transition, run the candidate release's
`./venv/bin/python -m message_config --config /path/to/private/config.json`
before activation. This read-only command reports legacy keys/customization
status without private text or recipient values. Privately reconcile customized
wording in `messages.py`, move referee recipients to
`club.notifications.referee_targets`, and remove `club.texts`. Keep the previous
configuration with its compatible application release for rollback. See the
[message migration guide](README.MD#notification-wording-and-migration); this
step uses no providers/database and creates no notification trigger.

## 2. Activate in a maintenance window

Keep a second SSH session open for recovery. Verify that the host preflight and
first-cutover acceptance are complete. Activation takes an exclusive lock,
pauses the daily and cleanup timers, waits for active jobs, stops public ingress
and the web service, and checks that no database user remains. It creates a
validated recovery snapshot and checks schema compatibility before switching
`current`. It then waits for local web readiness before reopening Caddy and
checking public HTTPS health. Allow for a short outage.

```bash
RELEASE_SHA='same-full-40-character-sha-used-for-prepare'
sudo /usr/local/sbin/nuligahelper-deploy activate --sha "$RELEASE_SHA"
```

Save the printed `deployment_id`, outcome, previous release, snapshot path,
schema result, and timer states in the private operator record. An outcome of
`public_ready` means the web is serving the new release; the daily and cleanup
timers are still paused.

If the command exits with `migration_required`, it leaves ingress and writers
stopped. Review the candidate migration and its backup implications. Only for
a recognized older revision, run the guarded migration from the **candidate**
release while all database users remain stopped, then continue using the ID
printed by activation:

```bash
RELEASE_SHA='same-full-40-character-sha-used-for-prepare'
DEPLOYMENT_ID='paste-deployment-id-from-activate-output'
sudo -u nuligahelper "/opt/nuligahelper/releases/$RELEASE_SHA/venv/bin/python" -B \
  "/opt/nuligahelper/releases/$RELEASE_SHA/manage_db.py" \
  --db /var/lib/nuligahelper/nuliga_helper.db migrate-schema --confirm-stopped
sudo /usr/local/sbin/nuligahelper-deploy continue --deployment-id "$DEPLOYMENT_ID"
```

Record the migration's backup path and successful postflight. An unknown,
newer, divergent, or damaged schema is a refusal: keep services stopped and
investigate it through the private recovery runbook. Never initialize a
replacement database to make a release start.

For the birth-date revision, postflight must retain person identities, contacts,
memberships, status, MV appointments, assignments and audit snapshots. Existing
dates remain unknown. After application verification, use authorized self/admin
maintenance or `set-birth-date PERSON_ID YYYY-MM-DD` to complete the roster;
review current/future eligibility warnings without deleting existing appointments.
The [data inventory](DATA-INVENTORY.md) records visibility and backup scope.

## 3. Decide when scheduled work resumes

After `public_ready`, record a `hold` decision first. This computes and returns
`pending_catchup` while leaving the timers paused. Review that result and the
prior timer states before deciding whether to resume them. Resuming a
persistent timer after a missed 09:00 run can immediately send notifications
or perform cleanup. Use the ID from activation:

```bash
DEPLOYMENT_ID='paste-deployment-id-from-activate-output'
sudo /usr/local/sbin/nuligahelper-deploy resume-timers \
  --deployment-id "$DEPLOYMENT_ID" --catchup hold
```

When the effects of any catch-up run are accepted, use the same command with
`--catchup run` to restore the timers' prior states. Leaving the decision at
`hold` leaves them paused. Do not start the daily service merely as a smoke
test; it can send real messages and upload a backup.

## 4. Verify and retain recovery material

Confirm the `current` link resolves to the selected SHA and check web/Caddy
health, timer states, and recent journals. Use the site's approved public
health URL from the private deployment configuration:

```bash
readlink -f /opt/nuligahelper/current
systemctl status nuligahelper-web.service caddy.service --no-pager
systemctl list-timers 'nuligahelper-*'
sudo journalctl -u nuligahelper-web -u caddy --since '1 hour ago' --no-pager
curl --fail --max-time 10 'https://your-approved-host.example/healthz'
```

Check the schedule and a read-only database view, including a few assignment
and audit IDs, without putting member data in shared records. Verify the
database is still the approved file and the next monitor check succeeds. Keep
the prior release and validated deployment snapshot until the release is
accepted under the private retention policy.

## If activation fails

Keep ingress and timers closed and inspect the deployment record and journals.
Before public writes can have been accepted, and only when the database schema
is unchanged and compatible with the prior release, the guarded code rollback
can select that prior release while retaining the **current** database:

```bash
DEPLOYMENT_ID='paste-deployment-id-from-activate-output'
sudo /usr/local/sbin/nuligahelper-deploy rollback-code --deployment-id "$DEPLOYMENT_ID"
```

After a schema migration or possible public writes, stop and follow the private
recovery decision path. Do not automatically restore an older snapshot over
new assignments or audit entries. The [database restore and migration guide](README.MD#sqlite-production-profile-and-backups)
describes the guarded SQLite procedures; the private runbook governs the host
and service recovery sequence.
