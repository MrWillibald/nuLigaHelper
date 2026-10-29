# Private deployment

For repeat updates to an existing installation, follow the tracked
[deployment guide](../../DEPLOYMENT.md). It uses a pinned commit on protected `master`
and the installed release command. The first production cutover has been
accepted.

The operator-triggered release workflow has tracked, non-secret
templates in `release-assets/`. Its deploy command now implements separate
preparation, guarded activation, explicit timer resumption, and conditional
code-only rollback. Activation and compatible rollback wait up to 60 seconds
for Gunicorn's loopback listener and local `/healthz` after the `Type=simple`
web service starts; Caddy and timers remain closed if readiness fails or times
out. Host preflight and the first-cutover operator review in the private
operations runbook are complete, so these commands are approved for production
releases from protected `master`. `release-assets/README.md` explains the
operator command sequence and recovery gates for setting up a host.
The production database remains `/var/lib/nuligahelper/nuliga_helper.db`, and
`/etc/nuligahelper/web.env` remains private and shared by the services.

The application exposes a fixed `/healthz` readiness route and public
`/impressum` and `/datenschutz` routes. Legal-page content, privacy retention
settings, host monitoring, service units and launch evidence are supplied by the
private deployment. Missing or invalid legal content returns a generic 503 page.
The private deployment uses a short cleanup policy (five retention periods and
batch size), two public legal pages, and a launch record with one club officer's
decision. The launch record checks the restore rehearsal, privacy notice and
retention settings; technical host checks run separately.

Production uses a stable `NULIGAHELPER_SECRET`, an absolute `NULIGAHELPER_DB`
shared by the web app, daily job, cleanup and backup,
exact `NULIGAHELPER_TRUSTED_HOSTS`, and provider settings from environment
variables. Production `config.json` contains club, recipient and operational settings,
not email, Twilio or Dropbox provider sections. Host setup and public launch
require separate operator review.

The local development overrides `NULIGAHELPER_TODAY` and `NULIGAHELPER_DRY_RUN` must
not appear in `/etc/nuligahelper/web.env` or any unit environment: `runtime_check.py`,
web startup, the daily job and the operations preflight/launch check refuse to run
and name the variable.

Authentication does not make the Flask development server suitable for public use.
Generic Gunicorn and systemd templates are tracked in `release-assets/`;
hostnames, credentials, environment files, and installed host configuration
remain private and outside the release source.


---
Back to the [README](../../README.MD) · [Operations index](README.md)
