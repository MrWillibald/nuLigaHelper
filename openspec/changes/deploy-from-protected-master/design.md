## Context

See `proposal.md` for the motivation and `specs/production-deployment/spec.md` for the behavioral contract. Production currently runs the latest `devel/webui` application with a systemd daily job and SQLite at `/var/lib/nuligahelper/nuliga_helper.db`. At planning time the local `devel/webui` head is `519c982fd1c13bcc6d59749aade2c3144d7c57b9`; the actual server commit must be read on the host before its first deployment. The local `master` tracking ref is an ancestor and still contains the old Excel application, so promotion to `master` is a distinct first step.

The existing units use `/opt/nuligahelper`, a shared `/etc/nuligahelper/web.env`, and one service account. The private `deploy/` directory and `requirements-production.txt` are ignored by Git, so a GitHub checkout alone does not contain all current runtime assets. The application already has guarded Alembic migration and validated SQLite snapshot/restore primitives. `add-production-operations-and-privacy` owns service identity, permissions, monitoring, alerts, and broader recovery policy; this change owns release selection and cutover orchestration.

## Goals / Non-Goals

**Goals:**

- Make each deployment identify one pull-request and CI-gated commit and one prepared, independently testable release.
- Keep the current SQLite file, configuration, and secret stable through code updates.
- Keep failure before public reopening recoverable without rebuilding the previous release.
- Make schema changes and timer catch-up visible operator decisions.

**Non-Goals:**

- Automatic deployment on push, GitHub-hosted SSH access to production, or zero-downtime serving.
- Replacing SQLite, Caddy, Gunicorn, systemd, or the existing schema migration and restore algorithms.
- Importing the historical Excel workbook; production already uses SQLite.
- Automatically restoring a database after a release has accepted public writes.

## Decisions

### 1. Promote a verified baseline to protected master

Add an offline pull-request check and configure GitHub branch protection/rules so `master` requires a pull request and the successful `offline-tests` check before merge. This is a solo-maintainer repository: the operator reviews the diff, but the ruleset cannot supply independent human approval and currently requires zero approving reviews. Do not describe this as second-person review enforcement. First promote the application currently running on production from `devel/webui`; compare the server's actual commit, application files, and schema revision with the promoted result. A merge commit may have a different SHA while carrying the same application code. New CI or deployment files also change the repository tree, so the first cutover is migration-free only after checking the actual application and schema differences.

The operator initiates deployment on the VPS. A root-managed read-only GitHub credential is used only when repository access requires one; the application service account has no Git credential. The deploy command fetches `master`, resolves an exact SHA, verifies that any requested SHA is reachable from the fetched branch, and records it. It does not track the moving branch name during the rest of the deployment. GitHub settings enforce the PR and offline-check gates; local ancestry verification does not pretend to prove that a GitHub administrator could not alter those settings.

Automatic GitHub Actions deployment was rejected because the operator wants to choose the release time and because it would place production access in a hosted workflow. In-place `git pull` was rejected because it changes files underneath running processes and provides no prepared prior release.

### 2. Use an internal current link and per-commit release directories

Keep `/opt/nuligahelper` as a root-owned directory with the existing service group and mode. Store immutable candidates under `/opt/nuligahelper/releases/<full-sha>/` and a root-owned `current` symlink pointing to one of those directories. Installed units use `/opt/nuligahelper/current` for working directory, interpreter, code, and runtime configuration. A root-owned deployment command switches the symlink through a temporary link and atomic rename, after recording the previous target. Release directories are never service-writable.

This layout fits the existing permission preflight: it requires the configured application root itself to be a directory, but allows a symlink inside that root when the resolved target remains inside the same root. It also preserves the approved `app_dir` value `/opt/nuligahelper`. All web, daily, cleanup, monitoring, preflight, preview, and other code-running units must be reviewed together so they do not mix releases. Caddy continues to proxy to the one loopback Gunicorn worker.

The deployment source cache and optional read-only GitHub credential live outside the service-readable application tree. Use a source archive or equivalent clean export of the pinned commit, excluding `.git`, local configuration, and test artifacts from the runtime release. Limit retained releases after acceptance, while preserving the immediately prior release for rollback.

### 3. Version generic release assets and pin production dependencies

Create a tracked, non-secret production deployment asset directory separate from the currently ignored private `deploy/` tree. It contains the deploy command source, generic systemd/Gunicorn configuration, and operator instructions. Keep hostnames, contacts, legal content, approval records, environment files, keys, and production state outside Git. Move or reproduce the current production requirements as a tracked lock file with exact direct and transitive versions for the supported server Python platform. Build a fresh virtual environment inside each candidate release and run the full offline suite with synthetic data before activation.

This split avoids accidentally publishing the private `deploy/operations.json` and other host records. Reusing a mutable global virtual environment was rejected because a dependency update could break both the candidate and the previous release at once.

### 4. Separate preparation from activation

The operator runs a `prepare` phase that fetches and pins the source, checks disk capacity, exports it to a new release directory, installs locked dependencies, runs `test/run_tests.sh`, and validates relevant unit and Gunicorn syntax. Failure in this phase removes only an incomplete candidate; the current release, database, and services continue untouched. A separate operator `activate <sha>` phase serializes deployments with a host lock and confirms the candidate is the one prepared from that SHA.

The deployment command and installed units are root-managed host assets, not files the application account can rewrite. Updating those host assets is an explicit reviewed installation step and receives syntax validation before reload. A `dry-run` or plan output reports the selected SHA, source/current release, schema state, expected writer units, and required maintenance actions without reading secret values aloud.

### 5. Quiesce writers, snapshot, and gate schema changes

Activation first records the enabled and active state of the relevant systemd units. Stop and disable the daily and cleanup timers so a reboot during maintenance does not silently restart them. If either oneshot service is active, wait for its normal completion under a bound or abort; do not kill it during notification or backup work. Remove public ingress for this application's dedicated Caddy service, stop the web service, and verify no configured database user or unexpected open handle remains. The operator keeps a separate SSH session for recovery.

Create a self-contained SQLite snapshot with the existing online-backup primitive, validate integrity and foreign keys, fsync it, and retain it in a root-only deployment recovery directory outside `/var/lib/nuligahelper`. Never copy only the live `.db` file or manipulate its WAL/SHM sidecars. The deploy command checks the candidate's schema head against a read-only classification of the live database. An exact match proceeds. A recognized older revision stops for the operator to invoke the candidate release's `manage_db.py migrate-schema --confirm-stopped`, which creates its own additional backup and postflight checks. Unknown or newer revisions refuse activation. The initial same-tree promotion should require no migration, but this is verified on the server rather than assumed.

The current daily and cleanup timers both use `OnCalendar=09:00 Europe/Berlin` and `Persistent=true`. Schedule maintenance outside that window. If a scheduled run was missed, leave timers disabled and present the operator with a pending-run warning. The operator either authorizes the immediate catch-up run after reviewing notification/cleanup effects or keeps the timers paused for incident handling. The deploy command never silently resumes them across that boundary.

Stopping only the web service was rejected because the daily job and cleanup service can also write SQLite. A deployment snapshot taken before they finish would not necessarily include their final changes.

### 6. Accept the new release before reopening scheduled work

Atomically switch `current`, start the web service, and check the actual running unit, local `/healthz` through the configured one-hop proxy boundary, database revision/readiness, and the expected loopback listener. Keep public ingress closed until those checks pass. Reopen Caddy and verify the configured HTTPS hostname and `/healthz`, then allow a bounded observation check of application and Caddy state. Only then present the recorded timer state and catch-up decision for operator-controlled resumption. Do not start the daily service as a smoke test: it can scrape, upload a backup, and send member messages.

Write a root-readable deployment record with the selected and previous SHA, source tree identity, schema before/after, snapshot path, unit/timer actions, validation results, and outcome. Record paths and identifiers, never environment values or member data. The existing journal and monitoring remain the primary operational signals.

### 7. Keep rollback conditional on database compatibility

If activation fails before public ingress reopens and no schema migration occurred, switch `current` back and restart the prior web service, then verify readiness before restoring ingress and timers. If migration occurred, the previous release is not assumed compatible: use the existing validated `restore-snapshot --confirm-stopped` procedure only after confirming that no public writes were accepted, or activate a reviewed compatible release. If public traffic has resumed, never automatically replace the database with the predeployment snapshot because that would discard accepted changes. Apply the existing privacy re-clean and restoration checks before reopening a historical snapshot.

Retain the prior release and deployment snapshot until acceptance criteria are met. Cleanup of older releases and snapshots follows the operator-approved retention and disk policy, not an unconditional recursive deletion.

## Risks / Trade-offs

- [A release starts against the wrong database] -> Validate the absolute configured database path against the approved operations configuration and leave it outside releases.
- [A new release changes the schema and breaks rollback] -> Require an explicit migration gate, preserve a validated predeployment snapshot, and refuse code-only rollback when revisions differ.
- [A persistent timer immediately runs after maintenance] -> Record last/next trigger state, require an operator catch-up decision, and never run the daily job for a smoke test.
- [The ignored private `deploy/` tree cannot be fetched from GitHub] -> Extract generic assets to tracked files while keeping site values outside Git; check the archive contents during preparation.
- [The 1 GB RAM or 30 GB SSD host cannot hold two environments and a snapshot] -> Measure available memory and bytes before preparation/cutover and fail before interrupting services if headroom is inadequate.
- [An interruption leaves services paused] -> Keep deployment state and the prior release on disk, make activation resumable or explicitly recoverable, and let existing monitoring identify stale/disabled units.
- [Caddy serves another application later] -> The initial stop/start procedure assumes this dedicated site; revisit ingress isolation before sharing the proxy service.

## Migration Plan

1. Read the running host's actual SHA, source tree, schema revision, unit/timer state, and current backup status. Rehearse a validated restore using a copy, without changing production data.
2. Add the offline pull-request check, enable protection for `master`, promote the running `devel/webui` application baseline through a PR and the applicable checks, and compare application files and schema revisions. The baseline PR was merged before `offline-tests` became required; subsequent PRs must pass it. Do not equate a new merge-commit SHA with a changed application without checking its contents.
3. Implement and test the tracked generic assets, dependency lock, root-owned deploy command, and all unit path updates on a representative non-production installation. Reconcile the private operations runbook, whose older fresh-database steps do not describe this existing production database.
4. Install the host deploy command and prepare the selected `master` commit without interrupting services. The legacy current release has `deploy/gunicorn.conf.py` but lacks `release-assets/gunicorn.conf.py`. Before installing the new web unit, provision that generic compatibility file in the legacy release under root ownership and service-group read access, or defer the unit replacement until maintenance while preserving a tested rollback configuration. Verify the new unit can start against both the candidate and the retained prior release, then validate syntax and permissions; installing it against the unmodified legacy release would break web startup.
5. During an operator-selected maintenance window, run activation with writer quiescence, validated snapshot, schema gate, release switch, readiness checks, controlled public reopening, and deliberate timer resumption. Record the result and observe the service through the agreed window.
6. On failure, keep ingress and writers paused until code/database compatibility is established. Reactivate the previous release when safe; restore the validated snapshot only through the existing guarded procedure when a deliberate data recovery decision is necessary.
