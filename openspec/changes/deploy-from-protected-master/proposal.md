## Why

Production already runs the SQLite/web application from `devel/webui`, but there is no repeatable way to promote reviewed GitHub code to that server while preserving its live database. Establish an operator-triggered release process before moving production updates to protected `master`.

## What Changes

- Make protected `master` the source of deployable commits. Promote the currently running `devel/webui` baseline first, and record the exact source commit and tree used for the first cutover.
- Add a manual, server-side deployment workflow that fetches and pins a commit, prepares a separate root-owned release with its dependencies, validates it, and activates it through a stable service path.
- Keep `/var/lib/nuligahelper/nuliga_helper.db`, secrets, and site configuration outside releases. Quiesce the web, daily, and cleanup database writers before the cutover, create and validate a recovery snapshot, and require an explicit guarded schema migration when revisions differ.
- Verify the activated web service, public route, database readiness, and timer state before accepting a release. Retain the prior release and document code rollback versus database restore decisions.
- Put reusable, non-secret deployment assets and production dependency definitions under version control while keeping host-specific values private. Reconcile the process with the existing production operations runbook and systemd units.

## Capabilities

### New Capabilities

- `production-deployment`: Protected-branch release selection, operator-controlled preparation and activation, live-data preservation, validation, and rollback.

### Modified Capabilities

- None. Existing schema-migration and database-backup requirements remain authoritative; this change orchestrates their use during deployment.

## Impact

- Affects GitHub branch governance, deployment tooling and documentation, production dependency packaging, and paths in the installed systemd units. Public Caddy routing and application behavior remain under their existing contracts.
- Uses the existing production SQLite database at `/var/lib/nuligahelper/nuliga_helper.db` and the existing guarded migration and restore commands. No application data or schema is changed by this proposal itself.
- Coordinates with `add-production-operations-and-privacy` for service ownership, permissions, alerts, monitoring, and launch or rollback runbooks; it does not replace those controls.
