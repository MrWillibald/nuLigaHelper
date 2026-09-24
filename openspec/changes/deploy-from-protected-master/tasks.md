## 1. Establish the Production Baseline and Promotion Gate

- [x] 1.1 Inventory the live host's exact commit/tree, database revision and path, installed web/daily/cleanup units and timers, and latest validated backup; verify a redacted baseline record matches the operator-confirmed `/var/lib/nuligahelper/nuliga_helper.db` and identifies any difference from the planned `devel/webui` head.
- [ ] 1.2 Add a GitHub pull-request check that installs the declared dependencies and runs the full offline suite with synthetic data; verify the check passes on a review branch and requires no production credentials or server access.
- [ ] 1.3 Configure protected `master` to require reviewed pull requests and the project test check, and verify the repository settings or a controlled direct-push refusal records those gates without changing production.
- [x] 1.4 Promote the running `devel/webui` application baseline to `master` through the required review path, and verify differences in application files and schema head against the live commit before treating the first cutover as migration-free.

## 2. Package a Reproducible, Non-Secret Release

- [x] 2.1 Extract generic systemd, Gunicorn, and deployment guidance from the ignored private `deploy/` tree into tracked release assets; verify `git ls-files` includes every runtime asset required by a clean GitHub checkout and a secret/host-data review finds none in the tracked assets.
- [ ] 2.2 Add a tracked production dependency lock with exact direct and transitive versions for the supported server Python, and verify a clean virtual environment installs it and passes `test/run_tests.sh` without reading production configuration.
- [ ] 2.3 Define `/opt/nuligahelper/releases/<sha>` and the internal `current` link with root ownership and service-group read access, and verify the existing permission preflight accepts a synthetic prepared release but rejects an external or service-writable link target.
- [ ] 2.4 Update each installed code-running unit to resolve its interpreter, working directory, and code through `current`, while retaining the one shared environment and SQLite path; verify `systemd-analyze verify` and inspected unit properties show no mixed release paths or secret values.

## 3. Prepare and Record Candidate Releases

- [x] 3.1 Implement a root-managed, operator-invoked source selection command that fetches `master`, pins a full SHA, checks requested-commit reachability, and records the prior release; verify tests reject a non-master commit and show that later branch movement cannot change a prepared candidate.
- [x] 3.2 Export the pinned source to a new release, build its isolated environment, run the offline suite and runtime syntax checks, and mark it prepared only on success; verify injected fetch, install, and test failures leave the active release and services unchanged.
- [ ] 3.3 Add deployment serialization, disk/headroom preflight, and a non-sensitive deployment record covering source, previous release, schema, snapshot, checks, and outcome; verify concurrent attempts refuse cleanly and synthetic secret/contact canaries do not appear in output or records.

## 4. Guard Activation and Data Preservation

- [ ] 4.1 Implement the maintenance entry sequence for daily and cleanup timers, active oneshot jobs, public ingress, and the web service; verify synthetic active-job cases wait or refuse without killing a running notification/backup task and no database writer remains before cutover.
- [ ] 4.2 Create and validate a protected, durable predeployment SQLite snapshot after writers quiesce using the existing backup primitive; verify a WAL-mode fixture retains committed data, integrity and foreign-key checks pass, and snapshot failure prevents activation.
- [ ] 4.3 Add a candidate-version schema preflight that proceeds only at the required head, pauses for an explicit guarded migration on a recognized older revision, and refuses missing, unknown, corrupt, divergent, or newer databases; verify each case with synthetic databases and no implicit migration.
- [ ] 4.4 Atomically switch the internal `current` link only after all cutover gates pass, and verify interrupted or failed-switch fixtures leave either the old or new complete release addressable with the previous target recorded.
- [ ] 4.5 Verify the started web unit, loopback listener, local readiness, candidate database revision, and configured public HTTPS health before accepting activation; verify failed checks keep scheduled work paused and never invoke the daily job as a smoke test.
- [ ] 4.6 Detect whether the 09:00 persistent daily or cleanup timer would catch up immediately, require a recorded operator decision before resuming it, and verify simulated missed-trigger cases never start a notification or cleanup run silently.

## 5. Rehearse Rollback and Hand Off Operations

- [ ] 5.1 Implement code-only rollback for a compatible unchanged database and a guarded stop for schema-incompatible or post-traffic recovery; verify tests reactivate the prior release without replacing live data in the first case and refuse automatic snapshot restore in the latter cases.
- [ ] 5.2 Reconcile `deploy/OPERATIONS.md`, the current runtime guide, and `README.MD` with the existing production database, operator commands, GitHub promotion, timer catch-up, snapshot retention, migration gate, and rollback decisions; verify a walkthrough finds no remaining fresh-database or in-place-update instruction in the active deployment path.
- [ ] 5.3 Rehearse prepare, activation, failed health, compatible rollback, and migration-required recovery on a representative non-production installation with synthetic data and units; verify the recorded evidence identifies the exact commit, database revision, snapshot, unit states, and recovered data without exposing secrets.
- [ ] 5.4 Run `test/run_tests.sh`, strict OpenSpec validation, and the host preflight against the candidate release, then complete the first production cutover only after operator review; verify the live database remains at its approved path with preserved row counts and sample assignments/audits, the web service is healthy, and timers are resumed deliberately.
