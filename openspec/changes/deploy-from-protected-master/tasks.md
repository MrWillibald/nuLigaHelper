## 1. Establish the Production Baseline and Promotion Gate

- [x] 1.1 Inventory the live host's exact commit/tree, database revision and path, installed web/daily/cleanup units and timers, and latest validated backup; verify a redacted baseline record matches the operator-confirmed `/var/lib/nuligahelper/nuliga_helper.db` and identifies any difference from the planned `devel/webui` head.
- [x] 1.2 Add a GitHub pull-request check that installs the declared dependencies and runs the full offline suite with synthetic data; verify the check passes on a review branch and requires no production credentials or server access.
- [x] 1.3 Configure protected `master` to require pull requests and the `offline-tests` check, and verify the active repository ruleset records those gates without changing production; independent approval is not required in this solo-maintainer repository.
- [x] 1.4 Promote the running `devel/webui` application baseline to `master` through a pull request, and verify differences in application files and schema head against the live commit before treating the first cutover as migration-free. The baseline PR predated the required offline check.

## 2. Package a Reproducible, Non-Secret Release

- [x] 2.1 Extract generic systemd, Gunicorn, and deployment guidance from the ignored private `deploy/` tree into tracked release assets; verify `git ls-files` includes every runtime asset required by a clean GitHub checkout and a secret/host-data review finds none in the tracked assets.
- [x] 2.2 Add a tracked production dependency lock with exact direct and transitive versions for the supported server Python, and verify a clean virtual environment installs it and passes `test/run_tests.sh` without reading production configuration.
- [x] 2.3 Define `/opt/nuligahelper/releases/<sha>` and the internal `current` link with root ownership and service-group read access, and verify the existing permission preflight accepts a synthetic prepared release but rejects an external or service-writable link target.
- [x] 2.4 Update each installed code-running unit to resolve its interpreter, working directory, and code through `current`, while retaining the one shared environment and SQLite path; verify `systemd-analyze verify` and inspected unit properties show no mixed release paths or secret values.

## 3. Prepare and Record Candidate Releases

- [x] 3.1 Implement a root-managed, operator-invoked source selection command that fetches `master`, pins a full SHA, checks requested-commit reachability, and records the prior release; verify tests reject a non-master commit and show that later branch movement cannot change a prepared candidate.
- [x] 3.2 Export the pinned source to a new release, build its isolated environment, run the offline suite and runtime syntax checks, and mark it prepared only on success; verify injected fetch, install, and test failures leave the active release and services unchanged.
- [x] 3.3 Add deployment serialization, disk/headroom preflight, and a non-sensitive deployment record covering source, previous release, schema, snapshot, checks, and outcome; verify concurrent attempts refuse cleanly and synthetic secret/contact canaries do not appear in output or records.

## 4. Guard Activation and Data Preservation

- [x] 4.1 Implement the maintenance entry sequence for daily and cleanup timers, active oneshot jobs, public ingress, and the web service; verify synthetic active-job cases wait or refuse without killing a running notification/backup task and no database writer remains before cutover.
- [x] 4.2 Create and validate a protected, durable predeployment SQLite snapshot after writers quiesce using the existing backup primitive; verify a WAL-mode fixture retains committed data, integrity and foreign-key checks pass, and snapshot failure prevents activation.
- [x] 4.3 Add a candidate-version schema preflight that proceeds only at the required head, pauses for an explicit guarded migration on a recognized older revision, and refuses missing, unknown, corrupt, divergent, or newer databases; verify each case with synthetic databases and no implicit migration.
- [x] 4.4 Atomically switch the internal `current` link only after all cutover gates pass, and verify interrupted or failed-switch fixtures leave either the old or new complete release addressable with the previous target recorded.
- [x] 4.5 After `Type=simple` start, poll the candidate web unit, expected loopback-only listener, and local readiness under a fixed 60-second monotonic deadline; verify candidate database revision as a hard gate before public HTTPS reopening. Test delayed startup success, failed-unit and timeout refusal, ingress/timers remaining closed until readiness, and no daily-job smoke test.
- [x] 4.6 Detect whether the 09:00 persistent daily or cleanup timer would catch up immediately, require a recorded operator decision before resuming it, and verify simulated missed-trigger cases never start a notification or cleanup run silently.

## 5. Rehearse Rollback and Hand Off Operations

- [x] 5.1 Implement code-only rollback for a compatible unchanged database with the same 60-second monotonic readiness poll before public reopening; verify delayed prior-release startup succeeds without replacing live data, timeout or failed-unit leaves ingress/timers closed and records failure, and schema-incompatible or post-traffic recovery still refuses automatic rollback or snapshot restore.
- [x] 5.2 Reconcile `deploy/OPERATIONS.md`, the current runtime guide, and `README.MD` with the existing production database, operator commands, GitHub promotion, timer catch-up, snapshot retention, migration gate, rollback decisions, and the bounded readiness wait and timeout recovery; verify a walkthrough finds no remaining fresh-database or in-place-update instruction in the active deployment path.
- [x] 5.3 Rehearse prepare, activation, failed health, delayed-start compatible rollback, readiness timeout, and migration-required recovery on a representative non-production installation with synthetic data and units; verify the recorded evidence identifies the exact commit, database revision, snapshot, unit states, and recovered data without exposing secrets. See `rehearsal.md`.
- [ ] 5.4 After the readiness fix is merged through protected `master`, run `test/run_tests.sh`, strict OpenSpec validation, and the host preflight against that candidate release, then complete the first production cutover only after operator review; verify the live database remains at its approved path with preserved row counts and sample assignments/audits, the web service is healthy, and timers are resumed deliberately.
