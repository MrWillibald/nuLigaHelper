## 1. Resolve Production Inputs and Change Boundaries

- [ ] 1.1 Reconcile this change's ownership and integration seams with `deploy-production-web-runtime`, `add-persistent-auth-abuse-controls`, and `make-sqlite-production-safe`; verify a written matrix names the single owning artifact for runtime units, abuse-record cleanup, health/database readiness, backup, and restore behavior with no competing definitions.
- [ ] 1.2 Define the final production filesystem paths, public hostname, service-account name, alert destination, stale/disk/inode/certificate thresholds, and observation window from operator input; verify the production configuration/preflight rejects every unresolved placeholder without printing sensitive values.
- [ ] 1.3 Obtain the operator- and reviewer-approved data inventory, retention periods, processor facts, legal bases, legal-page content, and approval evidence without drafting legal assertions; verify every required approval/version field is complete and traceable before the privacy/legal launch gates can pass.

## 2. Establish Least-Privilege Host and Secret Handling

- [ ] 2.1 Add provisioning documentation or automation for a dedicated no-login, unprivileged service account and separate root-owned application/configuration and service-owned state paths; verify process identity, shell, groups, writable paths, and administrator-only provisioning steps on a representative host.
- [ ] 2.2 Define and enforce owners, groups, modes, parent-directory access, and service umask for `config.json`, `.nuligahelper_secret`, the systemd environment file, SQLite database/sidecars, state markers, and local backups; verify expected effective access as root, the service identity, and an unrelated unprivileged account, including newly created SQLite sidecars/replacement files.
- [ ] 2.3 Add a root-managed, non-repository systemd environment-file workflow shared by web and daily units, with paths rather than secret values in unit arguments; verify unit/command-line inspection does not reveal values and missing/unreadable `NULIGAHELPER_SECRET` or configuration fails startup with only the setting name logged.
- [ ] 2.4 Reconcile and harden the production service/timer definitions from the runtime change with the dedicated identity, environment source, restrictive umask, state paths, restart limits, and failure hooks; verify `systemd-analyze verify` (or the target host's equivalent), installed-unit inspection, and synthetic service/timer starts succeed.

## 3. Add Redacted Logging and Actionable Failure Visibility

- [ ] 3.1 Standardize production operation/phase logging for web startup, daily work, notifications, privacy cleanup, backup, and monitoring with stable identifiers, severity, timestamps, and safe error categories; verify tests with synthetic secret/token/code/e-mail/phone canaries find none of those values in captured output.
- [ ] 3.2 Configure and document journald as the primary log sink with operator-approved age/size bounds and restricted reader access; verify effective retention/disk settings and journal permissions on the representative host.
- [ ] 3.3 Separate daily application-work and backup outcomes, return non-zero status for either failure, and atomically record non-sensitive last-success markers only after each phase completes; verify tests cover application failure, backup-after-success failure, complete success, interruption, and stale markers.
- [ ] 3.4 Implement an operator-configured failure-alert path for web-service failure, timer failure/absence, daily failure, backup failure, cleanup failure, and stale success, including component/time/runbook reference but no raw exception or personal data; verify a synthetic alert reaches the approved destination without scraping, changing schedule data, uploading a backup, or sending member notifications.

## 4. Implement Health, Host Checks, and Monitoring

- [ ] 4.1 Add the fixed, unauthenticated, non-mutating health route with the minimum safe local readiness check agreed with `make-sqlite-production-safe`; verify healthy and unavailable-database tests return only generic fixed bodies/statuses and expose no versions, paths, counts, personal data, configuration, or stack traces.
- [ ] 4.2 Implement a local monitoring/preflight command with machine-usable exit status for web unit state, timer presence/failure, daily/backup/cleanup marker age, filesystem bytes and inodes, and synchronized system time; verify synthetic healthy, warning, critical, missing-timer, stale, and unsynchronized cases produce redacted actionable results.
- [ ] 4.3 Add public TLS hostname/trust/current-validity/remaining-lifetime checks against the endpoint supplied by `deploy-production-web-runtime`; verify invalid, mismatched, expired, warning-window, and healthy certificate fixtures or controlled test endpoints map to the documented states without accessing private-key material.
- [ ] 4.4 Schedule monitoring and connect critical results to the failure-alert path; verify a missed/disabled schedule and each critical host condition become visible even when no application process explicitly failed.
- [ ] 4.5 Write the monitoring and incident runbook with exact status/journal/health/time/disk/certificate commands, expected output, alert ownership, escalation, safe restart guidance, backup cautions, duplicate-notification warnings, and recovery checks; verify an operator walkthrough can diagnose each synthetic failure using the runbook without displaying secrets or contact data.

## 5. Implement Privacy Lifecycle Operations

- [ ] 5.1 Add the versioned privacy data-inventory and retention-policy format covering authentication/session data, pending/rejected/approved/inactive registrations and persons, assignments/audits, journals, local/Dropbox backups, and enabled processors; verify schema/preflight validation blocks missing owners, purposes, locations, periods, dispositions, approval, or processor-review data.
- [ ] 5.2 Implement bounded, idempotent aggregate-preview and apply operations for expired authentication artifacts and any personal identifiers in abuse-control records, using operator-approved cutoffs; verify tests cover expiry/grace boundaries, reruns, interruption, approved distinct abuse-data retention, aggregate-only output, and inability to reuse cleaned credentials.
- [ ] 5.3 Implement bounded, idempotent aggregate-preview and apply operations for rejected and abandoned registrations while preserving approved active/inactive roster people and required approved evidence; verify tests cover each registration state, cutoff boundary, rerun, interruption, and absence of names/contact/authentication values in output.
- [ ] 5.4 Schedule privacy cleanup with distinct success state, non-zero failure and redacted alerts; verify the timer, aggregate evidence, stale/failure monitoring, first-run preview procedure, and policy-change preview procedure on synthetic data.
- [ ] 5.5 Ensure routine lifecycle operations cannot modify append-only assignment audits and document their approved purpose/access/retention/escalation implications; verify regression tests compare audit rows before and after every cleanup class and the runbook prohibits ad hoc SQL edits.
- [ ] 5.6 Implement or reconcile observable expiry/pruning for every approved local and Dropbox backup location, consuming the consistent backup mechanism from `make-sqlite-production-safe`; verify synthetic retained/expired/prune-failure cases and confirm logs/alerts reveal no backup contents or credentials.
- [ ] 5.7 Document and test the quarantined restore procedure—quiesce writers, restore safely, verify integrity/permissions, reapply due cleanup/deletion obligations, then health-check before reopening public traffic; verify a historical synthetic backup cannot be returned to service before cleanup and policy checks pass.
- [ ] 5.8 Complete the third-party processor and privacy-request runbooks for configured e-mail, Twilio, Dropbox, hosting/TLS/monitoring, and any disabled integration, including identity verification, live/log/audit/backup implications, processor follow-up, evidence, and escalation; verify operator walkthroughs distinguish deactivation, erroneous-record deletion, registration cleanup, and privacy requests without promising unsupported backup erasure.

## 6. Integrate Reviewed Public Legal Information

- [ ] 6.1 Define a versioned, non-executable, autoescaped deployment-data format and approval manifest for operator-supplied `Impressum` and `Datenschutzerklärung` content; verify missing files, known placeholders, malformed content, missing approval, or version mismatch block production preflight and no example invents legal facts.
- [ ] 6.2 Add unauthenticated legal routes and responsive German page rendering using the established visual language; verify guest, member, verified-but-unapproved, narrow-screen, keyboard-navigation, autoescaping, and no-private-payload tests for both pages.
- [ ] 6.3 Add consistent `Impressum` and `Datenschutzerklärung` links to the shared public layout, including schedule, login, registration, registration-status, and normal HTML error pages; verify route/template tests assert both links from every public-facing page without weakening the default-deny state-changing-route guard.
- [ ] 6.4 Document approved legal-content correction/deployment and review reopening when operator facts, data handling, retention, backups, or processors change; verify a content-only release can identify old/new versions and approval evidence without modifying schedule/account data or treating technical checks as legal approval.

## 7. Validate Launch and Rollback Readiness

- [ ] 7.1 Create a versioned launch checklist recording operator, release, sibling-change versions, permissions/secrets, database backup/restore evidence, log redaction/retention, health/alerts, timers/markers, time/disk/certificate checks, retention/processor review, legal-content approval, observation window, and explicit go/no-go; verify any absent blocking evidence keeps the checklist in no-go state with owner/remediation recorded.
- [ ] 7.2 Create the paired rollback checklist with triggers, public-traffic removal, writer quiescing, diagnostics capture, prior release/config restoration, database compatibility decision, optional verified restore plus lifecycle re-clean, permission/health/monitoring checks, and duplicate-notification caution; verify a tabletop rollback identifies when the current database is retained versus when restore/escalation is required.
- [ ] 7.3 Add offline automated coverage for permission-policy parsing, secret/log redaction, operation outcomes/markers, health responses, monitoring thresholds, privacy previews/cleanup, audit preservation, legal routing/rendering, and launch-gate validation; verify the focused tests pass with synthetic data and never read production `config.json` or contact external services.
- [ ] 7.4 Run `test/run_tests.sh` and the reconciled production preflight against a representative non-production installation; verify the full offline suite passes and archive redacted evidence for every launch requirement and sibling-change integration seam.
- [ ] 7.5 Perform the final operator/legal/privacy review and synthetic alert/restore/tabletop exercises, then record explicit launch or no-go approval; verify public routing is not enabled until all blocking checklist entries are complete.

## NON-BLOCKING FOLLOW-UP PROPOSAL BACKLOG

The following ideas are explicitly out of scope for this change, use plain backlog bullets rather than completion tasks, and do not count against change completion:

- Notification delivery idempotency and a dispatch ledger.
- Production database schema upgrade/migration strategy.
- Reproducible dependency/version locking and a release-update process.
- Optional broader database scalability/PostgreSQL reassessment only if measured load requires it.
- Optional API/OpenAPI/FastAPI reassessment only if an API-first/mobile use case emerges.
