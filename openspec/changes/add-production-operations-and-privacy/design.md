## Context

See `proposal.md` for motivation. Today the project is operated through shell entry points, stores credentials and notification configuration in `config.json`, accepts a persistent application secret from the environment or `.nuligahelper_secret`, writes a local SQLite database, and uploads database backups to Dropbox. The public runtime, persistent authentication-abuse controls and SQLite production-safety work are being planned separately.

This change crosses host provisioning, systemd, Flask, database maintenance, logs, backups and operator/legal documentation. The deployment target is a Raspberry Pi-class Linux host, so the design favors native systemd/journald capabilities and small local checks over a new monitoring stack. The current no-migration policy remains in force; any schema change still requires explicit owner coordination rather than migration machinery added here.

## Goals / Non-Goals

**Goals:**

- Establish one auditable least-privilege filesystem and process model for both the web service and daily job.
- Make failures and staleness detectable without putting personal or secret data into health responses, logs or alerts.
- Turn privacy retention, backup implications, processor review and public legal pages into explicit launch gates with named operator inputs.
- Keep deployment, monitoring, cleanup and rollback procedures reproducible and testable offline where possible.
- Define clean integration seams with the three related production changes.

**Non-Goals:**

- Choosing the public reverse proxy, WSGI server, TLS automation or network topology owned by `deploy-production-web-runtime`.
- Designing rate limits, challenge lockouts or abuse-state semantics owned by `add-persistent-auth-abuse-controls`.
- Changing SQLite transaction, connection, concurrency or online-backup algorithms owned by `make-sqlite-production-safe`.
- Supplying legal advice, legal wording, retention periods, processor contracts or operator identity details.
- Adding a hosted monitoring/SaaS dependency, changing notification-delivery semantics, introducing schema migrations, or replacing SQLite.

## Decisions

### 1. Separate immutable application, sensitive configuration and mutable state

Use a conventional three-part production layout, with final paths documented by deployment artifacts:

- root-owned application/release files, not writable by the service account;
- a root-managed configuration area containing `config.json`, the persistent secret/environment source and operator-supplied legal content/approval metadata;
- a service-owned state area containing the SQLite database, sidecars and non-secret status markers needed for stale-success checks.

The dedicated account has no shell and receives read access only to runtime/configuration files plus write access only to state paths. Baseline modes are directory `0700` where only the service needs access, secret/config files `0640` or stricter with root ownership and the service's private group when runtime read access is required, and database/sidecar files `0600` owned by the service. Deployment verification inspects parent directories, existing files and effective access; it also verifies the service umask so replacement databases and SQLite sidecars remain restricted.

This separates code deployment from data and secrets and prevents a compromised application from rewriting its executable code or unit definitions. Keeping the whole checkout service-owned was rejected because it expands persistence opportunities. Running as the interactive Raspberry Pi user was rejected because it mixes operator and application authority.

### 2. Use one root-managed systemd environment source and explicit application file references

The systemd units reference a root-managed environment file outside the repository. It contains environment-only values such as `NULIGAHELPER_SECRET`, `NULIGAHELPER_DB`, and paths to configuration/legal data; `config.json` remains a separately permissioned file because it already contains structured credentials and templates. Units never embed values directly and command lines receive paths, not secrets. Both web and daily units use the same environment source so secret rotation cannot accidentally split session-signing behavior.

Provisioning validates required names, file readability as the service account and restrictive modes before restart. It never prints values. A missing production secret is fatal; no production unit may generate an ephemeral replacement.

Systemd's environment-file mechanism was selected for compatibility with the existing application contract and Raspberry Pi distributions. Per-unit inline `Environment=` was rejected because unit inspection exposes values and duplicates configuration. systemd credentials can be reconsidered later where the deployed systemd version and application loading contract support them consistently, but are not required for this baseline.

### 3. Treat journald as the sole primary local log sink and redact at source

Services write normal output/error to journald with unit identity. Application logging uses stable operation/event identifiers, severity and phase fields. Redaction happens before formatting: values from secret/config/authentication/contact fields are never passed as log arguments, and exception handling exposes an error category plus protected detailed traceback only when that traceback has been reviewed not to include request/config payloads. Health and alert paths refer to journal units and correlation identifiers rather than copying arbitrary exception text.

The deployment configures/documentedly verifies bounded journal age/size appropriate to host capacity and restricts journal membership/access to authorized operators. A release-time test suite feeds synthetic canary secrets, codes, tokens, e-mail addresses and phone numbers through representative failures and asserts they do not appear in captured logs.

Separate rotating application log files were rejected because they duplicate retention and permission controls. Logging full request bodies/configuration and filtering later was rejected because redaction misses are difficult to prove and already-written secrets remain exposed.

### 4. Model operations as distinct outcomes with durable non-sensitive success state

The daily entry point exposes at least two named phases: application work and backup. Each phase records a timestamped success marker only after completing successfully; markers contain no member data. Any failed phase yields a non-zero result, and the journal retains which prior phases succeeded. The web service relies on systemd active/failed state. systemd timers schedule daily and privacy-cleanup work and make missed/failed execution inspectable.

A small local monitoring command checks unit/timer state, marker age, backup outcome, disk/inodes, time synchronization and certificate validity. It returns machine-usable exit codes and redacted summaries. A systemd timer runs it, while `OnFailure=` and/or a small failure-notification unit calls an operator-configured alert adapter. The adapter destination is a required deployment input and can use an existing local mail path or another operator-approved channel; it must not reuse member-notification code in a way that could message members. A synthetic alert test proves delivery without scraping, database mutation, backup upload or ordinary notifications.

Relying only on an operator periodically reading journals was rejected because silent timer omission and stale success would remain invisible. Treating backup errors as warnings was rejected because recoverability is a launch requirement.

### 5. Provide one generic readiness endpoint and keep detail local

Add a fixed public health route returning only a small constant healthy body/status or generic unhealthy body/status. It performs no scrape, message or backup and no write. Its readiness check is limited to process responsiveness and the minimum safe local database usability check agreed with `make-sqlite-production-safe`; detailed component results remain in protected logs/local monitoring output. It exposes no version, counts, paths or dependency names.

A detailed public diagnostics endpoint was rejected due to reconnaissance and privacy risk. A liveness-only endpoint that always succeeds while the database is unusable was rejected because it would route traffic to a non-ready service. If `deploy-production-web-runtime` uses a reverse proxy, that change configures the proxy/upstream check to consume this route without changing its response contract.

### 6. Express privacy policy as reviewed configuration plus idempotent maintenance operations

Create a versioned operational privacy policy/inventory template whose production values are supplied and approved by the operator/reviewer. It lists data classes, purposes, locations, processors, retention triggers/periods, disposition and responsible owner. Required unresolved fields fail preflight rather than receiving guessed defaults.

Implement bounded, idempotent cleanup operations for expired authentication artifacts and rejected/abandoned registrations. Each supports aggregate preview, applies explicit UTC cutoffs derived from approved policy, commits safely in batches if needed, and emits only aggregate evidence. The scheduled cleanup runs under systemd and participates in failure alerting. Any abuse-control data model supplied by `add-persistent-auth-abuse-controls` must expose its records/retention categories to this cleanup contract.

Audits are deliberately excluded from routine cleanup because the current specification makes them append-only and readable after person deletion. The policy records their separate purpose, retention and escalation path. A request that conflicts with this guarantee triggers separate review/change work, never direct SQL edits.

Hard-coded retention periods were rejected because the operator and legal/privacy reviewer must own them. A single generic “delete old rows” command was rejected because different data classes have different triggers and audit/backup constraints.

### 7. Handle backup deletion through expiry and restore quarantine

The approved policy inventories local and Dropbox backup locations, access, encryption expectations and retention. Live cleanup affects future backups; historical backups remain access-restricted until their approved expiry rather than being rewritten ad hoc. Backup pruning is observable and failure-alerted.

Restore is a quarantined workflow: stop/quiesce writers, restore with the procedure produced by `make-sqlite-production-safe`, verify integrity/permissions, reapply all currently due lifecycle cleanup and known deletion obligations, then start the service and validate health. No historical backup is exposed publicly before that sequence completes.

Immediate mutation of every historical backup after each deletion was rejected because it can undermine integrity and is operationally risky; indefinite backup retention was rejected because it defeats deletion policy.

### 8. Store legal content as reviewed deployment data, not generated prose

The application provides public routes, shared links and an autoescaped rendering structure. Operator-specific Impressum and privacy text live as versioned deployment content separate from route/application logic, with a manifest or equivalent evidence containing content version and approval status. The renderer supports the required headings, paragraphs, lists and links without accepting arbitrary executable template code. Deployment preflight detects absent content, known placeholders and missing approval metadata. Review—not code—confirms substantive accuracy against the privacy inventory and enabled processors.

Hard-coding invented legal text was rejected because the project cannot establish the operator's facts or give legal advice. Raw untrusted HTML/template execution was rejected because it creates an injection path. Requiring a code change for every approved textual correction was rejected because it unnecessarily couples legal updates to application behavior.

### 9. Make related OpenSpec changes explicit prerequisites at integration points

This change may be implemented in parallel where files do not overlap, but final production acceptance observes these contracts:

- `deploy-production-web-runtime` supplies the concrete production WSGI/reverse-proxy/TLS units and public endpoint; this change adds/validates service identity, secret source, hardening, health consumption, certificate monitoring and runbook evidence.
- `add-persistent-auth-abuse-controls` owns abuse rules and persistent record semantics; this change requires those records to be inventoried, redacted and cleaned under an approved retention rule.
- `make-sqlite-production-safe` owns connection/concurrency/backup consistency; this change uses its safe backup/restore mechanism and adds permissions, outcome visibility, retention and restore quarantine.

Where both changes propose the same unit or script, implementation SHALL reconcile them into one maintained artifact rather than installing competing definitions. Launch is blocked until the implemented versions are mutually consistent and each required change's validation is green.

## Risks / Trade-offs

- **[A compromised service account can still read production contacts and credentials needed at runtime]** → Minimize writable paths, keep code/units root-owned, restrict local login/groups, redact output and rotate affected credentials after compromise.
- **[Environment variables can be read by sufficiently privileged host users]** → Restrict host administration, environment-file access and diagnostic dumps; do not put values in command lines or units. Consider systemd credentials later only as a compatible hardening improvement.
- **[Redaction regressions can leak data into journals or alerts]** → Centralize safe logging, test synthetic canary values, forbid payload logging and keep journal access/retention bounded.
- **[A success marker can become stale or be written too early]** → Write atomically only after phase completion and alert on age as well as explicit failures.
- **[An operator alert channel can itself fail]** → Include alert self-tests and a second manual status path in the runbook; periodically verify delivery and stale state.
- **[Automated cleanup can delete too much]** → Require approved cutoffs, aggregate preview on first run/policy changes, bounded idempotent operations, backups and post-run counts; never include audits in generic cleanup.
- **[Backups temporarily retain data deleted from the live database]** → Document this explicitly in reviewed policy/privacy content, restrict access, enforce expiry and re-clean after restoration.
- **[Operator-supplied legal content may be incomplete or outdated]** → Fail the launch gate on placeholders/missing approval and reopen review whenever processing or operator facts change.
- **[Multiple planned production changes can conflict in systemd or backup artifacts]** → Assign ownership by integration seam, reconcile before launch and validate one final installed configuration.
- **[Raspberry Pi storage constraints make logs/backups/SQLite compete for space]** → Bound journals/backups, monitor bytes and inodes with warning/critical thresholds, and test rollback/restore capacity before launch.

## Migration Plan

1. Complete or reconcile the runtime, abuse-control and SQLite-safety prerequisites at the integration points above; record exact versions included in the candidate release.
2. Obtain operator-approved retention values, processor inventory, alert destination, health monitor endpoint/hostname, disk/certificate thresholds, legal content and review evidence. Unresolved placeholders remain blocking.
3. Create the dedicated service account and production directory layout. Install root-owned application/unit files, sensitive configuration and service-owned state with the documented umask and modes.
4. Back up the existing database using the approved consistent method, prove restore in a non-production location, and record integrity evidence before changing service ownership or startup.
5. Install/reload the reconciled systemd service, timers, monitoring/failure units and journal limits. Validate configuration and start with outbound member notifications disabled in a synthetic or staging context.
6. Run permission, secret non-disclosure, log-redaction, health, time-sync, disk/inode, certificate, timer, stale-marker, cleanup-preview and synthetic alert tests.
7. Install approved legal deployment content and verify unauthenticated links/pages across schedule, authentication, status and error views. Record reviewer/content version.
8. Execute the launch checklist and take explicit go/no-go approval. Enable public routing only after all blocking evidence passes, then monitor service health, journal errors, daily/backup markers and certificate/disk state through the defined observation window.

Rollback uses the paired checklist: remove public traffic, stop/quiesce service writers, capture diagnostics without secrets, restore the prior root-owned release/configuration, and assess database compatibility before touching data. Prefer retaining the current database when the prior release is compatible. If restoration is required, use only a verified backup, reapply lifecycle cleanup and deletion obligations, restore permissions, then validate health and monitoring before reopening traffic. Do not manually rerun the daily job until duplicate-notification risk has been assessed.

## Open Questions

- Which operator-controlled alert destination and backup-failure escalation contacts will be approved for production?
- What warning/critical values will the operator approve for disk/inode capacity, marker staleness and certificate lifetime?
- What exact retention periods, legal bases, processor details and legal-page content/version will the operator and appropriate reviewer approve?
- What final filesystem paths and public hostname will `deploy-production-web-runtime` establish on the target host?
