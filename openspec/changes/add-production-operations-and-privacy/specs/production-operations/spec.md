## Purpose

Defines the least-privilege, observable and recoverable operating baseline required to launch and maintain nuLigaHelper as a public production service.

## ADDED Requirements

### Requirement: Production processes run as a dedicated unprivileged identity

The production web service, daily job and related maintenance commands SHALL run as a dedicated service account that is not used for interactive administration, has no login shell, has no unnecessary supplementary groups or privilege escalation, and can access only the application paths and operating-system resources required for its duties. Production documentation SHALL identify which one-time provisioning actions require an administrator and which recurring actions run as the service account.

#### Scenario: Service processes are inspected

- **WHEN** an operator inspects the running web service or daily job
- **THEN** each application process runs as the documented dedicated service account
- **AND** it does not run as root or as an operator's personal account

#### Scenario: Interactive login is attempted

- **WHEN** a direct interactive login is attempted for the service account
- **THEN** the operating system refuses the login

### Requirement: Sensitive files and state use restrictive permissions

The production deployment SHALL define and verify restrictive ownership and permissions for the application directories, `config.json`, `.nuligahelper_secret`, any systemd environment file, the SQLite database and its sidecar files, and locally retained backups. Secret-bearing files SHALL be readable only by the administrator and, where runtime access is required, the dedicated service identity. The database and its containing directory SHALL be writable only by the service identity and explicitly authorized administrators; unrelated local users SHALL have no access. Permission verification SHALL account for newly created SQLite sidecar and replacement files rather than checking only the initial database file.

#### Scenario: Production permissions pass verification

- **WHEN** the documented permission check is run after provisioning or deployment
- **THEN** every sensitive file and containing directory has the documented owner, group and mode
- **AND** the service identity has exactly the access needed to operate
- **AND** an unrelated unprivileged local account cannot read secrets or database content and cannot modify application state

#### Scenario: Permissions drift

- **WHEN** a sensitive path becomes broader than the documented policy or is owned by an unexpected identity
- **THEN** the production readiness check fails with the affected path and expected remediation
- **AND** it does not print the file's contents

### Requirement: systemd receives secrets without exposing them

Production secrets and sensitive environment values SHALL be supplied through a root-managed, non-repository environment or credential file with restrictive permissions and referenced by systemd rather than embedded in unit files, command-line arguments or source-controlled deployment files. The same persistent `NULIGAHELPER_SECRET` SHALL be available to the web service and daily job. Startup SHALL fail clearly when a required secret or configuration reference is absent or unreadable, without logging the value.

#### Scenario: Service starts with provisioned secrets

- **WHEN** systemd starts the service with the approved secret source present and correctly permissioned
- **THEN** the process receives the required values
- **AND** process command lines, unit definitions and journal entries do not reveal those values

#### Scenario: Required secret is unavailable

- **WHEN** the secret source is missing, unreadable or lacks a required value
- **THEN** startup fails rather than generating or silently substituting a production secret
- **AND** the journal identifies the missing setting by name without exposing any secret value

### Requirement: Production logs are useful, bounded and redacted

Production service and job output SHALL be captured by journald with stable unit identifiers, timestamps, severity and enough operation context to diagnose startup, scraping, notification and backup outcomes. The operator documentation SHALL define journal retention limits, disk-use bounds and who may read production logs. Logs SHALL NOT contain credentials, secret values, authentication codes or tokens, session/CSRF material, full configuration payloads, full database rows, request bodies, or unnecessary e-mail addresses and phone numbers. Personal records SHALL be referred to by non-sensitive internal/event identifiers where diagnosis requires correlation.

#### Scenario: Operator diagnoses a failed operation

- **WHEN** a web startup, daily run or backup operation fails
- **THEN** the corresponding journal identifies the unit, operation phase, time, severity and actionable error category
- **AND** it provides enough context to select a runbook action without disclosing prohibited data

#### Scenario: Sensitive input causes an error

- **WHEN** an error involves a secret, authentication artifact or contact value
- **THEN** the journal records a redacted diagnostic
- **AND** the sensitive value is not present in either structured fields or free-form text

#### Scenario: Journal retention is inspected

- **WHEN** an operator performs the documented retention check
- **THEN** configured age and/or size bounds match the runbook
- **AND** access to application logs is limited to explicitly authorized operators

### Requirement: Service, daily-job and backup failures are actionable

The web service, each scheduled daily run and each backup attempt SHALL expose success or failure through process exit status and systemd state. A failed daily run SHALL distinguish an application/synchronization failure from a backup failure even if both occur in one invocation. Production monitoring SHALL detect service failure, timer failure or absence, stale last-success state and backup failure, and SHALL notify an operator through a configured channel with the affected component, failure time and a runbook reference. Alert tests SHALL not send real member notifications or mutate production schedule data.

#### Scenario: Web service fails

- **WHEN** the production web process exits unexpectedly or repeatedly fails to start
- **THEN** systemd records a failed or restart-exhausted state
- **AND** the configured monitoring path notifies the operator with a service-specific diagnostic reference

#### Scenario: Daily application work fails

- **WHEN** scraping, synchronization or notification processing ends unsuccessfully
- **THEN** the scheduled unit exits unsuccessfully
- **AND** monitoring identifies the daily application run as failed without reporting a successful run merely because later cleanup completed

#### Scenario: Backup fails after daily work

- **WHEN** daily application work succeeds but the database backup fails
- **THEN** the overall scheduled result is unsuccessful
- **AND** monitoring identifies the backup phase as failed while preserving evidence that the preceding application phase succeeded

#### Scenario: Scheduled run becomes stale

- **WHEN** no successful daily run or backup has been observed within the operator-approved interval
- **THEN** monitoring alerts even if no explicit failed process was captured

### Requirement: A minimal health check reveals no sensitive details

The public runtime SHALL provide a non-mutating health check suitable for service supervision. A healthy response SHALL confirm only that the process can serve requests and complete the minimum safe dependency check required for readiness. An unhealthy response SHALL use a non-success status and a generic body. The response SHALL NOT expose version details, host paths, configuration, database contents or counts, personal data, upstream credentials, stack traces, or administrative controls, and invoking it SHALL NOT scrape nuLiga, send messages or start a backup.

#### Scenario: Healthy service is checked

- **WHEN** a monitor requests the health check while the service and required local dependency are ready
- **THEN** it receives a success status with a fixed non-sensitive response
- **AND** no production data or infrastructure detail is disclosed

#### Scenario: Required local dependency is unavailable

- **WHEN** the process is running but cannot safely use the required local production state
- **THEN** the health check returns an unhealthy status with a generic response
- **AND** detailed diagnostics remain available only through the protected journal

### Requirement: Operators have a monitoring and incident runbook

Production documentation SHALL provide commands and expected results for checking the web service, timers, last daily success, last backup success, health response, journal diagnostics, disk state, time synchronization and certificate validity. It SHALL identify alert ownership, escalation, common remediations, safe restart rules, backup/restore cautions and how to verify recovery. Every check SHALL be usable without displaying secrets or member contact data.

#### Scenario: Operator receives an alert

- **WHEN** an operator follows the runbook for a service, daily-job or backup alert
- **THEN** the runbook leads from component status to redacted diagnostics, corrective action and recovery verification
- **AND** it states when to stop and escalate rather than risk data loss or duplicate notifications

### Requirement: Host readiness and expiry checks gate production operation

A repeatable preflight SHALL verify synchronized system time, adequate filesystem capacity for the database, SQLite sidecars, journal and local deployment operations, and—when TLS termination is in scope—the public certificate's hostname, trust, current validity and remaining lifetime. The runbook SHALL define operator-approved warning and critical thresholds. Periodic monitoring SHALL alert before disk exhaustion or certificate expiry and when time synchronization is lost. A certificate check SHALL target the actual public endpoint defined by `deploy-production-web-runtime`, regardless of where TLS terminates.

#### Scenario: Time is not synchronized

- **WHEN** the host clock is not synchronized within the documented production condition
- **THEN** preflight fails or periodic monitoring alerts
- **AND** the runbook explains that authentication expiry, logs and schedules may be unreliable until corrected

#### Scenario: Disk capacity is below threshold

- **WHEN** free capacity or inode availability crosses a documented warning or critical threshold
- **THEN** monitoring identifies the affected filesystem and threshold
- **AND** launch is blocked at the critical threshold

#### Scenario: Certificate approaches expiry

- **WHEN** the public certificate has fewer than the documented warning days remaining, is untrusted, mismatches the hostname or is currently invalid
- **THEN** monitoring alerts with the endpoint and validity condition but no private-key material
- **AND** launch is blocked for an invalid, untrusted or mismatched certificate

### Requirement: Launch and rollback are checklist-driven

Production SHALL have a versioned launch checklist that records responsible operator, release identifier, configuration and permission verification, dependency-change completion, backup and restore evidence, health and monitoring tests, legal/privacy approval, host checks and go/no-go approval. The paired rollback checklist SHALL define triggers, service quiescing, restoration of the prior application/configuration state, database compatibility checks, health verification and post-rollback monitoring. Rollback SHALL avoid restoring an older database blindly and SHALL warn about duplicate outbound notifications before manually rerunning the daily job.

#### Scenario: Launch gate is incomplete

- **WHEN** any blocking checklist item lacks evidence or approval
- **THEN** the production launch remains a no-go
- **AND** the incomplete owner and remediation are recorded

#### Scenario: Rollback is invoked

- **WHEN** a blocking production fault meets a documented rollback trigger
- **THEN** the operator can restore the last known-good application state using the checklist
- **AND** database compatibility is assessed before any database replacement
- **AND** health, service state and monitoring are verified before declaring recovery
