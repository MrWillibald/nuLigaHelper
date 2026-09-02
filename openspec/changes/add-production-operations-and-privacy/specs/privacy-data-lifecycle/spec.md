## Purpose

Defines an approved, repeatable privacy data lifecycle for short-lived authentication data, rejected registrations, durable audits, backups and external processors.

## ADDED Requirements

### Requirement: Production has an approved data inventory and retention schedule

Before public launch, the operator SHALL approve a versioned data inventory and retention schedule that identifies each relevant data class, purpose, storage location, owner, access roles, retention trigger, retention period and deletion or anonymization action. At minimum it SHALL cover authentication challenges/tokens, sessions or server-side session state if present, pending and rejected registrations, active and inactive person/contact records, assignment data, append-only assignment audits, application journals and database backups. Retention periods and legal bases SHALL be supplied or approved by the operator and appropriate legal/privacy reviewer; the implementation SHALL NOT invent them.

#### Scenario: Retention schedule is reviewed for launch

- **WHEN** the privacy launch gate is evaluated
- **THEN** every required data class has an approved owner, purpose, retention rule and disposition
- **AND** unresolved placeholders or unapproved periods block launch

#### Scenario: A new personal-data store is introduced

- **WHEN** a later release adds a data class, storage location or processor containing personal data
- **THEN** release approval requires the inventory, retention procedure and public privacy information to be reviewed and updated as applicable

### Requirement: Expired authentication artifacts are routinely removed

Authentication challenges, codes, signed-token records, abuse-control records containing personal identifiers, and server-side session records where applicable SHALL become unusable at their security expiry and SHALL be physically deleted or irreversibly minimized after the operator-approved operational grace period. Cleanup SHALL be repeatable, bounded, safe to rerun and scheduled independently enough that a failed cleanup becomes visible. Security records introduced by `add-persistent-auth-abuse-controls` MAY have a distinct approved retention period but SHALL be included in the same inventory and cleanup evidence.

#### Scenario: Expired authentication data is cleaned

- **WHEN** scheduled cleanup runs after an authentication artifact's expiry and approved grace period
- **THEN** the artifact is deleted or irreversibly minimized according to policy
- **AND** it cannot be used to authenticate or reconstruct the original secret/code

#### Scenario: Cleanup is rerun

- **WHEN** an operator safely reruns cleanup for the same cutoff after an interruption
- **THEN** already processed records cause no failure or unintended deletion of records outside the cutoff

#### Scenario: Cleanup fails

- **WHEN** authentication-data cleanup cannot complete
- **THEN** the operation exits unsuccessfully and produces a redacted operator alert
- **AND** the diagnostic reports data class and cutoff, not token values or contact data

### Requirement: Rejected and abandoned registrations follow an approved deletion rule

Rejected registrations and abandoned unverified or unapproved registrations SHALL be deleted or irreversibly minimized after their separately defined operator-approved retention periods. The disposition SHALL remove unnecessary contact and consent-submission data while retaining only information explicitly justified by the approved schedule. Pending registrations SHALL not be mistaken for active roster persons, and cleanup SHALL not delete an approved person's record.

#### Scenario: Rejected registration reaches its cutoff

- **WHEN** a rejected registration is older than its approved retention cutoff
- **THEN** its unnecessary personal and contact data is deleted or irreversibly minimized
- **AND** it no longer appears in registration administration

#### Scenario: Abandoned registration reaches its cutoff

- **WHEN** an unverified or unapproved registration has had no qualifying activity within its approved retention interval
- **THEN** it is disposed of according to the abandoned-registration rule

#### Scenario: Approved person shares historical registration data

- **WHEN** cleanup evaluates data associated with an approved active or inactive roster person
- **THEN** it preserves the roster record and applies only the explicitly approved disposition to obsolete registration-only data

### Requirement: Cleanup evidence is privacy-preserving

Automated and manual lifecycle operations SHALL record execution time, policy/version, cutoff, outcome and aggregate counts by data class. Their normal logs, dry runs and alerts SHALL NOT list authentication values, names, e-mail addresses, phone numbers or complete record payloads. A dry-run mode or equivalent pre-execution report SHALL allow an authorized operator to validate scope using aggregate counts before the first production cleanup or a policy change.

#### Scenario: Operator previews cleanup

- **WHEN** an authorized operator runs the documented preview for a cutoff
- **THEN** it reports aggregate candidate counts by data class without changing records
- **AND** it does not disclose personal or authentication data

#### Scenario: Cleanup completes

- **WHEN** a lifecycle operation succeeds
- **THEN** protected operational evidence records its policy version, cutoff, aggregate disposition counts and success

### Requirement: Append-only audit implications are explicit

The approved schedule SHALL document that assignment audit entries are append-only application records and may retain historical name snapshots after a roster record is deactivated or deleted. Routine authentication and registration cleanup SHALL NOT alter or delete assignment audit entries. The operator and legal/privacy reviewer SHALL explicitly approve the audit purpose, access, retention period and handling of a valid deletion or restriction request; any requirement to redact or remove audit data SHALL be handled by a separately reviewed change that preserves audit integrity and readability rather than an ad hoc database edit.

#### Scenario: Registration cleanup runs

- **WHEN** rejected-registration or authentication cleanup is executed
- **THEN** assignment audit rows and their historical snapshots remain unchanged

#### Scenario: Deletion request intersects audit history

- **WHEN** an operator handles a request concerning a person represented in assignment audit snapshots
- **THEN** the runbook identifies the approved audit exception or escalation path
- **AND** the operator does not directly edit append-only audit rows

### Requirement: Backup retention and restoration preserve deletion intent

The operational policy SHALL define backup access, encryption expectations, retention/expiry, deletion and restoration procedures for every local and Dropbox database backup. Deletion from the live database SHALL not require rewriting immutable historical backups unless the approved policy says otherwise, but expired backups SHALL be removed on schedule and shall not be kept indefinitely. After restoring a backup, the operator SHALL reapply all lifecycle cutoffs and completed deletion obligations before returning the service to public use, and SHALL verify the restored database using the safety procedure coordinated with `make-sqlite-production-safe`.

#### Scenario: Live data is deleted under policy

- **WHEN** a lifecycle rule removes personal data from the live database
- **THEN** new backups no longer contain that data
- **AND** older backups remain protected and expire according to their approved retention rule

#### Scenario: Backup reaches retention expiry

- **WHEN** a backup passes its approved retention cutoff
- **THEN** the backup is deleted from each documented storage location
- **AND** deletion success or failure is visible to the operator without exposing backup contents

#### Scenario: Historical backup is restored

- **WHEN** an operator restores a backup that may predate lifecycle cleanup
- **THEN** the service remains unavailable to the public until due cleanup/deletion obligations are reapplied
- **AND** integrity, permissions and current policy compliance are verified before service resumes

### Requirement: Third-party processing is inventoried and controlled

Before public launch, the operator SHALL maintain a reviewed inventory of each third party that may process production data, including at least the configured e-mail provider, Twilio when SMS is enabled, Dropbox when backup is enabled, and the public hosting/TLS/monitoring providers where applicable. For each enabled processor, the inventory SHALL state purpose, data categories, transfer path, operator-approved contractual or legal review status, relevant retention/deletion controls and incident contact. A processor SHALL not receive data merely for health monitoring when a non-sensitive local signal suffices.

#### Scenario: Processor is enabled for launch

- **WHEN** a production integration can transmit personal data to a third party
- **THEN** that processor has a completed reviewed inventory entry
- **AND** the public privacy content supplied by the operator reflects the processing where required

#### Scenario: Optional processor is disabled

- **WHEN** SMS, remote backup or another optional integration is not enabled
- **THEN** no production data is sent to that processor
- **AND** monitoring does not create a new personal-data transfer as a substitute

### Requirement: Privacy requests and exceptional deletion are runbook-driven

The operator runbook SHALL describe identity verification, authorization, data-location checks, approved deletion/restriction actions, audit and backup implications, processor follow-up, completion evidence and escalation for privacy requests or erroneous records. It SHALL distinguish normal deactivation, deletion of a record created in error, rejected-registration cleanup and a privacy request, and SHALL avoid promising immediate erasure from protected backups when the approved policy instead relies on access restriction and expiry.

#### Scenario: Operator receives a privacy request

- **WHEN** a person requests access, correction, restriction or deletion
- **THEN** the operator follows the reviewed procedure across live data, logs, audits, backups and enabled processors
- **AND** unresolved legal or audit conflicts are escalated rather than improvised

#### Scenario: Erroneous roster record is deleted

- **WHEN** an admin uses the existing deletion path for a record that should never have existed
- **THEN** the runbook explains the surviving audit snapshot and backup lifecycle implications
- **AND** records held by enabled processors are addressed according to the approved policy
