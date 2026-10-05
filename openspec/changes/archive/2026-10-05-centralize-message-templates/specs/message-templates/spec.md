# Spec Delta

## Purpose

Provide consistent notification content with explicit dynamic-value contracts and one developer-maintained source, so message editing remains reviewable across channels and workflows.

## ADDED Requirements

### Requirement: Notifications use one authoritative message source

The system SHALL obtain retained notification subjects, bodies and reusable text fragments from one developer-maintained source. This SHALL cover game and day-block reminders, rescheduling, referee and new-game alerts, authentication and account notifications, Spielfest variants, operator alerts and retained dormant notification features. Configuration SHALL remain responsible for recipient and operational settings rather than authoritative message wording.

#### Scenario: Message wording is updated

- **WHEN** a developer updates a retained message in the authoritative source
- **THEN** every caller of that message uses the updated wording
- **AND** no duplicate caller-local or configuration text overrides it

#### Scenario: Settings-only configuration is used

- **WHEN** the application loads a valid configuration containing club, recipient and operational settings without legacy text definitions
- **THEN** notification messages remain available from the authoritative source
- **AND** delivery still uses the configured recipients and operational settings

### Requirement: Message placeholders have explicit named contracts

Every retained dynamic message SHALL declare its required named values and their meaning. Rendering SHALL place those values into the intended semantic positions and SHALL refuse incomplete or unknown message contexts before dispatch. Supplied values SHALL remain plain text and SHALL NOT be recursively interpreted as template expressions.

#### Scenario: Game reminder is rendered

- **WHEN** a game reminder receives its declared recipient, task and event values
- **THEN** the result identifies the recipient and task with the correct event date and start time
- **AND** any included age class or opponents appear in their intended positions

#### Scenario: Referee alert is rendered

- **WHEN** a referee alert receives a recipient name, age class, event date, event time and notified-person names
- **THEN** each value appears in the part of the message corresponding to its meaning
- **AND** the age class is not substituted for the date and the date is not substituted for the event details

#### Scenario: Required context is missing

- **WHEN** rendering is requested without a required named value or with an unknown message identifier
- **THEN** the system reports a rendering error identifying the message or missing field
- **AND** no partial message is dispatched
- **AND** diagnostics do not expose authentication codes, rendered message contents or contact values

#### Scenario: A display value contains formatting characters

- **WHEN** a person or team name supplied to a message contains literal braces
- **THEN** the braces appear as ordinary text
- **AND** the value does not introduce additional substitutions

#### Scenario: A subject value contains a line break

- **WHEN** a supplied value would introduce a line break into a notification subject
- **THEN** rendering refuses that subject before dispatch
- **AND** the value cannot create an additional mail header

### Requirement: Channel variants preserve required information

The system SHALL retain explicit e-mail and SMS message variants wherever their content differs, and SHALL document intentional shared variants. Each supported variant SHALL preserve the information and action required by its notification purpose. Text centralization SHALL preserve existing recipient selection, delivery preference, authentication-route selection, notification triggers and failure handling.

#### Scenario: A helper receives an SMS reminder

- **WHEN** a helper receives a supported SMS reminder instead of e-mail
- **THEN** the SMS uses its intended channel variant
- **AND** it still identifies the duty and the relevant event or meeting time

#### Scenario: Authentication uses an explicitly selected route

- **WHEN** an eligible person requests an authentication code through a selected contact route
- **THEN** the centralized text is rendered for that existing workflow
- **AND** delivery still uses the explicitly selected route under the existing authentication requirements

#### Scenario: An approved user receives a welcome

- **WHEN** a verified registration is successfully approved and a welcome notification is rendered
- **THEN** its content retains the exact welcome greeting, approval confirmation and invitation required by `user-accounts`
- **AND** informational delivery remains best effort after the account transition commits

#### Scenario: An operator alert is rendered

- **WHEN** an operational alert is sent through the configured operator channel
- **THEN** the content identifies affected components, occurrence time and the existing runbook reference
- **AND** recipient selection remains separate from member-notification delivery

### Requirement: German message wording is clear and consistent

Retained messages SHALL use correct German spelling and grammar, consistent task terminology and readable date/time presentation. Greetings, sign-offs, paragraph structure and recipient actions SHALL be consistent for the intended audience. Wording cleanup SHALL preserve required content and SHALL NOT introduce new task policies or unsupported event information.

#### Scenario: Task label changes

- **WHEN** a notification receives the current display label of a task
- **THEN** its wording uses that label consistently
- **AND** an outdated hardcoded label does not appear in the subject or body

#### Scenario: Spielfest reminder is rendered

- **WHEN** an assigned helper receives a Spielfest reminder
- **THEN** the message identifies the Spielfest and its relevant timing
- **AND** it does not invent an ordinary home-versus-away matchup

#### Scenario: A day-block reminder includes a meeting time

- **WHEN** a day-block reminder is rendered with its calculated or explicitly supplied task time
- **THEN** the wording identifies that value as the task's meeting or performance time
- **AND** it does not describe it as a match kickoff time

### Requirement: Legacy text configuration has an explicit migration path

The system SHALL provide a non-mutating preflight that reports legacy message-setting keys and identifies customized templates for developer reconciliation without exposing message contents or contact values. Legacy text configuration SHALL remain readable during transition, SHALL produce a deprecation warning and SHALL NOT override the authoritative message source. Existing recipient metadata SHALL remain usable through a documented compatibility mapping until migrated, and provider settings SHALL retain their existing contract.

#### Scenario: Existing configuration contains customized texts

- **WHEN** preflight examines a configuration containing customized legacy templates
- **THEN** it reports the affected template keys and their migration status
- **AND** documentation explains how to reconcile the intended wording and placeholders before deployment
- **AND** preflight does not rewrite configuration, send notifications or print private text

#### Scenario: Legacy text section remains after deployment

- **WHEN** a readable existing configuration still contains legacy message text definitions
- **THEN** their presence alone does not reject startup
- **AND** the system warns that those setting keys are deprecated
- **AND** rendered wording comes from the authoritative source

#### Scenario: Referee recipients have not yet migrated

- **WHEN** the new recipient setting is absent and the legacy referee-recipient setting is present
- **THEN** the existing recipients remain available through the compatibility mapping
- **AND** no recipient values are treated as message templates

#### Scenario: Both recipient settings are present

- **WHEN** both the documented new referee-recipient setting and its legacy equivalent are present
- **THEN** the new setting takes precedence
- **AND** conflicting legacy metadata is reported without printing recipient contact values

### Requirement: Template cleanup does not enable dormant notifications

Centralizing or cleaning up retained templates SHALL NOT enable disabled send paths or create notification triggers. Legacy definitions SHALL be removed only when their lack of supported consumers is established and recorded.

#### Scenario: Newspaper templates are extracted

- **WHEN** retained newspaper templates and schedule fragments move to the authoritative source
- **THEN** the newspaper feature retains its existing enabled or disabled state
- **AND** the daily job does not start sending newspaper messages because those templates exist

#### Scenario: An unused early-task definition is removed

- **WHEN** a legacy early-task template has no supported consumer
- **THEN** its removal is recorded in the migration inventory
- **AND** active preparation reminders continue to use their retained templates and existing triggers
