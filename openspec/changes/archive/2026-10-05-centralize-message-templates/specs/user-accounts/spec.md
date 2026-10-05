# Spec Delta

## MODIFIED Requirements

### Requirement: Approved registrants receive a welcome notification

After an administrator successfully approves a verified registration, the system SHALL send the newly active user an automatic welcome notification using e-mail when usable and otherwise their stored phone number. The notification SHALL confirm that the registration was approved and invite the user to sign in and take open duties in the Heimspielplan. The greeting SHALL say "herzlich Willkommen beim nuLigaHelper des TuS Raubling Handball!"

#### Scenario: Registration is approved

- **WHEN** an administrator successfully approves a verified registration
- **THEN** the newly active user receives a welcome notification through their preferred automatic notification channel
- **AND** the message confirms approval and invites the user to sign in and take open duties in the Heimspielplan
- **AND** the message contains the agreed TuS Raubling Handball welcome greeting

#### Scenario: Approval notification delivery fails

- **WHEN** welcome-notification delivery fails after approval was committed
- **THEN** the user remains active and assignable
- **AND** the delivery failure is logged

#### Scenario: Approval request is stale or repeated

- **WHEN** an approval request does not transition a verified registration to active because it is stale, repeated, or otherwise invalid
- **THEN** no welcome notification is sent

#### Scenario: Registration is rejected

- **WHEN** an administrator rejects a verified registration
- **THEN** no approval welcome notification is sent
