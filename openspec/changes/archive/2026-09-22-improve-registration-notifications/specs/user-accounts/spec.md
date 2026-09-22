## ADDED Requirements

### Requirement: Verified registrations notify every active administrator

After a registrant successfully verifies their contact, the system SHALL send an automatic notification to every active person who is currently an administrator. Each notification SHALL use that administrator's stored e-mail address when usable and otherwise fall back to their stored phone number. The notification SHALL greet the administrator, identify the registrant and every selected team, and direct the administrator to review the registration under "Helfer verwalten". The system SHALL NOT send this approval notification before contact verification succeeds.

#### Scenario: Verified registration has several active administrators

- **WHEN** a registrant verifies their contact and several active administrators exist
- **THEN** every active administrator receives an approval notification through their preferred automatic notification channel
- **AND** the message identifies the registrant and all selected teams
- **AND** the message tells the administrator to review the registration under "Helfer verwalten"

#### Scenario: Administrator has both contact routes

- **WHEN** an active administrator has both a usable e-mail address and phone number
- **THEN** the approval notification is sent by e-mail
- **AND** no SMS is sent to that administrator for the same notification

#### Scenario: Inactive administrator or missing contact

- **WHEN** a person has administrator status but is inactive
- **THEN** no approval notification is attempted for that person
- **AND** when an active administrator has no usable contact route, that administrator is skipped and the outcome is logged without preventing notifications to the other administrators

#### Scenario: Registration has not been verified

- **WHEN** a visitor creates a pending registration but has not successfully verified the selected contact
- **THEN** no administrator receives an approval notification for that registration

#### Scenario: Administrator notification delivery fails

- **WHEN** delivery to one administrator fails after contact verification was committed
- **THEN** the registration remains verified and awaiting approval
- **AND** the failure is logged without preventing delivery attempts to the remaining active administrators

### Requirement: Approved registrants receive a welcome notification

After an administrator successfully approves a verified registration, the system SHALL send the newly active user an automatic welcome notification using e-mail when usable and otherwise their stored phone number. The notification SHALL confirm that the registration was approved and invite the user to sign in and take open duties in the Heimspielplan. The greeting SHALL say "herzlich willkommen beim nuLigaHelper des TuS Raubling Handball!"

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

