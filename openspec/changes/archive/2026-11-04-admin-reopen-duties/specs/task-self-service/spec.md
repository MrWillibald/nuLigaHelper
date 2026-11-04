# Spec Delta

## ADDED Requirements

### Requirement: Administrators reopen occupied duties

A current active administrator SHALL be able to switch an occupied game or day-block duty back to open after explicit confirmation, before or after its cutoff. The slot SHALL become empty, the helper SHALL no longer be responsible and any open replacement request SHALL close. Duties dated before today SHALL be refused; open or unparseable dates SHALL be permitted. No other tier or entry point SHALL offer this operation.

#### Scenario: Admin reopens a duty before cutoff

- **WHEN** an administrator confirms reopening an occupied game or day-block duty before its cutoff
- **THEN** the slot is empty and the former helper is no longer assigned
- **AND** the position counts as open in staffing and vacancy reporting

#### Scenario: Admin reopens a locked duty

- **WHEN** an administrator confirms reopening an occupied duty after its 09:00 day-before cutoff or with a saved lock
- **THEN** the slot is empty despite the lock
- **AND** every other occupied duty of that date keeps its lock

#### Scenario: Open replacement request is closed

- **WHEN** an administrator reopens a duty with an open replacement request
- **THEN** the request closes in the same transaction as the slot change
- **AND** a takeover naming that request is refused as stale

#### Scenario: Reopened slot after cutoff is claimed

- **WHEN** an eligible person confirms a claim of a slot reopened after the cutoff
- **THEN** the claim succeeds under the existing late-claim rules
- **AND** the new assignment is immediately locked

#### Scenario: Past duty is refused

- **WHEN** an administrator attempts to reopen a duty dated before today
- **THEN** the request is refused and the assignment and its statistics remain unchanged

#### Scenario: Duty of an open date is reopened

- **WHEN** an administrator reopens a duty whose game or block has an open or unparseable date
- **THEN** the slot becomes empty under the same rules as a current duty

#### Scenario: Non-administrator attempts to reopen

- **WHEN** a guest, member, MV or a former administrator submits a reopen request
- **THEN** the request is refused and the assignment is unchanged
- **AND** the schedule offers them no reopen control

#### Scenario: Occupant changed in the meantime

- **WHEN** a reopen request names an occupant who no longer holds the slot
- **THEN** the request is refused with a conflict and the current occupant
- **AND** no helper is removed or notified

#### Scenario: Reopening leaves a non-youth sale group without an adult

- **WHEN** an administrator reopens the seller who supplied adult coverage at a non-youth game
- **THEN** the slot becomes empty and the remaining seller stays assigned
- **AND** a later claim of the last sale slot must satisfy the collective adult rule

### Requirement: Reopened helpers are notified

After a reopening has been saved, the system SHALL notify the removed helper through the ordinary channel preference with German wording naming the duty and date. A failed or skipped notification SHALL NOT undo the reopening; the administrator SHALL be told that the helper was not reached. Refused or conflicting requests SHALL NOT send a notification.

#### Scenario: Helper is notified

- **WHEN** a reopening is saved for a helper with an e-mail address or phone number
- **THEN** one message is sent, preferring e-mail and falling back to SMS

#### Scenario: Helper cannot be reached

- **WHEN** a reopening is saved and the helper has no contact or delivery fails
- **THEN** the slot stays open
- **AND** the administrator sees a warning to inform the helper personally
- **AND** the warning contains no contact data

## MODIFIED Requirements

### Requirement: Admins assign anyone

An admin SHALL retain roster-wide authority to confirm empty-slot claims and manage unlocked replacement requests and takeovers for active persons, subject to the same binding, current-state, eligibility and cutoff rules. Admin status SHALL NOT permit ordinary release of an offered commitment or any occupied-assignment change after lock, other than the explicit administrator reopening of a duty.

#### Scenario: Admin reassigns a task

- **WHEN** an admin confirms an eligible takeover for an active replacement request before cutoff
- **THEN** responsibility transfers in one transaction without a committed vacancy

#### Scenario: Admin cannot override a lock

- **WHEN** an admin attempts an ordinary release, replacement request or takeover of a locked occupied duty
- **THEN** the change is refused and the responsible helper remains recorded

### Requirement: Day-before reminder locks occupied duties

At the authoritative day-before reminder boundary, every occupied game and day-block duty SHALL become locked and any outstanding replacement request SHALL close with the original helper retained. Weekly reminders SHALL NOT lock. After lock the web assignment service SHALL refuse occupied-duty replacement requests, releases, takeovers and changes for every tier without an admin override; only the explicit administrator reopening of a duty is exempt. Empty offered positions SHALL remain claimable within existing date/eligibility rights and successful late claims SHALL be immediately locked. The authoritative cutoff SHALL be 09:00 Europe/Berlin on the day before the duty, independently of the reminder job or delivery. Assignment writes SHALL recheck the aware local clock inside serialization and after retries, consult saved locks and persist due dated cutoff/lock/request-closure transitions idempotently. A missing marker, stale page, missed job, contactless helper, failed delivery or restart SHALL NOT postpone the deadline. Reminder batches SHALL persist due transitions before dispatch without advancing the deadline on an early call. Locked state SHALL remain visible on authoritative reads even before cron persists its marker.

#### Scenario: Weekly reminder

- **WHEN** the weekly reminder is processed
- **THEN** unlocked assignments remain eligible for replacement requests and takeovers

#### Scenario: Unanswered search at cutoff

- **WHEN** cutoff is reached with an active replacement request
- **THEN** the request closes and its original helper stays assigned and locked

#### Scenario: Takeover crosses cutoff while waiting

- **WHEN** a takeover waits for write access until the cutoff is effective
- **THEN** it is refused and the original helper remains responsible

#### Scenario: Empty position after cutoff

- **WHEN** an eligible helper confirms a genuinely empty offered position after cutoff
- **THEN** it is assigned and immediately locked

#### Scenario: Missed job cannot postpone cutoff

- **WHEN** it is 09:00 Europe/Berlin or later on the day before the duty and no reminder batch has run
- **THEN** occupied-duty replacement requests, releases and takeovers are refused
- **AND** authoritative state shows the locked occupant and closes its request without claiming delivery

#### Scenario: Early reminder does not advance cutoff

- **WHEN** a day-before reminder is processed before 09:00 Europe/Berlin
- **THEN** it does not prematurely lock an otherwise unlocked duty
- **AND** the fixed deadline still applies at 09:00 independently of delivery

#### Scenario: No in-service post-lock replacement

- **WHEN** any tier requests a replacement, takeover, ordinary release or occupant change for an occupied locked duty
- **THEN** the service refuses it and preserves the occupant
- **AND** later replacement arrangements remain outside this service
