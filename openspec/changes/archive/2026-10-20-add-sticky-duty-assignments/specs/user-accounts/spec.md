# Spec Delta

## MODIFIED Requirements

### Requirement: Persons are deactivated rather than deleted

The system SHALL let an admin deactivate a person so that the club's history stays
intact. A deactivated person SHALL keep every past assignment and audit entry, SHALL NOT
appear in any person selection, SHALL NOT be assignable, and SHALL NOT be able to log in.
An admin SHALL be able to reactivate them. Deactivation SHALL apply to persons with and
without contact data alike. Deactivation SHALL be refused atomically if it would cancel a current/future binding game or day-block commitment, including an unlocked or replacement-wanted commitment. On refusal the account, MV appointments, assignments and replacement state SHALL remain unchanged and no successful cancellation audit SHALL be appended. A person without affected commitments MAY be deactivated under existing authority; their past assignments and audits SHALL remain. Reactivation SHALL NOT restore historically released assignments.

#### Scenario: Admin deactivates a person who left the club

- **WHEN** an admin successfully deactivates a person without blocking current/future commitments
- **THEN** the person can no longer log in and is offered in no person selection
- **AND** their assignments on games that have already taken place remain recorded and
  keep showing their name
- **AND** the statistics keep counting those past assignments

#### Scenario: Future assignments are freed

- **WHEN** an administrator attempts to deactivate a person with a current/future binding game or day-block commitment
- **THEN** deactivation is refused before any account, MV, assignment or replacement-state change
- **AND** the responsible helper remains active and assigned whether before or after cutoff
- **AND** no successful cancellation audit is written

#### Scenario: Deactivated MV is stood down

- **WHEN** a person without blocking commitments is successfully deactivated and is recorded as the MV of a team
- **THEN** that team is left without an MV
- **AND** the team's registrations fall back to admin approval

#### Scenario: Deactivated person attempts to log in

- **WHEN** a deactivated person requests a login
- **THEN** no session is created
- **AND** the response does not reveal whether the account exists

#### Scenario: Reactivation restores the person

- **WHEN** an admin reactivates a deactivated person
- **THEN** the person appears on the roster again, is assignable again and can log in
  again
- **AND** the historically released assignments are not restored

#### Scenario: Deactivation is not deletion

- **WHEN** an admin successfully deactivates a person without blocking current/future commitments
- **THEN** no audit entry and no past assignment is removed

### Requirement: Deletion stays available for records created in error

The system SHALL keep an admin-only delete action for roster entries that should never
have existed, SHALL explain that binding commitments prevent deletion, and SHALL point to
deactivation as the way to retire someone who did serve once no blocking commitments remain. Audit entries SHALL survive the
deletion. Erroneous-record deletion SHALL be refused atomically if any assignment, including a past one, would be removed. The person, assignments, MV appointments and replacement state SHALL remain unchanged on refusal and no successful deletion audit SHALL be appended. Deletion without affected binding commitments SHALL preserve readable historical audits and existing admin-only authority; administrator status SHALL NOT grant a release or replacement override.

#### Scenario: Admin deletes a mistaken entry

- **WHEN** an admin deletes a person
- **THEN** the interface explains that binding commitments block deletion and that deactivation preserves history
- **AND** an affected binding commitment makes deletion fail atomically without removing the person or their responsibility
- **AND** deletion is permitted only when no binding commitment would be cancelled
- **AND** historical audit entries remain readable

#### Scenario: Only admins may delete

- **WHEN** a member or an MV attempts to delete a person
- **THEN** the request is refused
