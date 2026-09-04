# User Accounts Specification

## Purpose

Registration, approval and passwordless login for the people who use the web interface,
so that every write to the schedule can be attributed to a known member of the club
without the system ever storing a password.

## Requirements

### Requirement: MVs can create users for managed teams

The system SHALL allow an MV to create an active roster person for any team for which
that person is currently the MV. The MV SHALL be able to choose among all of their managed
teams, but SHALL NOT create a person for another team. The existing separate e-mail and
phone fields, including the ability to leave either or both empty, SHALL remain unchanged.

#### Scenario: MV creates a user for a managed team

- **WHEN** an MV submits a valid new-person form with one of their managed teams selected
- **THEN** an active person is created in that team
- **AND** the person appears in the roster and becomes assignable

#### Scenario: MV manages multiple teams

- **WHEN** an MV manages more than one team
- **THEN** the new-person form offers every team they manage
- **AND** the MV may choose any one of those teams

#### Scenario: MV creates a user without contact data

- **WHEN** an MV creates a person without an e-mail address or phone number
- **THEN** the person is created as an active, assignable roster person
- **AND** the person has no way to log in or receive notifications, as for an admin-created
  contactless person

#### Scenario: MV attempts to create a user for another team

- **WHEN** an MV submits or forges a team choice outside their managed teams
- **THEN** the request is refused
- **AND** no person is created

### Requirement: Person identity is an internal identifier

The system SHALL identify every person by an internal identifier that is stable for the
lifetime of the record and is never displayed in the user interface. Names SHALL be
mutable display data and SHALL NOT be required to be unique.

#### Scenario: Person renamed without losing history

- **WHEN** a person's name is changed
- **THEN** all existing task assignments, team membership, account state and audit
  records for that person remain attached to them
- **AND** the schedule shows the new name

#### Scenario: Two persons share a name

- **WHEN** a second person is created or approved with a name that already exists on the
  roster
- **THEN** the system accepts it
- **AND** every place that lists or offers a person for selection shows the person's team
  alongside the name so the two can be told apart

#### Scenario: Name is not an identity key

- **WHEN** a request refers to a person
- **THEN** the reference is the internal identifier
- **AND** a request that identifies a person only by name is rejected

### Requirement: Contact addresses are validated and canonicalized

The system SHALL validate every e-mail address and phone number on the server before it
is used for registration, authentication, or a person contact-data write. The system
SHALL store and compare accepted contact addresses in a canonical form. Browser-side
validation MAY provide earlier feedback but SHALL NOT replace server validation.

#### Scenario: Valid e-mail address submitted

- **WHEN** a visitor or member submits a syntactically valid e-mail address
- **THEN** the system accepts it
- **AND** stores or looks it up using the same canonical representation

#### Scenario: Invalid e-mail address submitted

- **WHEN** an e-mail address does not pass server-side syntax validation
- **THEN** the system shows an explanatory validation error
- **AND** sends no authentication message
- **AND** creates or changes no person record

#### Scenario: Valid international phone number submitted

- **WHEN** a visitor selects a country calling code and enters a valid national phone number
- **THEN** the system accepts the number
- **AND** combines and stores it in E.164 form

#### Scenario: Invalid phone number submitted

- **WHEN** the submitted country calling code and national number do not form a possible
  and valid telephone number
- **THEN** the system shows an explanatory validation error
- **AND** sends no authentication message
- **AND** creates or changes no person record

#### Scenario: Existing contact data is edited to an invalid value

- **WHEN** an admin or member submits invalid e-mail or phone contact data while editing
  a person
- **THEN** the write is rejected
- **AND** the previously stored contact data remains unchanged

#### Scenario: Canonically equivalent contact is reused

- **WHEN** a submitted address is equivalent to an existing contact after canonicalization
- **THEN** uniqueness and account lookup treat both representations as the same contact

### Requirement: Authentication forms present one guided two-step flow

The system SHALL present registration and login as matching, responsive two-step forms
that keep requesting and entering a code on the same page. Every input SHALL have a
visible German label and a short description of the expected value or its purpose.
The registration form SHALL show the e-mail and SMS contact fields before the route
selection. The route choices SHALL be available only for contact values that are
present and valid.

#### Scenario: Visitor opens registration

- **WHEN** an unauthenticated visitor opens the registration page
- **THEN** the page shows fields for name, desired team, e-mail address and SMS number
- **AND** the e-mail and SMS fields appear before the contact-route selection
- **AND** the route selection explicitly offers E-Mail and SMS
- **AND** SMS shows a country-calling-code selector and national-number field
- **AND** the page shows the consent control, code-request action, registration-code field
  and final registration action in their intended order

#### Scenario: Route choices reflect contact validity

- **WHEN** a visitor enters no value or an invalid value for a contact field
- **THEN** the corresponding route cannot be selected
- **AND** the visitor must correct or clear the invalid field before submitting registration

#### Scenario: Visitor opens login

- **WHEN** an unauthenticated visitor opens the login page
- **THEN** the page uses the same layout and contact controls as registration
- **AND** shows the code-request action before the login-code field and final login action

#### Scenario: Code is requested

- **WHEN** a visitor submits valid step-one data to request a code
- **THEN** the same page presents the code-confirmation step
- **AND** preserves the chosen contact context
- **AND** allows the visitor to return to and change the step-one data

#### Scenario: Action hierarchy is shown

- **WHEN** either authentication form is displayed
- **THEN** the code-request button uses a less saturated secondary treatment
- **AND** the final registration or login button uses the established saturated primary treatment

#### Scenario: Form is used without client-side scripting

- **WHEN** client-side JavaScript is unavailable
- **THEN** both code request and code confirmation remain usable through server-rendered form submissions
- **AND** server validation still rejects invalid or incomplete contact data

### Requirement: Self-registration with one or more contact routes and a desired team

The system SHALL let an unauthenticated visitor register by supplying a name, a desired
team, and at least one valid contact route from e-mail and phone. The visitor MAY supply
both routes. Every supplied contact SHALL be validated server-side; an invalid supplied
contact SHALL block registration until it is corrected or cleared. The system SHALL obtain
the registrant's consent to publish their name on the public schedule before issuing the
registration code. The selected valid route SHALL receive the verification code, while
all supplied valid contacts SHALL be stored canonically on the pending registration.

#### Scenario: Registration submitted with both contacts

- **WHEN** a visitor submits a valid name, desired team, e-mail address and phone number,
  confirms consent, and selects either E-Mail or SMS
- **THEN** the system records both canonical contacts on a pending registration
- **AND** sends a six-digit verification code only through the selected route
- **AND** the registrant is not yet on the roster

#### Scenario: Registration submitted with one contact

- **WHEN** a visitor submits a valid name, desired team and exactly one valid contact,
  confirms consent, and selects that contact's route
- **THEN** the system records the supplied canonical contact on a pending registration
- **AND** sends a six-digit verification code through that route

#### Scenario: Registration without a contact

- **WHEN** a visitor requests a registration code without supplying either contact
- **THEN** the registration is rejected because a contact is required to prove control of
  the identity

#### Scenario: Registration with an invalid supplied contact

- **WHEN** a visitor submits an invalid e-mail address or phone number in a non-empty
  contact field
- **THEN** the registration is rejected with an explanatory validation error
- **AND** no authentication message is sent
- **AND** no person record is created or changed

#### Scenario: Registration without consent

- **WHEN** a visitor requests a registration code without confirming the consent notice
- **THEN** the registration is rejected with an explanatory message

#### Scenario: Registration for a contact already in use

- **WHEN** a visitor requests registration with either supplied contact already belonging
  to a person on the roster
- **THEN** the system presents exactly the same code-entry state as it does for an unused
  contact
- **AND** no second person is created or existing person modified
- **AND** any account-exists message is sent only through the selected route

### Requirement: Verification proves the channel, approval grants roster membership

The system SHALL treat channel verification and club membership as separate gates. A
verified registrant SHALL be able to log in and see their own registration status, and
SHALL NOT appear on the roster, in any person selection, or in any assignment until the
registration has been approved.

#### Scenario: Channel verified

- **WHEN** the registrant enters the registration code within its validity period
- **THEN** the account becomes able to log in
- **AND** the registration is queued for approval

#### Scenario: Verified but unapproved person is not assignable

- **WHEN** a registration has been verified but not yet approved
- **THEN** the person does not appear on the roster page, in any person selection, or in
  the statistics
- **AND** any attempt to assign that person to a task is rejected

#### Scenario: Verified but unapproved person signs in

- **WHEN** a verified but unapproved person logs in
- **THEN** they see that their registration is awaiting approval
- **AND** they have no more access than a guest to the rest of the interface

### Requirement: Registrations are approved by the MV of the requested team or by an admin

The system SHALL route a registration to the MV of the team named in the registration.
An MV SHALL be able to approve or reject registrations for their own team only. An admin
SHALL be able to approve or reject any registration. Approval SHALL place the person in
the team named in the registration and SHALL NOT allow the approver to choose a different
team.

#### Scenario: MV approves a registration for their own team

- **WHEN** the MV of the requested team approves a pending registration
- **THEN** the person joins that team, appears on the roster and becomes assignable

#### Scenario: MV cannot act on another team's registration

- **WHEN** an MV attempts to approve or reject a registration for a team they are not MV
  of
- **THEN** the request is rejected

#### Scenario: Requested team has no MV

- **WHEN** a registration names a team that has no MV, or names the support team
- **THEN** only an admin can approve or reject it

#### Scenario: Registration rejected

- **WHEN** an approver rejects a registration
- **THEN** the person does not join the roster
- **AND** the account cannot be used to log in

### Requirement: Passwordless login by e-mail or SMS code

The system SHALL authenticate a person by proving control of any selected contact route
already stored on their record, using a single-use six-digit numeric code sent by e-mail or
SMS. A person with both stored routes SHALL be able to request login through either route.
The selected route SHALL determine delivery even when the person has both contact
channels. The system SHALL NOT store passwords.

#### Scenario: Login by e-mail code

- **WHEN** a person requests a login code for an e-mail address belonging to an account
  that may log in
- **THEN** a six-digit single-use code is sent to that address
- **AND** entering it within its validity period signs the person in

#### Scenario: Login by SMS code

- **WHEN** a person requests a login code for a phone number belonging to an account that
  may log in
- **THEN** a six-digit single-use code is sent by SMS
- **AND** entering it within its validity period signs the person in

#### Scenario: Person with both contacts can use either route

- **WHEN** an eligible person has both an e-mail address and phone number stored
- **THEN** the person can authenticate by requesting and entering a code through either
  route independently

#### Scenario: Active person completes login

- **WHEN** an active person enters a valid login code
- **THEN** the system creates their session
- **AND** redirects them to the Heimspielplan

#### Scenario: Login code is single use

- **WHEN** a login code that has already been used is presented again
- **THEN** the sign-in is refused and a new one must be requested

#### Scenario: Login code expires

- **WHEN** a login code is presented after its validity period has passed
- **THEN** the sign-in is refused and a new one must be requested

#### Scenario: Person without any contact channel cannot log in

- **WHEN** a login is requested for a person who has neither e-mail nor phone
- **THEN** no message is sent and no session is created
- **AND** the person remains assignable by others

### Requirement: Automatic notifications prefer e-mail

The system SHALL use a person's stored e-mail address for automatic notifications when
one is available and valid. It SHALL use the stored phone number as fallback only when no
usable e-mail address is available. This preference SHALL apply independently of the
contact route selected for registration or login.

#### Scenario: Person has both contacts for an automatic notification

- **WHEN** an automatic notification is sent to a person with both a valid e-mail address
  and phone number
- **THEN** the notification is sent by e-mail
- **AND** no SMS is sent for that notification

#### Scenario: Person has only phone for an automatic notification

- **WHEN** an automatic notification is sent to a person without a usable e-mail address
  but with a valid phone number
- **THEN** the notification is sent by SMS

### Requirement: Login and registration do not disclose who is registered

The system SHALL return the same visible response shape and generic user-facing message whether a supplied contact is known, unknown, ineligible, or currently throttled. Abuse-control decisions SHALL NOT reveal whether a person, account, or contact exists, and unknown contacts SHALL participate in contact-based throttling without storing the raw contact as abuse-control state.

#### Scenario: Unknown contact submitted

- **WHEN** a login is requested for a contact channel that belongs to nobody
- **THEN** the response has the same status, form state, and generic message as the response for an eligible known contact
- **AND** no message is sent

#### Scenario: Known contact is throttled

- **WHEN** a login or registration code request for a known contact is refused by an abuse control
- **THEN** the response has the same status, form state, and generic message as a request that was accepted for delivery
- **AND** no additional message is sent

#### Scenario: Unknown contact is repeatedly submitted

- **WHEN** the same unknown canonical contact is submitted repeatedly
- **THEN** those attempts contribute to the same contact-scoped limit without creating a person or exposing the contact in abuse-control records
- **AND** the visible response remains the same shape as for a known contact

### Requirement: Delivery of login and approval messages is rate limited

The system SHALL apply configurable abuse limits to login code requests, authentication-code confirmation attempts, and registration code requests. Limits SHALL cover the originating client and the canonical contact or person/account when one can be resolved; registration involving multiple contacts SHALL apply the relevant contact limits to every supplied canonical contact. SMS SHALL have stricter configurable delivery limits than e-mail because it incurs monetary cost. Enforcement state SHALL survive application restarts and SHALL be shared by all application workers using the deployment's database. Each attempt that is admitted against a limit SHALL be reserved atomically so concurrent requests cannot collectively exceed that limit.

#### Scenario: Repeated login requests for the same person

- **WHEN** login messages for the same person/account or canonical contact are requested more often than a configured limit allows
- **THEN** further requests are refused until the applicable window permits them again
- **AND** no additional message is sent

#### Scenario: Many login requests from one client

- **WHEN** a single attributed client requests logins for many different contacts in quick succession
- **THEN** further requests from that client are refused until the applicable window permits them again
- **AND** no authentication message is sent for a refused request

#### Scenario: Registration is attempted repeatedly

- **WHEN** registration code requests exceed a configured client, supplied-contact, or resolved-person limit
- **THEN** further requests are refused until every applicable limit permits the operation
- **AND** no new pending person, replacement code, account-exists message, or approval notification is created or sent by the refused request

#### Scenario: Authentication codes are guessed repeatedly

- **WHEN** login or registration code confirmations exceed a configured client or challenge-subject limit
- **THEN** further confirmation attempts are rejected with the normal invalid-or-expired-code outcome until the applicable window permits them again
- **AND** a correct code submitted while refused is not consumed and does not establish or advance an account

#### Scenario: Process restarts during a cooling-off window

- **WHEN** the application restarts after an identity or client has reached a limit
- **THEN** the remaining limit window continues to be enforced from shared persistent state

#### Scenario: Concurrent attempts reach the last allowance

- **WHEN** concurrent requests contend for the final allowance under the same limit
- **THEN** at most the configured number of attempts is admitted
- **AND** every other contender is refused without sending a message or changing account state

#### Scenario: Abuse-control storage is unavailable

- **WHEN** the system cannot atomically evaluate and record an applicable authentication abuse limit
- **THEN** the protected action fails closed without sending an authentication message, consuming a code, creating a registration, or establishing a session
- **AND** the visitor receives the normal generic response for that action

### Requirement: Authentication client attribution trusts only configured proxies

The system SHALL derive the client address from forwarding metadata only when the direct request peer and proxy chain satisfy explicit trusted-proxy configuration. Requests that do not satisfy that trust boundary SHALL be attributed to their direct peer, and attacker-supplied forwarding headers SHALL NOT select or rotate an abuse-control identity. Production deployment documentation SHALL state the trusted proxy topology and the matching application configuration.

#### Scenario: Request arrives directly

- **WHEN** an authentication request arrives from a peer that is not configured as a trusted proxy
- **THEN** the direct peer address is used for client-scoped controls
- **AND** any forwarded-client headers are ignored

#### Scenario: Request arrives through the configured proxy chain

- **WHEN** an authentication request arrives through the explicitly configured trusted proxy topology with valid forwarding metadata
- **THEN** the verified originating client address is used for client-scoped controls

#### Scenario: Trusted proxy metadata is missing or malformed

- **WHEN** a trusted proxy request has missing, malformed, or structurally unexpected client-address metadata
- **THEN** the request is not attributed to an attacker-selected address
- **AND** authentication abuse controls fail safely using the trusted direct boundary or refuse the protected action

### Requirement: Authentication abuse-control data is bounded and privacy safe

The system SHALL automatically delete expired abuse-control records after a configurable retention period that is no shorter than the longest enforcement or reporting window. Cleanup SHALL be bounded per execution and SHALL NOT be required to complete before each authentication request can proceed. Stored limiter identities and security logs SHALL NOT contain raw names, e-mail addresses, phone numbers, authentication codes, signed challenges, session values, or reusable secrets. Security logs SHALL record sufficient non-personal information to distinguish admitted, throttled, failed-closed, and cleanup outcomes and the applicable action, channel, and limit dimension.

#### Scenario: Expired records accumulate

- **WHEN** abuse-control records are older than the configured retention cutoff
- **THEN** bounded cleanup removes them over one or more runs
- **AND** live records needed by any configured limit remain available

#### Scenario: Cleanup overlaps authentication traffic

- **WHEN** cleanup and authentication requests execute concurrently
- **THEN** limit enforcement remains correct
- **AND** each cleanup run performs only bounded work so routine authentication is not blocked by an unbounded purge

#### Scenario: A request is throttled

- **WHEN** an authentication or registration action is refused by a configured limit
- **THEN** a security log entry identifies the decision, action, channel, and coarse limit dimension
- **AND** the entry contains no raw personal/contact data, code, challenge, session value, or secret

#### Scenario: Logs and limiter records are inspected

- **WHEN** an operator inspects abuse-control records and normal application logs
- **THEN** raw contacts and authentication credentials cannot be recovered from those records

### Requirement: SMS authentication has layered cost safeguards

The system SHALL support configurable SMS authentication-delivery caps for each person/account or canonical contact and for the application globally. The global cap SHALL cover SMS code delivery across login and registration, SHALL use a configurable operational period, and SHALL be enforced in shared persistent state. Reaching any applicable SMS cap SHALL prevent Twilio dispatch while preserving the normal anti-enumeration response. Deployment documentation SHALL also instruct operators to configure independent Twilio-side spending limits or usage alerts and explain that provider-side controls are the final cost backstop.

#### Scenario: Per-account or contact SMS cap is reached

- **WHEN** an SMS authentication request would exceed its configured person/account or canonical-contact cap
- **THEN** no SMS is dispatched
- **AND** the visitor receives the same response shape and generic message as for an accepted request

#### Scenario: Global SMS cap is reached

- **WHEN** an SMS authentication request would exceed the configured application-global cap for its operational period
- **THEN** no SMS is dispatched regardless of client or destination
- **AND** e-mail authentication remains governed by its own configured limits
- **AND** the refusal is logged without personal or contact data

#### Scenario: Concurrent SMS requests reach a cap

- **WHEN** concurrent SMS requests contend for the final allowance of a per-account, contact, or global cap
- **THEN** no more than the configured number of SMS dispatches is admitted

#### Scenario: Operator prepares production SMS delivery

- **WHEN** an operator enables Twilio-backed authentication SMS in production
- **THEN** the deployment documentation provides steps to configure application caps and independent provider-side spending limits or usage alerts
- **AND** it states that application controls do not replace Twilio account protections

### Requirement: Sessions expire after one hour of inactivity

The system SHALL end a session after one hour without a request from it, and SHALL extend
the session on each request. A person SHALL be able to end their session explicitly.

#### Scenario: Session kept alive by use

- **WHEN** a signed-in person keeps using the interface
- **THEN** the session remains valid

#### Scenario: Session times out

- **WHEN** a signed-in person makes no request for more than one hour
- **THEN** the next request is treated as unauthenticated

#### Scenario: Logout

- **WHEN** a signed-in person logs out
- **THEN** the session is invalid immediately

### Requirement: Persons without contact data remain fully usable

The system SHALL allow an admin to create and maintain persons that have neither e-mail
nor phone. Such persons SHALL be assignable and SHALL be shown on the schedule like any
other person, SHALL receive no notifications, and SHALL have no way to log in.

#### Scenario: Admin creates a person without contact data

- **WHEN** an admin creates a person with no e-mail and no phone
- **THEN** the person joins the roster and can be assigned to tasks
- **AND** the schedule shows their name for the tasks they hold

#### Scenario: Contactless person is skipped by notifications

- **WHEN** notifications are sent for a game a contactless person is assigned to
- **THEN** that person is skipped and the skip is recorded in the run log

### Requirement: Persons are deactivated rather than deleted

The system SHALL let an admin deactivate a person so that the club's history stays
intact. A deactivated person SHALL keep every past assignment and audit entry, SHALL NOT
appear in any person selection, SHALL NOT be assignable, and SHALL NOT be able to log in.
An admin SHALL be able to reactivate them. Deactivation SHALL apply to persons with and
without contact data alike.

#### Scenario: Admin deactivates a person who left the club

- **WHEN** an admin deactivates a person
- **THEN** the person can no longer log in and is offered in no person selection
- **AND** their assignments on games that have already taken place remain recorded and
  keep showing their name
- **AND** the statistics keep counting those past assignments

#### Scenario: Future assignments are freed

- **WHEN** a person holding tasks on games that have not yet taken place is deactivated
- **THEN** those slots are released so the schedule does not show them as covered
- **AND** each release is recorded as a change made by the deactivating admin

#### Scenario: Deactivated MV is stood down

- **WHEN** the person being deactivated is recorded as the MV of a team
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
- **AND** the assignments freed at deactivation are not restored

#### Scenario: Deactivation is not deletion

- **WHEN** an admin deactivates a person
- **THEN** no audit entry and no past assignment is removed

### Requirement: Deletion stays available for records created in error

The system SHALL keep an admin-only delete action for roster entries that should never
have existed, SHALL warn that it removes the person's assignments, and SHALL point to
deactivation as the way to retire someone who did serve. Audit entries SHALL survive the
deletion.

#### Scenario: Admin deletes a mistaken entry

- **WHEN** an admin deletes a person
- **THEN** the interface first states that assignments will be removed and that
  deactivation preserves them
- **AND** on confirmation the person and their assignments are removed
- **AND** the audit entries naming that person remain readable

#### Scenario: Only admins may delete

- **WHEN** a member or an MV attempts to delete a person
- **THEN** the request is refused

### Requirement: Members maintain their own profile

The system SHALL let a signed-in member change their own name, e-mail address and phone
number, including clearing the contact fields. Clearing every contact channel SHALL leave
the person assignable but unable to log in until an admin restores a channel.

#### Scenario: Member updates their own data

- **WHEN** a member changes their own name or contact data
- **THEN** the change is saved and used for future notifications

#### Scenario: Member cannot edit another person

- **WHEN** a member attempts to change another person's data
- **THEN** the request is rejected

#### Scenario: Member removes their last contact channel

- **WHEN** a member clears both e-mail and phone
- **THEN** the change is saved with a warning that they will no longer be able to log in
- **AND** their existing assignments are unaffected
