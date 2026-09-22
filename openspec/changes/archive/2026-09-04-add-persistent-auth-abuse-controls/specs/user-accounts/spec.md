## MODIFIED Requirements

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

## ADDED Requirements

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
