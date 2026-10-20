# Spec Delta

## MODIFIED Requirements

### Requirement: Every claim enforces current eligibility atomically

The system SHALL enforce age rules for new claims and responsibility takeovers made through every entry point and access tier, including administrators and CLI operations. A takeover SHALL evaluate the complete resulting group after replacing the original occupant; the outgoing person SHALL NOT count toward post-takeover collective adult Verkauf coverage. Authorization and existing slot/person uniqueness rules SHALL still apply. Eligibility SHALL be evaluated against authoritative game, incoming person, and post-change group state within the same serialized write transaction, and SHALL be reevaluated after a contention retry. A refused claim or takeover SHALL alter neither the responsible occupant nor successful-mutation audit history. A refused takeover SHALL preserve the original occupant and replacement-request state. Opening a replacement request SHALL NOT vacate the slot or remove its occupant from current eligibility reporting. Binding protections and the day-before lock SHALL apply even when the existing occupant is ineligible.

#### Scenario: Administrator attempts an underage appointment

- **WHEN** an administrator claims a timing duty for a person below its minimum
- **THEN** the claim is refused without an age-rule override

#### Scenario: Concurrent younger-person sale claims

- **WHEN** two concurrent claims by nonqualifying sellers target different free sale slots
- **THEN** they cannot both complete a sale group lacking an adult
- **AND** any rejected claim leaves the accepted assignment intact

#### Scenario: Candidate list became stale

- **WHEN** the group changes after candidates were loaded and a later claim no longer meets age rules
- **THEN** the server refuses the claim based on current stored state

#### Scenario: Takeover would remove the only qualifying adult seller

- **WHEN** a full Verkauf group contains one qualifying adult and a younger person attempts to take over that adult's assignment
- **THEN** the takeover is refused because the resulting full group would lack adult coverage
- **AND** the original adult remains assigned and responsible with the replacement request unchanged

#### Scenario: Eligible seller takes over

- **WHEN** an authorized takeover before the day-before lock leaves the resulting Verkauf group compliant with current stored age data
- **THEN** the age rule permits the atomic transfer subject to all other assignment rules
- **AND** no intermediate vacancy or release of the original seller is stored

#### Scenario: Takeover eligibility changes during contention

- **WHEN** a takeover waits or retries and current stored data no longer permits the incoming person or resulting sale group
- **THEN** eligibility is reevaluated and the takeover is refused without replacing the occupant

### Requirement: Withdrawals retain existing rights

An occupant's age deficiency or status as the only qualifying adult seller SHALL NOT prevent an otherwise authorized replacement request before the day-before lock. The request SHALL retain that occupant and their responsibility; it SHALL NOT be treated as a release or as proof that an eligible replacement exists. There SHALL be no ordinary release of an offered occupied duty. A takeover SHALL satisfy current individual and collective age eligibility as well as authorization, compare-and-swap, one-duty and binding-assignment rules. Neither an age deficiency nor administrator status SHALL bypass the occupied-duty lock. Retained removed duties SHALL allow existing authorized release-only retirement before the fixed cutoff and refuse ordinary release afterward. Destructive account, game, date or block maintenance SHALL be refused when it would cancel a binding commitment; it SHALL NOT introduce a replacement override.

#### Scenario: Adult seller withdraws

- **WHEN** the only qualifying adult seller opens an authorized replacement request before the day-before lock
- **THEN** the request is shown while that adult stays assigned and responsible
- **AND** opening the request does not itself create a vacancy or remove adult coverage

#### Scenario: Ineligible occupant is removed

- **WHEN** an authorized replacement request concerns a stored assignment whose occupant no longer meets its age rule
- **THEN** the request can be opened before the lock
- **AND** the current occupant and the visible eligibility deficiency remain until an eligible atomic takeover

#### Scenario: Locked ineligible occupant cannot be released

- **WHEN** any tier attempts to release or replace an occupied duty after the day-before lock because its occupant is ineligible
- **THEN** the web service refuses the change
- **AND** the occupant, responsibility and visible deficiency remain recorded
