# Spec Delta

## MODIFIED Requirements

### Requirement: Withdrawals retain existing rights

An occupant's age deficiency or status as the only qualifying adult seller SHALL NOT prevent an otherwise authorized replacement request before the day-before lock. The request SHALL retain that occupant and their responsibility; it SHALL NOT be treated as a release or as proof that an eligible replacement exists. There SHALL be no ordinary release of an offered occupied duty. A takeover SHALL satisfy current individual and collective age eligibility as well as authorization, compare-and-swap, one-duty and binding-assignment rules. Neither an age deficiency nor administrator status SHALL bypass the occupied-duty lock. Retained removed duties SHALL allow existing authorized release-only retirement before the fixed cutoff and refuse ordinary release afterward. Destructive account, game, date or block maintenance SHALL be refused when it would cancel a binding commitment, except removal of block duties of a cancelled home-game date by synchronization; it SHALL NOT introduce a replacement override.

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
