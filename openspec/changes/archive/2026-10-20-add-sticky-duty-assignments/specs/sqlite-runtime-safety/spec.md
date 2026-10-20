# Spec Delta

## MODIFIED Requirements

### Requirement: Assignment compare-and-swap semantics survive database contention

The system SHALL preserve per-slot compare-and-swap behavior and assignment/request/lock audit atomicity under the supported concurrent topology. Waiting for a database lock SHALL NOT turn a stale claim, replacement-request change or takeover into an overwrite: after access is obtained, the operation SHALL still apply only if its expected occupant and applicable request state match current committed state. Authorization, active-person status, offeredness/capacity, one-duty rules, current eligibility including the resulting Verkauf group, and the day-before boundary SHALL be revalidated under the same write serialization and after retries. The fixed boundary SHALL be 09:00 Europe/Berlin on the day before the duty, regardless of reminder-job or delivery success. Every assignment write SHALL recheck the aware local clock and saved lock state after obtaining access and after retries, persist due cohort marker/lock/request-closure transitions idempotently under serialization, and refuse post-deadline replacement even if no batch ran. Due reminder processing SHALL commit those transitions before dispatch without advancing the deadline on an early invocation. A successful assignment, request, takeover or lock transition and its audit entries SHALL commit together, while a refused or failed requested mutation SHALL commit neither its state change nor an audit claiming success. A takeover SHALL transfer responsibility without a committed intermediate vacancy. Any separately committed reminder-batch lock transition SHALL remain distinguishable from a refused takeover. A successful claim of an empty duty after the boundary SHALL commit its occupied and locked state together. Network delivery SHALL NOT occur while this serialization is held.

#### Scenario: Concurrent claims retain one winner

- **WHEN** two independent connections claim the same empty assignment slot
- **THEN** exactly one claim and its audit entry are committed
- **AND** the other caller receives a conflict containing the current occupant

#### Scenario: Stale takeover waits behind another writer

- **WHEN** a takeover waits for another successful takeover and its expected occupant or request state is stale when the lock clears
- **THEN** the waiting takeover is refused as a conflict
- **AND** the winner and its audit remain unchanged
- **AND** no successful-transfer audit is written for the refused takeover

#### Scenario: Stale release waits behind another writer

- **WHEN** a legacy ordinary-release request waits behind a successful replacement and then obtains write access with a stale expected occupant
- **THEN** it cannot overwrite or release the saved assignment
- **AND** the replacement and its audit remain unchanged
- **AND** neither stale state nor waiting for a writer creates an ordinary-release permission

#### Scenario: Lock timeout during assignment mutation

- **WHEN** an assignment mutation cannot obtain the database lock within 5 seconds
- **THEN** no assignment or audit change from that mutation is committed
- **AND** the failure is reported as temporary unavailability rather than as a successful claim or a stale-state conflict

#### Scenario: Two replacements contend for one request

- **WHEN** two otherwise valid takeovers contend for the same occupied slot and replacement request
- **THEN** exactly one responsibility transfer and its audit commit
- **AND** the other request receives the current authoritative state without changing the winner

#### Scenario: Cutoff batch wins while a takeover waits

- **WHEN** a takeover waits until 09:00 Europe/Berlin on the day before the duty or until a due cutoff transition has committed before it obtains write access
- **THEN** the saved cutoff is revalidated and the takeover is refused
- **AND** the original helper remains responsible and the committed lock has closed the request

#### Scenario: Late empty-slot claim races another claim

- **WHEN** two callers claim the same truly empty offered position after the day-before boundary
- **THEN** exactly one claim and its immediately locked state commit with their audit
- **AND** the loser receives a conflict without being able to replace the winner

#### Scenario: Cake configuration overlaps a takeover

- **WHEN** a cake configuration update and a takeover contend for the same block
- **THEN** both use current saved capacity and binding protection under shared serialization
- **AND** neither can silently remove an occupied position or transfer responsibility into a removed position

#### Scenario: Takeover commits before the cutoff batch

- **WHEN** a valid takeover commits before the fixed 09:00 Europe/Berlin day-before cutoff
- **THEN** the batch locks the new responsible helper and closes any remaining request against current authoritative state
- **AND** reminders are dispatched only after the cutoff transaction commits

#### Scenario: Dispatch fails after the cutoff commit

- **WHEN** the dated cutoff and all occupied-duty locks have committed and mail or SMS dispatch then fails or is skipped
- **THEN** the saved cutoff and locks remain durable
- **AND** contactless helpers, process restarts and retries cannot reopen the assignments
