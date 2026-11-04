# Spec Delta

## MODIFIED Requirements

### Requirement: Every assignment change is recorded

The system SHALL record one entry for every change to a game or day-block task
assignment, whether it was made by the person themselves, by an MV where permitted, by an admin or by permitted pre-cutoff obsolete-duty retirement. Replacement-request opening, authorized cancellation, successful-takeover closure and day-before closure, as well as first locking of an occupied assignment and release of a saved lock because its game was postponed to a later, not-yet-due date or its date became open or unparseable, SHALL be recorded as distinguishable lifecycle transitions even when the occupant does not change. A takeover SHALL record a linked before-and-after responsibility transfer atomically with the assignment and request closure; it SHALL NOT be recorded as an independently committed ordinary release followed by a claim. Repeated enforcement of an already saved lock SHALL NOT append duplicate transition entries. A postponement or open-date unlock SHALL be recorded once per released assignment with system origin and the unchanged responsible helper; it SHALL NOT be recorded as a release or claim. An administrator reopening SHALL be recorded once as a distinguishable `reopen` transition naming the administrator and the removed helper, atomically with the slot change and with closure of any open replacement request; it SHALL NOT be recorded as an ordinary release.

#### Scenario: Self-service claim recorded

- **WHEN** a member claims a game or day-block task slot
- **THEN** an entry is recorded naming that member as both the actor and the affected
  person

#### Scenario: Assignment by someone else recorded

- **WHEN** an authorized MV or admin assigns or releases a task for another person
- **THEN** an entry is recorded naming the actor and the affected person separately

#### Scenario: Destructive maintenance is refused

- **WHEN** account deactivation, erroneous-record deletion or game/day/block removal would cancel a binding commitment
- **THEN** the operation is refused atomically and retains person/container/assignment/request state
- **AND** no audit claims that the commitment was removed


#### Scenario: Cascading removal recorded

- **WHEN** an erroneous-record deletion would cascade to an assigned commitment, including a past assignment
- **THEN** deletion is refused atomically rather than cascading through that commitment
- **AND** no successful assignment-removal audit is written and historical snapshots stay unchanged

#### Scenario: Block removal recorded

- **WHEN** synchronization would remove a block containing an assigned commitment
- **THEN** removal is refused atomically with the block, occupant and request retained
- **AND** no successful-removal audit is written; snapshots remain readable after any later legitimate empty-container removal

#### Scenario: Refused change is not recorded as a change

- **WHEN** a game or day-block change is refused for lack of rights, a conflict, or a
  rule violation
- **THEN** no entry claiming the assignment changed is recorded

#### Scenario: Replacement request records continued responsibility

- **WHEN** an authorized caller opens a replacement request
- **THEN** an append-only lifecycle entry identifies the actor, occupied task and unchanged responsible helper
- **AND** the entry does not claim that the assignment was released

#### Scenario: Atomic takeover recorded

- **WHEN** an eligible authorized replacement takes over before the lock
- **THEN** linked audit evidence records the original and new responsible helper and the request closure
- **AND** that evidence commits atomically with the transfer without an intermediate vacancy

#### Scenario: Day-before lock closes a request

- **WHEN** the day-before boundary first locks an occupied duty with an open replacement request
- **THEN** lifecycle audit evidence identifies the lock and request closure while retaining the original responsible helper
- **AND** enforcing the same saved transition again creates no duplicate entry

#### Scenario: Postponement unlock recorded

- **WHEN** synchronization releases the saved lock of an occupied duty because its game moved to a later, not-yet-due date or has an open or unparseable date
- **THEN** one system lifecycle entry identifies the unlock, the game and position and the unchanged responsible helper
- **AND** the entry does not claim that the assignment was released or reassigned
- **AND** the same evidence is recorded when the first synchronization after this release repairs a game moved before it
- **AND** a later re-lock at a valid date's cutoff is recorded as a new lock transition

#### Scenario: Administrator reopening recorded

- **WHEN** an administrator reopens an occupied game or day-block duty
- **THEN** one `reopen` entry names the administrator as actor and the removed helper as affected person
- **AND** an open replacement request adds one closure entry in the same transaction
- **AND** a refused or conflicting reopen request appends no entry

#### Scenario: Confirmation is cancelled

- **WHEN** a user cancels or dismisses a pre-save assignment confirmation
- **THEN** no assignment or replacement-request mutation audit is appended

#### Scenario: Takeover is refused

- **WHEN** authorization, eligibility, stale expected state or the day-before lock refuses a takeover
- **THEN** no audit claims that responsibility transferred
- **AND** any independently committed reminder-batch lock transition remains separately identifiable
