# Assignment Audit Specification

## Purpose

Keeps a durable, append-only record of who assigned or released whom, for which task of
which game and when, so the club can answer after the fact why a task slot changed hands.

## Requirements

### Requirement: Every assignment change is recorded

The system SHALL record one entry for every change to a game or day-block task
assignment, whether it was made by the person themselves, by an MV where permitted, by an admin or by permitted pre-cutoff obsolete-duty retirement. Replacement-request opening, authorized cancellation, successful-takeover closure and day-before closure, as well as first locking of an occupied assignment, SHALL be recorded as distinguishable lifecycle transitions even when the occupant does not change. A takeover SHALL record a linked before-and-after responsibility transfer atomically with the assignment and request closure; it SHALL NOT be recorded as an independently committed ordinary release followed by a claim. Repeated enforcement of an already saved lock SHALL NOT append duplicate transition entries.

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

#### Scenario: Confirmation is cancelled

- **WHEN** a user cancels or dismisses a pre-save assignment confirmation
- **THEN** no assignment or replacement-request mutation audit is appended

#### Scenario: Takeover is refused

- **WHEN** authorization, eligibility, stale expected state or the day-before lock refuses a takeover
- **THEN** no audit claims that responsibility transferred
- **AND** any independently committed reminder-batch lock transition remains separately identifiable

### Requirement: An entry describes the change completely

Each entry SHALL record the time of the change, the acting person, the tier the actor
acted as, the kind of change, the affected person, and the target task slot. A game entry
SHALL identify the game, role, and slot. A day-block entry SHALL identify the season,
home-game date, preparation, cake-delivery or cleanup phase, displayed task name, and slot. Replacement-request and lock entries SHALL identify the unchanged responsible helper, the transition and its actor or system origin. Atomic takeover evidence SHALL identify the previous and new responsible helpers, the acting tier and target position without requiring a transient vacancy. Stored snapshots SHALL contain no full birth date or exact personal age; any necessary eligibility reason SHALL use the established derived threshold or reason only. New snapshot fields SHALL NOT rewrite existing historical snapshots.

#### Scenario: Entry content

- **WHEN** a game assignment change is recorded
- **THEN** the entry states when it happened, who made it, in which tier, what kind of
  change it was, whom it affected, and which task slot of which game was involved

#### Scenario: Day-block entry content

- **WHEN** a preparation, cake-delivery or cleanup assignment change is recorded
- **THEN** the entry states when it happened, who made it, in which tier, what kind of
  change it was, whom it affected, and which named slot of which dated block was involved

#### Scenario: Timestamps are precise

- **WHEN** entries are recorded
- **THEN** each carries a full date and time
- **AND** entries can be ordered by that time without relying on the German date format
  used for game and block dates

#### Scenario: Cake assignment history remains readable

- **WHEN** a cake position is assigned and later an attempted occupied date removal is refused
- **THEN** the audit identifies the dated cake-delivery block and original position
- **AND** that description remains readable after the quantity changes or the block is deleted

#### Scenario: Request and lock history is understandable

- **WHEN** an administrator reviews a request opening, cancellation, cutoff closure or first-lock entry
- **THEN** the entry distinguishes that lifecycle transition from an occupant change
- **AND** it identifies the duty and the responsible helper retained at that transition

#### Scenario: Transfer history identifies both responsibilities

- **WHEN** an administrator reviews a successful game or day-block takeover
- **THEN** the linked evidence identifies the outgoing and incoming helper, actor, tier, exact dated position and time
- **AND** the evidence remains readable after renaming, capacity change, deactivation or permitted removal

#### Scenario: Eligibility is recorded without private dates

- **WHEN** takeover audit evidence includes a necessary age-eligibility reason
- **THEN** it contains the applicable derived reason or threshold rather than a full birth date or exact personal age
- **AND** prior audit snapshots remain unchanged

### Requirement: The record survives the people and games it refers to

Each entry SHALL retain a readable description of the acting person, the affected person,
and the game or day block, including the names and labels in force at the time of the
change, so that the entry stays meaningful after those records are renamed or deleted.

#### Scenario: Person deleted afterwards

- **WHEN** a person named in an entry is deleted from the roster
- **THEN** the entry remains and still shows the name that person had at the time

#### Scenario: Person renamed afterwards

- **WHEN** a person named in an entry changes their name
- **THEN** the entry still shows the name in force when the change was made

#### Scenario: Role renamed afterwards

- **WHEN** a task role is renamed after an entry was recorded
- **THEN** the entry retains the role name recorded at the time

#### Scenario: Day block deleted after its games disappear

- **WHEN** a dated block is deleted because its date no longer has any home game
- **THEN** its prior and system-generated removal audit entries remain readable

### Requirement: The record is append-only

The system SHALL NOT offer any way to edit or delete an entry through the web interface
or the management CLI.

#### Scenario: No way to alter history

- **WHEN** any user, including an admin, is signed in
- **THEN** the interface offers no action that changes or removes an existing entry

### Requirement: Admins can review the record

The system SHALL show the recorded entries to admins in the web interface, most recent
first, and SHALL let them narrow the list down to a single game or a single person. The
review interface SHALL use the application's established visual language and SHALL present
entries in a readable form on both desktop and mobile screens. It SHALL remain read-only.

#### Scenario: Admin reviews recent changes

- **WHEN** an admin opens the record
- **THEN** the most recent changes are listed first with time, actor, affected person,
  game and task
- **AND** the page uses the same section, card, filter, button and typography treatment as
  the rest of the web interface

#### Scenario: Admin investigates one game

- **WHEN** an admin filters the record by a game
- **THEN** only the changes concerning that game are listed

#### Scenario: Admin investigates one person

- **WHEN** an admin filters the record by a person
- **THEN** only the changes concerning that person as actor or affected person are listed

#### Scenario: Activity entries are readable on mobile

- **WHEN** an admin opens the record on a narrow screen
- **THEN** each entry is presented as a readable stacked item or card rather than requiring
  horizontal scrolling across the full desktop table
- **AND** all audit fields remain available

#### Scenario: Activity review remains read-only

- **WHEN** an admin reviews the record
- **THEN** no action is offered that edits or deletes an existing entry

#### Scenario: Non-admin has no access

- **WHEN** a guest, a member or an MV requests the record
- **THEN** access is refused
