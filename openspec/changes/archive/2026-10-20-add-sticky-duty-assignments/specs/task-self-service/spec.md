# Spec Delta

## MODIFIED Requirements

### Requirement: Claiming and releasing a single task slot

The system SHALL claim an empty offered game position only after explicit binding confirmation. An ordinary withdrawal from an offered occupied position SHALL open a replacement request while retaining the occupant and responsibility, rather than release the slot. An eligible takeover SHALL transfer responsibility in one atomic operation before the cutoff. Each mutation SHALL compare the expected occupant and relevant assignment/request generation with current saved state. A stale request SHALL return a conflict and the current state without overwriting it. Legacy retirement and operational removal SHALL NOT be inferred from the ordinary release route.

#### Scenario: Slot claimed

- **WHEN** an authorized caller explicitly confirms a valid claim and the slot is still free
- **THEN** the person is recorded in that slot

#### Scenario: Slot released

- **WHEN** an authorized caller requests withdrawal from an offered occupied slot before cutoff
- **THEN** the original helper stays assigned and responsible
- **AND** the slot displays `Ersatz gesucht` instead of becoming free

#### Scenario: Concurrent claim of the same slot

- **WHEN** two callers confirm the same free slot
- **THEN** exactly one claim succeeds
- **AND** the other receives a conflict and current occupant without overwriting the winner

#### Scenario: Release of a slot someone else now holds

- **WHEN** a request names a previous occupant or obsolete request generation
- **THEN** the request is refused with a conflict and nothing is changed

#### Scenario: Claim of a slot filled in the meantime

- **WHEN** a caller confirms a claim after another person filled the slot while its dialog was open
- **THEN** the claim is refused with a conflict and the existing assignment is kept

#### Scenario: Atomic takeover fails

- **WHEN** a takeover fails any current-state check
- **THEN** no part of the requested transfer is saved
- **AND** the original assignment and any still-valid replacement request remain intact

### Requirement: Members act only on their own assignments

An active member SHALL be allowed to confirm an empty-slot claim for themselves, open or end a replacement search for their own unlocked assignment, and confirm a takeover for themselves of an unlocked offered duty marked as seeking replacement. They SHALL NOT assign another person, manage another person's replacement search, or change a locked occupied duty. Opening a request SHALL not authorize unrelated changes.

#### Scenario: Member claims a task

- **WHEN** a member confirms a permitted free-slot claim for themselves
- **THEN** the assignment is recorded

#### Scenario: Member releases their own task

- **WHEN** a member asks for a replacement for their unlocked offered task
- **THEN** the request is recorded while they remain responsible

#### Scenario: Member takes over a requested duty

- **WHEN** an eligible member confirms a takeover of an unlocked offered duty with an active request
- **THEN** the member becomes responsible atomically

#### Scenario: Member tries to assign someone else

- **WHEN** a member attempts to claim or take over for another person
- **THEN** the request is refused

#### Scenario: Member tries to release another person's task

- **WHEN** a member attempts to request withdrawal for another helper
- **THEN** the request is refused

### Requirement: MVs staff the games their team is responsible for

An MV SHALL retain authority to confirm claims for active members of the responsible team they manage and to manage replacement requests for that team's assignments on those games. A takeover on behalf of a team member SHALL require an active replacement request and an eligible replacement belonging to that responsible team. Both team conditions SHALL hold; additional memberships SHALL neither remove nor broaden authority. Ordinary member self-service rights SHALL remain available separately. MV rights SHALL NOT bypass binding confirmation or cutoff.

#### Scenario: MV assigns a team member

- **WHEN** an MV confirms a claim for an active member of the responsible team they manage
- **THEN** the valid claim is recorded

#### Scenario: MV and a game owned by another team

- **WHEN** an MV attempts team-scoped staffing on another responsible team
- **THEN** the request is refused

#### Scenario: MV and a person from another team

- **WHEN** an MV attempts team-scoped staffing for a person outside the game's responsible team
- **THEN** the request is refused

#### Scenario: Person also belongs to another team

- **WHEN** the selected person belongs to the responsible team and additional teams
- **THEN** those additional memberships do not prevent an otherwise permitted claim or takeover

#### Scenario: MV attempts a locked change

- **WHEN** an MV requests a replacement or takeover after the occupied duty locks
- **THEN** the change is refused

### Requirement: Admins assign anyone

An admin SHALL retain roster-wide authority to confirm empty-slot claims and manage unlocked replacement requests and takeovers for active persons, subject to the same binding, current-state, eligibility and cutoff rules. Admin status SHALL NOT permit ordinary release of an offered commitment or any occupied-assignment change after lock.

#### Scenario: Admin reassigns a task

- **WHEN** an admin confirms an eligible takeover for an active replacement request before cutoff
- **THEN** responsibility transfers in one transaction without a committed vacancy

#### Scenario: Admin cannot override a lock

- **WHEN** an admin attempts to replace or release a locked occupied duty
- **THEN** the change is refused and the responsible helper remains recorded

### Requirement: Existing assignment rules apply to every tier

The system SHALL enforce that a person holds at most one task per game, whoever makes the change, and SHALL refuse a claim that would give a person a second task in the same game. Membership in the team playing the game, or membership in neither the responsible team nor the support team, SHALL produce advisory warnings only and SHALL NOT block the assignment. A person SHALL receive the playing-team warning when that team occurs anywhere in their membership set.

Every new claim and takeover SHALL also satisfy the applicable individual age minimum or collective Verkauf age rule, based on the scheduled game date and current stored data. These restrictions SHALL apply to every tier and entry point without an administrator override. A replacement request SHALL preserve an ineligible occupant or sole qualifying adult seller. A takeover SHALL validate the resulting group without counting the outgoing helper. Eligibility deficiencies SHALL NOT grant a release or lock override.

#### Scenario: Second task in the same game refused

- **WHEN** a claim would give a person a second task in a game they already have a task in
- **THEN** the request is refused with an explanatory message

#### Scenario: Self-service claim by a player of the game

- **WHEN** a member claims an age-eligible slot in a game and their membership set includes the team playing that game
- **THEN** the claim succeeds
- **AND** the interface marks the assignment as one where the person's team plays itself

#### Scenario: Person belongs to playing and responsible teams

- **WHEN** a person's memberships include both the playing team and the responsible team for a game and the age rules permit the claim
- **THEN** the assignment remains permitted
- **AND** the playing-team warning takes precedence

#### Scenario: Self-service claim from outside the responsible team

- **WHEN** an age-eligible member whose memberships include neither the responsible team nor the support team claims a slot in a game that has a responsible team
- **THEN** the claim succeeds
- **AND** the interface marks the assignment as coming from outside that team

#### Scenario: Unapproved or deactivated person is never assignable

- **WHEN** a claim names a person whose registration is not approved, or who has been deactivated
- **THEN** the request is refused

#### Scenario: Authorized caller selects an underage helper

- **WHEN** a member, MV, administrator, or CLI caller claims an individually age-restricted task for a person below its minimum
- **THEN** the claim is refused and no successful assignment audit is added

#### Scenario: Sole adult seller releases their task

- **WHEN** the only qualifying adult seller requests replacement before cutoff
- **THEN** they remain assigned and continue to supply adult coverage
- **AND** any takeover must preserve adult coverage in the final occupied group

### Requirement: Task candidate lists prioritize suitable teams

The system SHALL present the people available in each task-assignment selection control as four consecutive, mutually exclusive categories in this order: people whose memberships include the game's responsible team, people whose memberships include the Supporter team, people belonging only to other teams or to no team, and people whose memberships include the team currently playing the game. People SHALL be ordered alphabetically by display name within each category, with a deterministic order for duplicate names.

Membership in the playing team SHALL place a person in the final category even when their memberships also include the responsible or Supporter team. Otherwise responsible-team membership SHALL take precedence over Supporter membership. An absent category, including the responsible-team category when no responsible team is selected, SHALL simply contribute no people. This ordering SHALL only arrange people whom the viewer is already authorized to assign; it SHALL NOT expand their assignment rights. A person already assigned to another task of the same game SHALL be omitted. The currently selected person SHALL remain visible in their own task control. Existing advisory hints for playing-team and outside-team candidates SHALL be retained and SHALL use the complete membership set.

For a free age-restricted game slot, selectable candidates SHALL be limited to claims permitted by the applicable individual or post-claim Verkauf group rule. A younger or unknown-date seller SHALL remain a candidate while the group can remain incomplete or already contains a qualifying adult seller. Candidates SHALL be refreshed when sale-group occupancy or eligibility changes. An existing occupant SHALL remain visible even when no longer eligible, with an explanatory deficiency rather than silently disappearing. Takeover candidates SHALL be offered only for an active unlocked replacement request and SHALL pass the resulting-group eligibility check with the outgoing helper excluded. Locked occupied duties SHALL offer no replacement candidates.

#### Scenario: Full roster is grouped and sorted

- **WHEN** an administrator opens a task selection for a game with a responsible team
- **THEN** eligible available people appear in responsible-team, Supporter, other-team, and playing-team order according to their complete membership sets
- **AND** people within each category are ordered alphabetically by display name
- **AND** playing-team and outside-team candidates retain their advisory hints

#### Scenario: Membership categories overlap

- **WHEN** a candidate belongs to several teams that would otherwise place them in several categories
- **THEN** the candidate appears exactly once
- **AND** playing-team membership has first precedence, followed by responsible-team membership, Supporter membership, and the other category

#### Scenario: Already assigned person is excluded

- **WHEN** a person holds one task for a game
- **THEN** that person is absent from every other task selection for that game
- **AND** the person remains visible as the selected value for their current task

#### Scenario: Released person returns in the correct position

- **WHEN** an atomic takeover is saved without reloading the schedule
- **THEN** the outgoing person is restored to other permitted selections for that game wherever current age rules permit their claim, in the correct membership category and alphabetical position
- **AND** the person's complete team label and advisory hint remain unchanged

#### Scenario: Limited assignment scope remains limited

- **WHEN** a member or MV opens a task selection whose candidates are restricted by their access tier
- **THEN** only candidates permitted for that viewer and the slot's age rule are selectable
- **AND** those candidates follow the applicable membership category and alphabetical ordering

#### Scenario: Timing candidate has an unknown birth date

- **WHEN** a free Zeitnehmer or Sekretär slot is opened
- **THEN** a person with an unknown birth date is not offered as an eligible candidate
- **AND** the interface explains unresolved eligibility without exposing dates or exact ages

#### Scenario: Adult seller is removed

- **WHEN** the only qualifying adult seller seeks a replacement alongside a younger seller
- **THEN** takeover candidates are limited to people who preserve adult coverage

#### Scenario: Stored occupant becomes ineligible

- **WHEN** a date or birth-date change makes an existing occupant ineligible
- **THEN** the occupant remains selected in their own control
- **AND** an authorized caller can seek an eligible replacement before cutoff without vacating their duty
- **AND** a locked assignment remains unchanged even when deficient

### Requirement: Self-service is limited to games that are still ahead

Members and MVs SHALL remain unable to claim or change past-game assignments. Empty offered current/future positions SHALL remain claimable after cutoff, with successful claims immediately locked. The former admin correction permission SHALL NOT authorize changes to locked occupied assignments, including past ones. Destructive account, game, date or block maintenance SHALL be refused when it would cancel a binding commitment; it SHALL NOT grant an ordinary historical correction or staffing override.

#### Scenario: Member claims a past game

- **WHEN** a member or MV attempts a claim, request or takeover for a past game
- **THEN** the request is refused

#### Scenario: Admin corrects a past game

- **WHEN** an admin tries to change a locked occupied past-game assignment
- **THEN** the ordinary assignment service refuses the change

#### Scenario: Late empty claim

- **WHEN** an eligible authorized person confirms a claim for a genuinely empty offered current/future duty after cutoff
- **THEN** the claim succeeds if current checks pass
- **AND** the new assignment is immediately locked

### Requirement: Day-block slots use compare-and-swap assignment

Preparation, cake-delivery and cleanup SHALL use the same binding claim, replacement-request and atomic takeover semantics as offered game duties. Every write SHALL compare current occupant and relevant assignment/request generation under serialization. A stale or failed operation SHALL leave current responsibility unchanged. Claims SHALL respect saved block configuration and capacity.

#### Scenario: Concurrent block claim

- **WHEN** two callers confirm the same empty configured block position
- **THEN** one succeeds and the other receives a conflict with the current occupant

#### Scenario: Stale block release

- **WHEN** a replacement request or takeover names an obsolete block occupant or request generation
- **THEN** it is refused with a conflict and the saved assignment remains

#### Scenario: Cake takeover races configuration

- **WHEN** a cake takeover overlaps a settings change
- **THEN** the saved capacity, request, occupant and cutoff are checked together without losing an assignment

### Requirement: Block assignment authority has no team scope

An active member, including an MV, SHALL confirm an empty block claim or takeover only for themselves and manage a replacement request only for a block duty they hold. An admin SHALL retain the corresponding authority for active roster persons. Blocks SHALL have no responsible team; MV status and membership SHALL not expand block authority. Non-admin past-date restrictions SHALL remain. No tier SHALL change a locked occupied block assignment through this service.

#### Scenario: Member claims own preparation slot

- **WHEN** a member confirms a permitted empty preparation claim for themselves
- **THEN** the assignment is recorded

#### Scenario: MV tries to assign another person

- **WHEN** an MV without admin rights attempts to claim or take over for another person in a block
- **THEN** the request is refused regardless of membership

#### Scenario: Admin staffs a cleanup block

- **WHEN** an admin confirms a valid empty cleanup claim for an active person
- **THEN** it is recorded without a responsible team

#### Scenario: Member changes a past block

- **WHEN** a member or MV attempts a block mutation after its home-game date
- **THEN** it is refused

#### Scenario: Locked block duty

- **WHEN** any tier attempts to change a locked occupied preparation, cake or cleanup position
- **THEN** the original responsible helper remains and the change is refused

### Requirement: Kasse replaces the per-game Unterstützung task

The current per-game role named Unterstützung SHALL be named Kasse. It SHALL remain assignable, auditable, included in assigned-person statistics and included in helper notifications when occupied. It SHALL also carry the additional ordering/security duties of the rejected second Ordner while Ordnungsdienst remains a separate required singleton. Kasse SHALL be offered and required only for adult games. Youth and unknown-category games SHALL not offer new Kasse claims. Retained Kasse assignments SHALL stay visible under existing rights, remain in personal statistics and helper reminders, and SHALL not affect required progress, open duties or MV follow-up. They SHALL offer existing authorized release-only retirement before the fixed cutoff, SHALL refuse ordinary release after cutoff and SHALL never offer a new claim, replacement request or takeover. Destructive account and game/date/block maintenance that would cancel a binding commitment SHALL be refused atomically. No optional game duties SHALL be offered.

The migration SHALL preserve every current Unterstützung assignment's game, person and position under Kasse. It SHALL retain historical audit snapshots with the role text recorded at the time. The newly introduced Reinigung positions SHALL not receive those migrated assignments.

#### Scenario: Empty Unterstützung slot

- **WHEN** every required position of an upcoming youth game is occupied and there is no retained Kasse assignment
- **THEN** the game is treated as completely staffed
- **AND** statistics and MV reminders do not report that removed position as missing

#### Scenario: Empty required Kasse position

- **WHEN** an adult game's Kasse position is empty
- **THEN** that position is reported as a missing required duty

#### Scenario: Occupied Unterstützung slot

- **WHEN** a person is assigned to Kasse
- **THEN** the assignment appears on the schedule and in that person's statistics
- **AND** the person receives the applicable per-game reminders

#### Scenario: Existing Unterstützung assignment migrated

- **WHEN** migration encounters a current Unterstützung assignment
- **THEN** it preserves the same game, position and person under Kasse
- **AND** historical audit snapshots retain their recorded role text
- **AND** the new Reinigung positions begin empty

#### Scenario: Role collision is refused

- **WHEN** migration would overwrite a pre-existing Kasse assignment
- **THEN** migration refuses the collision without silently discarding either assignment

#### Scenario: Existing Reinigung assignment migrated

- **WHEN** a recognized older database still contains original Reinigung assignments and is upgraded through the reviewed migration chain
- **THEN** the existing Reinigung-to-Unterstützung revision and the new Unterstützung-to-Kasse revision preserve the same game, position and person under Kasse
- **AND** historical audit snapshots retain their recorded role text
- **AND** the newly introduced Reinigung positions do not inherit those legacy assignments

#### Scenario: Migrated youth Kasse is release-only

- **WHEN** a youth game has a migrated Kasse assignment
- **THEN** its occupant is shown as an existing assignment subject to binding protection
- **AND** authorized release-only retirement is available before the fixed cutoff and ordinary release is refused afterward
- **AND** no new candidate, replacement claim or empty Kasse field is offered
- **AND** any explicitly authorized retirement appends an audit without changing historical snapshots

### Requirement: On-demand candidates preserve assignment selection behavior

Loading candidates after card expansion SHALL preserve the applicable candidate ordering, complete team labels, playing-team and outside-team hints, and the assignment scope of each slot. A person already assigned to another task in the same game or day-task block SHALL be excluded; the current occupant SHALL remain visible in their own slot even if they are no longer an active or age-eligible candidate. Free game slots SHALL offer only claims permitted by current individual or collective age rules. Unlocked replacement requests SHALL offer only valid takeover candidates. Locked occupied slots SHALL offer no editing candidates. Selecting any claim or takeover SHALL require pre-save binding confirmation; cancellation SHALL reconcile saved state without mutation. Successful claims and atomic takeovers SHALL keep sibling controls and affected Verkauf eligibility consistent without loading unrelated cards. Assignment writes SHALL retain their per-slot compare-and-swap checks and server-side age revalidation. This change SHALL introduce no new age limits for day-task blocks.

#### Scenario: Expand a staffed game

- **WHEN** a viewer opens a game with one or more occupied slots
- **THEN** each occupant remains selected in their own slot
- **AND** no occupant is offered in another slot of that game
- **AND** available age-eligible candidates retain the existing category order and hints

#### Scenario: Expand a staffed day-task block

- **WHEN** a viewer opens a preparation or cleanup block with an occupied slot
- **THEN** that occupant remains selected in their own slot and is absent from its other slots
- **AND** a task in another game or block does not exclude that person

#### Scenario: Change an assignment on an open card

- **WHEN** a claim or atomic takeover succeeds
- **THEN** the other controls in that same card update their candidate availability and affected sale-group eligibility
- **AND** the stored occupant remains authoritative if a later write detects a stale expectation

#### Scenario: Loaded age eligibility becomes stale

- **WHEN** a caller submits a claim after another change invalidates loaded eligibility
- **THEN** the server refuses the claim using current stored data
- **AND** the interface retains saved occupants and refreshes the affected candidates

## ADDED Requirements

### Requirement: End a replacement search while retaining the assigned helper
Before the existing cutoff, an active currently assigned helper SHALL be able to end their own open replacement search on an offered game duty or any day-block phase. The assignment identity, occupant, numbered position, responsibility and occupancy SHALL remain unchanged. Existing authorized administrator and responsible-team MV scope SHALL remain available without extending authority to unrelated people. The service SHALL compare the expected occupant and replacement generation under writer serialization, refresh actor rights and cutoff after waits/retries, and record exactly one successful `request_close` transition with the actual actor. Closed or stale requests SHALL NOT create duplicate successful audits. Reopening SHALL create a fresh token. The JSON route and non-JavaScript POST SHALL require CSRF; visible German `Ersatzsuche beenden` controls SHALL use the same checks. No new migration, release, notification broadcast or timer SHALL be introduced.

#### Scenario: The assigned helper ends their own search
- **GIVEN** an active helper is assigned and their replacement search is open before cutoff
- **WHEN** they submit `Ersatzsuche beenden` with the current occupant and request token
- **THEN** the search closes, one closure audit names the actor, and the same helper remains assigned and responsible

#### Scenario: Another helper or an old request cannot close the search
- **GIVEN** a takeover or a close-and-reopen changed the saved occupant or search generation
- **WHEN** an earlier page attempts cancellation, or an unrelated member attempts to close the current request
- **THEN** cancellation is refused and the current saved request, occupant and successful audit history remain unchanged

#### Scenario: Cancellation waits across the fixed cutoff
- **GIVEN** a cancellation is waiting for write access before the deadline
- **WHEN** it obtains access at or after 09:00 Europe/Berlin on the preceding calendar day
- **THEN** the occupied-duty lock is enforced and no user cancellation or responsibility change is saved

### Requirement: Binding confirmation precedes every assignment creation

Every UI flow creating a game or block assignment, including assignment on behalf of another person and takeovers, SHALL first show German guidance that the assignment is binding, the assigned helper remains responsible until an eligible replacement takes over, and occupied duties become fixed at the day-before reminder. The confirmation SHALL provide explicit `Bestätigen` and `Abbrechen` actions. Selection, dialog display, cancellation or dismissal SHALL perform no assignment write, successful-assignment audit or notification. Only explicit confirmation SHALL attempt the normal current-state-validated claim or takeover. The confirmation SHALL identify the selected duty and helper, preserve keyboard/focus access and have a usable non-JavaScript equivalent. It SHALL NOT falsely record another person's personal acknowledgment when an authorized MV/admin acts for them.

#### Scenario: Cancel a free duty

- **WHEN** a user selects an empty game, preparation, cake or cleanup duty and cancels or dismisses its confirmation
- **THEN** no assignment mutation, successful-assignment audit or notification occurs
- **AND** the duty remains empty if no other caller filled it

#### Scenario: Cancel preserves a concurrent claim

- **WHEN** another helper fills a duty while its confirmation is open and the user cancels
- **THEN** cancellation leaves that concurrent assignment untouched

#### Scenario: Confirm preserves compare-and-swap

- **WHEN** another helper fills a duty before explicit confirmation
- **THEN** confirmation returns the normal stale-slot conflict without overwriting them

#### Scenario: Confirm a late claim

- **WHEN** a user confirms a still-empty duty after cutoff
- **THEN** successful assignment is immediately locked and the notice makes that consequence clear

#### Scenario: Keyboard and non-JavaScript confirmation

- **WHEN** a user navigates the confirmation by keyboard or without JavaScript
- **THEN** they can read the commitment and explicitly confirm or cancel before a write
- **AND** dismissal returns focus to an appropriate control without submitting

### Requirement: Replacement requests retain responsibility until takeover

An authorized request before cutoff SHALL mark an offered occupied duty `Ersatz gesucht` while retaining its original helper and physical occupancy. No elapsed replacement-search interval SHALL free, return or roll back an assignment. A successful eligible takeover SHALL replace the helper atomically and close that request. Request generations SHALL prevent stale close/reopen requests from authorizing a later takeover. A request SHALL be a visible schedule status without a new outbound broadcast.

#### Scenario: Search remains occupied

- **WHEN** a helper requests replacement and nobody takes over
- **THEN** the same helper remains responsible and counted as assigned
- **AND** the request stays visible until closed or cutoff

#### Scenario: No automatic return

- **WHEN** an open request has existed for more than 24 hours
- **THEN** elapsed time alone changes neither occupant nor request

#### Scenario: Valid takeover

- **WHEN** an authorized eligible replacement confirms a takeover of the current request before cutoff
- **THEN** responsibility changes atomically and the request closes
- **AND** no committed intermediate vacancy exists

### Requirement: Day-before reminder locks occupied duties

At the authoritative day-before reminder boundary, every occupied game and day-block duty SHALL become locked and any outstanding replacement request SHALL close with the original helper retained. Weekly reminders SHALL NOT lock. After lock the web assignment service SHALL refuse occupied-duty replacement requests, releases, takeovers and changes for every tier without an admin override. Empty offered positions SHALL remain claimable within existing date/eligibility rights and successful late claims SHALL be immediately locked. The authoritative cutoff SHALL be 09:00 Europe/Berlin on the day before the duty, independently of the reminder job or delivery. Assignment writes SHALL recheck the aware local clock inside serialization and after retries, consult saved locks and persist due dated cutoff/lock/request-closure transitions idempotently. A missing marker, stale page, missed job, contactless helper, failed delivery or restart SHALL NOT postpone the deadline. Reminder batches SHALL persist due transitions before dispatch without advancing the deadline on an early call. Locked state SHALL remain visible on authoritative reads even before cron persists its marker.

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

- **WHEN** any tier requests a change to an occupied locked duty
- **THEN** the service refuses it and preserves the occupant
- **AND** later replacement arrangements remain outside this service

## REMOVED Requirements

### Requirement: No release cutoff

**Reason**: Unconditional vacancy-producing withdrawal conflicts with binding assignments and the agreed day-before lock.

**Migration**: Replace ordinary offered-duty release with a replacement request and atomic takeover before cutoff. Preserve weekly editing without vacancy. Retained removed duties keep authorized release-only retirement before cutoff; after cutoff they are locked for ordinary actions. Destructive maintenance that would cancel commitments is refused as described in design.md; the old release-after-reminder scenarios grant no ordinary lock override.
