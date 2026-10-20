# Spec Delta

## MODIFIED Requirements

### Requirement: Candidate reads use current assignment authority

Every request for assignment candidates SHALL use the viewer's current account status, tier, managed teams, game or block date, responsible team, slot occupant, active replacement request, and cutoff/locked state to determine which candidates may be returned. Candidate responses SHALL contain no e-mail addresses or phone numbers. Guests and verified but unapproved registrants SHALL receive no candidate list. Locked occupied duties SHALL return no replacement candidates. Takeover candidates SHALL be scoped to the incoming person the caller may assign; only active requests on unlocked offered positions SHALL permit takeovers. A member taking over for themselves SHALL not need authority to request withdrawal for the outgoing helper. Candidate reads SHALL never grant or imply authority to write; assignment mutations SHALL continue to enforce their own checks.

#### Scenario: Member requests candidates

- **WHEN** an active member requests candidates for a current or future game or day-task block
- **THEN** only the member's own active person record is returned for slots they may edit

#### Scenario: MV requests candidates

- **WHEN** an MV requests candidates for a game whose responsible team they manage
- **THEN** slots they may manage include active members of that responsible team and the MV's own active record
- **AND** a day-task block and games outside their managed teams retain ordinary self-service scope

#### Scenario: Admin requests candidates

- **WHEN** an admin requests candidates for a game or day-task block, including a past one
- **THEN** the active assignable roster is returned only for editable slots without contact data
- **AND** admin status grants no candidates or write permission for locked occupied slots

#### Scenario: Guest or pending registration requests candidates

- **WHEN** a guest or verified but unapproved registrant requests candidate data directly
- **THEN** the request is refused without returning roster entries or person identifiers

#### Scenario: Permission changes after page load

- **WHEN** a viewer's approval, MV record, team membership, or a game's responsible team changes before a candidate request
- **THEN** the response reflects the stored state at request time
