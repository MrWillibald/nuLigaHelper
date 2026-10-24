# Game Task Staffing Delta

## ADDED Requirements

### Requirement: Responsible-team members receive a seven-day vacancy invitation

When the daily job processes the date exactly seven calendar days after its effective today, the system SHALL send a personal invitation for each game on that date with at least one physically vacant required position and a saved responsible team. The invitation SHALL use that team's current active memberships and SHALL NOT depend on an appointed MV. The current MV SHALL receive the invitation only when they meet the same recipient rules as every other member. Pending, rejected and inactive people and people outside the responsible team SHALL NOT receive it. Each qualifying person SHALL receive at most one invitation per game per invocation, identified by internal person identity rather than name or membership count. Invitations SHALL prefer e-mail, fall back to SMS when no usable e-mail is present, and skip contactless recipients with the existing warning and zero sent-message count.

The trigger and advertised vacancies SHALL use only the game's category-dependent required positions. A position seeking replacement SHALL remain occupied. Age or classification deficiencies without a physical vacancy SHALL NOT trigger this invitation. Team-independent day-block vacancies SHALL NOT trigger a responsible-team invitation. Existing assigned-helper reminders and day-before MV follow-up SHALL retain their existing conditions. Sending an invitation SHALL NOT assign a duty, change responsibility, alter replacement state or lock an assignment.

#### Scenario: Active responsible-team members are invited a week ahead

- **WHEN** the daily job runs seven days before a game with an assigned responsible team and a vacant required position
- **THEN** each active, otherwise qualifying member of that team receives one personal invitation for that game
- **AND** pending, rejected and inactive people and non-members receive none

#### Scenario: A team needs no appointed MV for recruitment

- **WHEN** a game seven days away has required vacancies and a responsible team without an MV
- **THEN** its qualifying active team members still receive invitations
- **AND** no fallback team or administrator recipient is invented

#### Scenario: A qualifying MV receives one member invitation

- **WHEN** the responsible team's active MV holds no assignment in the game
- **THEN** the MV receives one invitation under the ordinary member rules
- **AND** the MV appointment does not create a second invitation

#### Scenario: Games outside the seven-day date are not recruited for

- **WHEN** the daily job runs and games have required vacancies six or eight days later
- **THEN** those games receive no invitation from this seven-day flow

#### Scenario: Missing responsible team gives no recruitment recipients

- **WHEN** a game seven days away has vacancies but no responsible team
- **THEN** no team vacancy invitation is sent for that game

#### Scenario: Required vacancies follow the game category

- **WHEN** an adult game has its five baseline positions filled but Kasse and Wischer are empty
- **THEN** its invitation lists one Kasse and one Wischer vacancy
- **AND** a youth or Spielfest game with all five required positions filled sends no invitation for empty Kasse or Wischer duties

#### Scenario: A replacement search or age deficiency is not a vacancy

- **WHEN** every required position is occupied but a replacement search or an age or classification deficiency exists
- **THEN** no team vacancy invitation is sent
- **AND** current occupants and existing follow-up behavior remain intact

#### Scenario: Team-independent block gaps do not recruit a team

- **WHEN** a date has an empty preparation, cleanup or cake position and its games have no required vacancies
- **THEN** no responsible-team vacancy invitation is generated for those blocks

#### Scenario: Contact routing and skipped counts stay consistent

- **WHEN** qualifying recipients include one person with both contacts, one with phone only and one with neither
- **THEN** the first receives only e-mail and the second receives SMS
- **AND** the contactless person is skipped with a warning and contributes zero to the sent count

#### Scenario: Multiple memberships and identical display names preserve identity

- **WHEN** a qualifying member belongs to several teams and another qualifying person has the same display name
- **THEN** each internal person identity receives one invitation for the game
- **AND** neither extra memberships nor a shared name duplicates or suppresses a recipient

### Requirement: Same-game assignments exclude vacancy-invitation recipients

The system SHALL exclude a responsible-team member from a game's vacancy invitation if that person holds any saved assignment in that game, including retained removed-duty assignments and assignments seeking a replacement. Exclusion SHALL use the internal game and person identities and SHALL apply regardless of task, offeredness or MV/admin status. An assignment in another game or in a day block SHALL NOT by itself exclude the member. No additional individual age filter SHALL limit this team invitation; actual claims SHALL continue to enforce the existing eligibility and one-duty rules.

#### Scenario: An already assigned member keeps only their helper reminder

- **WHEN** an active member of the responsible team already holds any duty in the game with vacancies
- **THEN** that person receives no vacancy invitation for that game
- **AND** their ordinary weekly assigned-helper reminder remains available

#### Scenario: Retained or offered replacement duties also exclude the helper

- **WHEN** a member holds a retained removed-duty assignment or a duty marked Ersatz gesucht in the game
- **THEN** that member receives no invitation to take another duty in that game
- **AND** the existing assignment and replacement state remain unchanged

#### Scenario: Duties in another container do not exclude membership

- **WHEN** an active responsible-team member holds a duty in another game or a day block but no duty in the game being recruited for
- **THEN** the member remains a recipient for that game's invitation

#### Scenario: The invitation is not an eligibility authorization

- **WHEN** an active, unassigned member with an unknown birth date belongs to the responsible team
- **THEN** the member remains a recipient
- **AND** any subsequent claim is checked under the existing task-specific eligibility rules
