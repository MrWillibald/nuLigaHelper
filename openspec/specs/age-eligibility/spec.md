# Age Eligibility Specification

## Purpose

Defines how stored birth dates determine eligibility for age-restricted game duties, including individual task minimums and the requirement for an adult within the explicitly assigned Verkauf group.

## Requirements

### Requirement: Birth dates are validated without inventing legacy data

Every newly created person SHALL have a valid calendar birth date that is not after the application's effective current day. This SHALL apply to public registration, administrative creation, MV creation, and CLI creation. Existing people whose birth date is unknown SHALL retain their identity and other data, SHALL be identifiable for authorized completion, and SHALL NOT receive an inferred or placeholder birth date. An unknown legacy date SHALL NOT prevent unrelated maintenance. Once a valid date is stored, ordinary maintenance SHALL NOT clear it.

#### Scenario: New person lacks a birth date

- **WHEN** a creation request omits the birth date
- **THEN** creation is refused with an explanatory validation error
- **AND** no partial person record is created

#### Scenario: Invalid or future birth date is submitted

- **WHEN** creation or correction supplies an impossible date or a date after the effective current day
- **THEN** the write is refused and previously stored data remains unchanged

#### Scenario: Existing date is unknown

- **WHEN** an existing person has no birth date
- **THEN** the person and existing assignments remain stored
- **AND** authorized maintenance can complete the date without replacing the person's identity

#### Scenario: Legacy contact data needs repair

- **WHEN** an authorized user corrects unrelated data on a legacy record whose unknown birth date remains unchanged
- **THEN** the correction can be saved without inventing a date

### Requirement: Age is evaluated on the scheduled game date

Eligibility SHALL use completed calendar years on the game's scheduled date rather than age on the assignment or notification date. A person SHALL reach an inclusive minimum on the corresponding birthday. A 29 February birthday SHALL advance on 1 March in a non-leap year. If the game's date cannot be determined, a new age-restricted claim SHALL be refused with an unresolved-date reason.

#### Scenario: Birthday occurs before a future game

- **WHEN** a person is below a task's minimum today but reaches it on or before the scheduled game date
- **THEN** the age check permits the claim if the other assignment rules allow it

#### Scenario: Game precedes the required birthday

- **WHEN** the game date is before the birthday on which the person reaches the minimum
- **THEN** the age check refuses the claim

#### Scenario: Leap-day birthday in a non-leap year

- **WHEN** a person born on 29 February is evaluated on 28 February and on 1 March of a non-leap year
- **THEN** their completed age advances on 1 March

#### Scenario: Scheduled game date is unknown

- **WHEN** a caller claims an age-restricted game duty whose date is missing or invalid
- **THEN** the claim is refused without falling back to the current day

### Requirement: Timing duties have game-category minimum ages

Zeitnehmer SHALL require a known birth date and completed age of at least 14 for youth games or at least 18 for adult games. Sekretär SHALL require a known birth date and completed age of at least 14 for youth games or at least 16 for adult games.

Adult M/F categories SHALL be recognized as adult, youth mA–mE/wA–wE categories and SPF SHALL be recognized as youth, and unsupported categories including GE SHALL remain unresolved rather than silently selecting a threshold. A new Zeitnehmer or Sekretär claim with an unresolved category SHALL be refused with an explanatory reason.

#### Scenario: Youth timing duties meet their minimum

- **WHEN** a person is exactly 14 on a youth game's date and is claimed for Zeitnehmer or Sekretär
- **THEN** the age check permits the claim

#### Scenario: Youth timing helper is under 14

- **WHEN** a person has not reached 14 on a youth game's date
- **THEN** a Zeitnehmer or Sekretär claim is refused

#### Scenario: Adult Zeitnehmer reaches 18

- **WHEN** a person turns 18 on the adult game's date
- **THEN** the Zeitnehmer age check permits the claim
- **AND** a person who is still 17 fails it

#### Scenario: Adult Sekretär reaches 16

- **WHEN** a person turns 16 on the adult game's date
- **THEN** the Sekretär age check permits the claim
- **AND** a person who is still 15 fails it

#### Scenario: Timing candidate lacks a birth date

- **WHEN** a Zeitnehmer or Sekretär claim names a person with an unknown birth date
- **THEN** the claim is refused because eligibility cannot be established

#### Scenario: Game category cannot be classified

- **WHEN** a caller claims Zeitnehmer or Sekretär for a game with an unsupported or missing category
- **THEN** the claim is refused with an unresolved-classification reason

### Requirement: Verkauf has collective adult coverage

A completely occupied Verkauf group SHALL contain at least one explicitly assigned person with a known birth date who is 18 or older on the game date, eligible from their 18th birthday. This rule SHALL apply to all game categories.

The rule SHALL NOT impose an individual minimum on every seller. A younger person or a legacy person with an unknown birth date SHALL be permitted in an incomplete Verkauf group. A claim filling the last free Verkauf slot SHALL be refused unless its resulting group contains a qualifying adult. Team membership, responsibility, or an unassigned candidate SHALL NOT satisfy this requirement.

#### Scenario: One younger seller is assigned first

- **WHEN** a younger person is claimed for an empty Verkauf group and another sale slot remains free
- **THEN** the claim is permitted under the sale age rule
- **AND** adult coverage remains outstanding

#### Scenario: Younger seller works alongside an adult

- **WHEN** one seller is at least 18 on the game date and a younger seller is claimed for the remaining sale slot
- **THEN** the claim is permitted under the sale age rule

#### Scenario: Only younger sellers would complete the group

- **WHEN** a claim would fill the final sale slot while all assigned sellers are younger than 18 on the game date
- **THEN** the claim is refused and stored assignments remain unchanged

#### Scenario: Adult birthday is reached

- **WHEN** a seller is evaluated on the day before their 18th birthday and on their 18th birthday
- **THEN** they do not satisfy adult coverage on the first date
- **AND** they satisfy it on the birthday

#### Scenario: Unknown birth date is not proof of adult qualification

- **WHEN** a final-slot claim would leave no seller with a known qualifying age
- **THEN** the claim is refused even if one assigned seller's age is unknown

#### Scenario: Responsible team has adult members

- **WHEN** adults belong to the responsible team but none is explicitly assigned to Verkauf
- **THEN** they do not satisfy the sale-group requirement

### Requirement: Every claim enforces current eligibility atomically

The system SHALL enforce age rules for claims made through every entry point and access tier, including administrators and CLI operations. Authorization and existing slot/person uniqueness rules SHALL still apply. Eligibility SHALL be evaluated against authoritative game, person, and post-claim group state within the same serialized write transaction, and SHALL be reevaluated after a contention retry. A refused claim SHALL alter neither assignments nor successful-mutation audit history.

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

### Requirement: Withdrawals retain existing rights

An otherwise authorized release SHALL remain allowed when it removes the only qualifying adult seller or an age-ineligible current occupant. Releasing the qualifying seller SHALL leave adult coverage outstanding and SHALL refresh the remaining group's eligibility. Existing past-game restrictions and compare-and-swap checks SHALL remain applicable.

#### Scenario: Adult seller withdraws

- **WHEN** the only qualifying adult seller releases their assignment before the game
- **THEN** the release succeeds
- **AND** the remaining sale group has an outstanding adult-coverage requirement

#### Scenario: Ineligible occupant is removed

- **WHEN** an authorized caller releases a stored assignment whose occupant no longer meets its age rule
- **THEN** the eligibility deficiency does not prevent the release

### Requirement: Changed or incomplete data produces visible eligibility deficiencies

The system SHALL retain existing assignment rows after migration, birth-date corrections, and game-date or category changes. Current/future assignments SHALL be reevaluated using current stored data and SHALL expose underage, unknown-birth-date, unresolved-game-date, unresolved-category, or missing-adult-seller deficiencies as applicable.

Physical vacancy counts SHALL remain distinct from these deficiencies. Current/future games with an eligibility deficiency SHALL remain outstanding in staffing summaries, statistics, and existing responsible-team MV follow-up even when every physical slot is occupied. Historical assignment records and append-only audits SHALL NOT be rewritten by this reevaluation.

#### Scenario: Legacy full roster has missing dates

- **WHEN** a future game's timing slots are occupied by people whose birth dates are unknown
- **THEN** their assignments remain visible
- **AND** the game reports unresolved timing eligibility despite its full occupancy

#### Scenario: Game moves before a helper's birthday

- **WHEN** a game moves to a date before its timing helper reaches the minimum
- **THEN** the assignment remains stored and an age deficiency becomes visible

#### Scenario: Birth date correction changes sale coverage

- **WHEN** an authorized correction shows that no assigned seller qualifies
- **THEN** the correction is retained
- **AND** the game reports outstanding adult coverage

#### Scenario: Full invalid roster has a responsible MV

- **WHEN** existing follow-up evaluates a current/future game with no physical vacancies but an age deficiency and an active responsible-team MV
- **THEN** the game remains eligible for that existing MV follow-up

#### Scenario: No responsible MV exists

- **WHEN** a deficient game has no active responsible-team MV
- **THEN** the deficiency remains visible
- **AND** this change does not invent a replacement assignee or notification recipient
