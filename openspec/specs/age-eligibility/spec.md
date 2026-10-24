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

Zeitnehmer SHALL require a known birth date and completed age of at least 14 for youth games or at least 16 for adult games. Sekretär SHALL require a known birth date and completed age of at least 14 for youth games or at least 16 for adult games.

Adult M/F categories SHALL be recognized as adult, youth mA–mE/wA–wE categories and SPF SHALL be recognized as youth, and unsupported categories including GE SHALL remain unresolved rather than silently selecting a threshold. A new Zeitnehmer or Sekretär claim with an unresolved category SHALL be refused with an explanatory reason.

#### Scenario: Youth timing duties meet their minimum

- **WHEN** a person is exactly 14 on a youth game's date and is claimed for Zeitnehmer or Sekretär
- **THEN** the age check permits the claim

#### Scenario: Youth timing helper is under 14

- **WHEN** a person has not reached 14 on a youth game's date
- **THEN** a Zeitnehmer or Sekretär claim is refused

#### Scenario: Adult Zeitnehmer reaches 16

- **WHEN** a person turns 16 on the adult game's date
- **THEN** the Zeitnehmer age check permits the claim
- **AND** a person who is still 15 fails it

#### Scenario: Adult Zeitnehmer reaches 18

- **WHEN** a person turns 18 on the adult game's date
- **THEN** the Zeitnehmer age check still permits the claim
- **AND** a person who is still 17 also passes the revised minimum of 16

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

A completely occupied Verkauf group SHALL contain at least one explicitly assigned person with a known birth date who is 18 or older on the game date, eligible from their 18th birthday. This rule SHALL apply to adult games and unresolved game categories, including GE, missing and unsupported classes. Only recognized youth games, including SPF, SHALL be exempt.

For games subject to collective adult coverage, the rule SHALL NOT impose an individual minimum on every seller. A younger person or a legacy person with an unknown birth date SHALL be permitted in an incomplete Verkauf group. A claim filling the last free Verkauf slot SHALL be refused unless its resulting group contains a qualifying adult. Team membership, responsibility, or an unassigned candidate SHALL NOT satisfy this requirement.

For recognized youth games, Verkauf SHALL impose neither an individual minimum nor collective adult coverage. Either sale position SHALL accept younger helpers or legacy helpers with unknown birth dates, even when the group becomes full. Youth-sale age eligibility SHALL NOT require a valid game date; existing general date, authorization and binding-duty checks SHALL still apply.

#### Scenario: One younger seller is assigned first

- **WHEN** a younger person is claimed for an empty Verkauf group at a non-youth game and another sale slot remains free
- **THEN** the claim is permitted under the sale age rule
- **AND** adult coverage remains outstanding

#### Scenario: Younger seller works alongside an adult

- **WHEN** a non-youth game has one seller at least 18 on the game date and a younger seller is claimed for the remaining sale slot
- **THEN** the claim is permitted under the sale age rule

#### Scenario: Only younger sellers would complete the group

- **WHEN** a claim for a non-youth game would fill the final sale slot while all assigned sellers are younger than 18 on the game date
- **THEN** the claim is refused and stored assignments remain unchanged

#### Scenario: Adult birthday is reached

- **WHEN** a seller for a non-youth game is evaluated on the day before their 18th birthday and on their 18th birthday
- **THEN** they do not satisfy adult coverage on the first date
- **AND** they satisfy it on the birthday

#### Scenario: Unknown birth date is not proof of adult qualification

- **WHEN** a non-youth final-slot claim would leave no seller with a known qualifying age
- **THEN** the claim is refused even if one assigned seller's age is unknown

#### Scenario: Responsible team has adult members

- **WHEN** a non-youth game has adults in the responsible team but none is explicitly assigned to Verkauf
- **THEN** they do not satisfy the sale-group requirement

#### Scenario: Two younger youth sellers

- **WHEN** two otherwise eligible younger helpers claim the two Verkauf positions at a recognized youth game
- **THEN** both claims pass the sale age check without an assigned adult

#### Scenario: Unknown birth dates at a Spielfest

- **WHEN** otherwise eligible legacy helpers with unknown birth dates fill both Verkauf positions at an SPF game
- **THEN** sale age eligibility permits both assignments and produces no adult-coverage deficiency

#### Scenario: Unresolved class receives no youth exemption

- **WHEN** a claim would complete a Verkauf group without an adult at a game with GE, a missing class or an unsupported class
- **THEN** the claim is refused under the retained collective adult-coverage rule

#### Scenario: Invalid date remains unresolved outside youth sales

- **WHEN** a caller claims either Verkauf position at a non-youth game whose date is missing or invalid
- **THEN** sale age eligibility refuses the claim even while another sale position remains free

#### Scenario: Invalid date does not introduce a youth-sale age requirement

- **WHEN** sale age eligibility is evaluated for a recognized youth game whose date is missing or invalid
- **THEN** it produces no age deficiency or age-based refusal
- **AND** other assignment checks retain their existing date requirements

### Requirement: Every claim enforces current eligibility atomically

The system SHALL enforce age rules for new claims and responsibility takeovers made through every entry point and access tier, including administrators and CLI operations. A takeover SHALL evaluate the complete resulting group after replacing the original occupant; the outgoing person SHALL NOT count toward post-takeover collective adult Verkauf coverage where that requirement applies. Authorization and existing slot/person uniqueness rules SHALL still apply. Eligibility SHALL be evaluated against authoritative game, incoming person, and post-change group state within the same serialized write transaction, and SHALL be reevaluated after a contention retry. A refused claim or takeover SHALL alter neither the responsible occupant nor successful-mutation audit history. A refused takeover SHALL preserve the original occupant and replacement-request state. Opening a replacement request SHALL NOT vacate the slot or remove its occupant from current eligibility reporting. Binding protections and the day-before lock SHALL apply even when the existing occupant is ineligible.

#### Scenario: Administrator attempts an underage appointment

- **WHEN** an administrator claims a timing duty for a person below its minimum
- **THEN** the claim is refused without an age-rule override

#### Scenario: Concurrent younger-person sale claims

- **WHEN** two concurrent claims by nonqualifying sellers at a non-youth game target different free sale slots
- **THEN** they cannot both complete a sale group lacking an adult
- **AND** any rejected claim leaves the accepted assignment intact

#### Scenario: Candidate list became stale

- **WHEN** the group changes after candidates were loaded and a later claim no longer meets age rules
- **THEN** the server refuses the claim based on current stored state

#### Scenario: Takeover would remove the only qualifying adult seller

- **WHEN** a full Verkauf group at a non-youth game contains one qualifying adult and a younger person attempts to take over that adult's assignment
- **THEN** the takeover is refused because the resulting full group would lack adult coverage
- **AND** the original adult remains assigned and responsible with the replacement request unchanged

#### Scenario: Eligible seller takes over

- **WHEN** an authorized takeover before the day-before lock leaves the resulting Verkauf group compliant with current stored age data
- **THEN** the age rule permits the atomic transfer subject to all other assignment rules
- **AND** no intermediate vacancy or release of the original seller is stored

#### Scenario: Takeover eligibility changes during contention

- **WHEN** a takeover waits or retries and current stored data no longer permits the incoming person or resulting sale group
- **THEN** eligibility is reevaluated and the takeover is refused without replacing the occupant

#### Scenario: Youth-sale takeover removes the sole adult

- **WHEN** an authorized younger helper or legacy helper with an unknown birth date takes over the only adult seller's position at a recognized youth game before cutoff
- **THEN** the sale age check permits the transfer without an adult remaining
- **AND** confirmation, expected occupant and request token checks still apply

#### Scenario: Youth category changes before confirmation

- **WHEN** a youth-sale candidate list is loaded but the saved game is changed to an adult or unresolved category before a final-slot claim or takeover commits
- **THEN** current collective adult coverage is enforced under serialization and after retries
- **AND** a refused write preserves occupants, request state and successful-mutation audit history

#### Scenario: Ordnungsdienst age changes before confirmation

- **WHEN** a previously eligible Ordner candidate no longer reaches 18 under current saved game or person data when a claim or takeover commits
- **THEN** the write is refused for every tier and entry point, including admin and CLI
- **AND** saved assignments and replacement-request state remain unchanged

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

### Requirement: Changed or incomplete data produces visible eligibility deficiencies

The system SHALL retain existing assignment rows after migration, eligibility-policy updates, birth-date corrections, and game-date or category changes. Current/future assignments SHALL be reevaluated using current stored data and the current duty-specific rules and SHALL expose underage, unknown-birth-date, unresolved-game-date, unresolved-category, or missing-adult-seller deficiencies as applicable.

Physical vacancy counts SHALL remain distinct from these deficiencies. Current/future games with an eligibility deficiency SHALL remain outstanding in staffing summaries, statistics, and existing responsible-team MV follow-up even when every physical slot is occupied. Historical assignment records and append-only audits SHALL NOT be rewritten by this reevaluation.

#### Scenario: Legacy full roster has missing dates

- **WHEN** a future game's timing slots are occupied by people whose birth dates are unknown
- **THEN** their assignments remain visible
- **AND** the game reports unresolved timing eligibility despite its full occupancy

#### Scenario: Game moves before a helper's birthday

- **WHEN** a game moves to a date before its timing helper reaches the minimum
- **THEN** the assignment remains stored and an age deficiency becomes visible

#### Scenario: Birth date correction changes sale coverage

- **WHEN** an authorized correction shows that a non-youth game has no assigned seller who qualifies
- **THEN** the correction is retained
- **AND** the game reports outstanding adult coverage

#### Scenario: Full invalid roster has a responsible MV

- **WHEN** existing follow-up evaluates a current/future game with no physical vacancies but an age deficiency and an active responsible-team MV
- **THEN** the game remains eligible for that existing MV follow-up

#### Scenario: No responsible MV exists

- **WHEN** a deficient game has no active responsible-team MV
- **THEN** the deficiency remains visible
- **AND** this change does not invent a replacement assignee or notification recipient

#### Scenario: New Ordnungsdienst minimum exposes a retained deficiency

- **WHEN** a current/future Ordner occupant is below 18 or has an unknown birth date when the revised policy is applied
- **THEN** the existing assignment, responsibility, request state and any saved lock remain intact
- **AND** staffing summaries, statistics and existing MV follow-up include the applicable deficiency
- **AND** an age deficiency grants no release or lock override

#### Scenario: Relaxed requirements resolve obsolete deficiencies

- **WHEN** current/future staffing is reevaluated for a 16- or 17-year-old adult-game Zeitnehmer or for youth-sale staffing without an adult
- **THEN** the former age deficiency disappears if no other requirement is violated
- **AND** physical occupancy, assignment records and append-only audits remain unchanged

#### Scenario: Category changes alter sale coverage requirements

- **WHEN** a current/future game changes from recognized youth to an adult or unresolved category without changing its sale assignments
- **THEN** collective adult coverage is reevaluated and any deficiency becomes visible
- **AND** changing back to recognized youth removes that sale age deficiency without altering assignments

### Requirement: Ordner requires an adult helper

Every new Ordner claim or takeover SHALL require a known birth date and completed age of at least 18 on the scheduled game date. This minimum SHALL apply to adult games, recognized youth games including SPF, and games with unresolved categories. A missing or invalid game date SHALL refuse new Ordner assignments with an unresolved-date reason. Unknown birth dates SHALL refuse new assignments with an unknown-birth-date reason. Full birth dates and exact personal ages SHALL remain private; public eligibility reasons SHALL identify only the duty, required minimum or unresolved condition.

#### Scenario: Ordnungsdienst reaches 18

- **WHEN** a helper turns 18 on a game's scheduled date
- **THEN** the Ordner age check permits the assignment
- **AND** the same helper fails the check for a game on the preceding day

#### Scenario: Ordnungsdienst applies to every category

- **WHEN** an underage helper claims Ordner for an adult, youth, SPF or unresolved-category game
- **THEN** the age check refuses the claim with a minimum-age-18 reason

#### Scenario: Unknown Ordnungsdienst birth date

- **WHEN** a new Ordner claim or takeover names a legacy helper with an unknown birth date
- **THEN** the write is refused because adult eligibility cannot be established

#### Scenario: Ordnungsdienst game date is unresolved

- **WHEN** the game's date is missing or invalid
- **THEN** a new Ordner claim or takeover is refused without substituting today's date

#### Scenario: Adult helper for an unresolved category

- **WHEN** a helper reaches 18 by a valid game date but the game's category is unresolved
- **THEN** the Ordner age check permits the assignment because its threshold is category-independent

#### Scenario: Public reasons use the current name

- **WHEN** an Ordner claim or takeover is refused for age, unknown birth date or an invalid game date, or a retained occupant has an eligibility deficiency
- **THEN** the human-readable reason identifies Ordner and the applicable threshold or unresolved condition
- **AND** it exposes no full birth date or exact personal age
- **AND** any machine-readable role identifier remains compatible with existing clients
