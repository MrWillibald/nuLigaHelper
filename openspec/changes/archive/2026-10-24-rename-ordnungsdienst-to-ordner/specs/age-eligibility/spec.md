# Spec Delta

## RENAMED Requirements

- FROM: `### Requirement: Ordnungsdienst requires an adult helper`
- TO: `### Requirement: Ordner requires an adult helper`

## MODIFIED Requirements

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
