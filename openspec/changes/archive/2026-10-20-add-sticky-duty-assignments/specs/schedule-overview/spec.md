# Spec Delta

## MODIFIED Requirements

### Requirement: Every game and day-task card starts compact

Every game, preparation block, and cleanup block SHALL start collapsed on page load,
regardless of whether a game has a responsible team or occupied slots. A collapsed game
card SHALL show its time, matchup or Spielfest identity, existing game metadata, its
responsible-team name or an open placeholder when unset, and a staffing progress bar.
A collapsed block card SHALL show its calculated time, block
label, and staffing progress bar. Each card SHALL have a labeled control that expands
and collapses that card independently. The expanded game SHALL reveal the responsible
team and all existing task fields, according to the game's offered set, plus occupied removed duties marked as existing assignments subject to binding protection and the explicit legacy retirement policy; the expanded
preparation or cleanup block SHALL reveal all three numbered assignment fields. All viewers SHALL be able to
expand cards, while assignment controls and private data SHALL continue to depend on
their existing access rights. Expansion SHALL be usable with a keyboard.

#### Scenario: Open and close a game

- **WHEN** a visitor opens the overview containing a game without a responsible team
- **THEN** that game starts collapsed like every other game and shows the open responsible-team placeholder
- **AND** expanding it reveals the responsible-team field and every task field
- **AND** collapsing it hides those fields again

#### Scenario: Responsible team remains visible in a collapsed game

- **WHEN** a visitor opens the overview containing a game with a responsible team
- **THEN** its collapsed summary shows that team's name without opening the card
- **AND** expanding it reveals the responsible-team field and every task field

#### Scenario: Open a day-task block

- **WHEN** a visitor expands a preparation or cleanup card
- **THEN** all three numbered assignment fields and their visible occupants are shown
- **AND** other cards retain their current expansion states

#### Scenario: Youth duties and retained assignments

- **WHEN** a viewer expands a youth game
- **THEN** only the five offered baseline positions appear as game duty fields
- **AND** empty Kasse and Reinigung fields and optional markers are absent
- **AND** any retained removed-duty occupant appears separately as an existing assignment with only controls allowed by binding protection and the explicit legacy retirement policy

### Requirement: Compact cards show accurate staffing progress

Each game progress bar SHALL count occupied positions in that game's offered/required-position set and SHALL exclude retained removed-duty positions and responsible-team selection from both numerator and denominator. Adult games SHALL have eight required positions; youth games SHALL have five. Each preparation and cleanup progress bar SHALL count
occupied slots among that block's three slots. Every bar SHALL have a readable count
and whole-number percentage, including at zero and full coverage, so meaning does not
depend on color alone. A successful permitted occupancy change SHALL update the affected
card's progress without requiring a page reload; an unsuccessful or stale operation
SHALL not show uncommitted progress.

#### Scenario: Optional support does not complete a game

- **WHEN** a youth game has three occupied required slots and an occupied retained removed Kasse slot
- **THEN** its compact card reports 3 of 5 required slots filled and 60 percent

#### Scenario: Block coverage

- **WHEN** one of a preparation block's three slots is occupied
- **THEN** its compact card reports 1 of 3 slots filled and 33 percent

#### Scenario: Assignment changes progress

- **WHEN** a viewer successfully claims a position or completes an atomic takeover on an expanded card
- **THEN** that card's count, percentage, and bar length reflect the saved assignment
- **AND** a rejected assignment leaves the displayed progress at its saved value

#### Scenario: Adult additional duties count

- **WHEN** an adult game has five of its eight required positions occupied
- **THEN** its compact card reports 5 of 8 required positions filled

#### Scenario: Retained removed duty released

- **WHEN** a retained Kasse or Reinigung assignment of a youth game is explicitly retired by an authorized caller before the fixed day-before cutoff
- **THEN** its required-duty progress is unchanged and the removed duty field disappears
- **AND** changing its sole Ordnungsdienst position updates progress

#### Scenario: Replacement wanted is still occupied

- **WHEN** a replacement request opens or closes without takeover
- **THEN** physical occupancy and the progress bar stay unchanged
- **AND** an atomic takeover changes responsibility without reducing progress

## ADDED Requirements

### Requirement: Assignment commitment and replacement state are visible

Every game and day-block duty SHALL show its assigned helper and persistent `Ersatz gesucht` status while an unlocked request is active, and SHALL explain when an occupied assignment is fixed. Guests SHALL receive only public names and read-only duty status without person IDs, roster payloads, contacts, full birth dates, exact ages or editing controls. Authorized claim and takeover flows SHALL provide the pre-save commitment confirmation defined by task self-service. Locked occupied duties SHALL not expose a release or replacement action; genuinely empty offered positions SHALL retain permitted claim actions. Existing filters, expansion behavior, responsive presentation and staffing-deficiency indications SHALL remain intact.

#### Scenario: Guest sees a replacement request

- **WHEN** a guest views an assignment seeking replacement
- **THEN** its helper name and `Ersatz gesucht` status remain visible
- **AND** no private identifiers, candidates or editing controls appear

#### Scenario: Cancel confirmation on any card

- **WHEN** a user cancels a claim confirmation in a game or any block phase
- **THEN** the card displays authoritative saved state without suggesting a claim succeeded

#### Scenario: Fixed and empty positions coexist

- **WHEN** a date has reached cutoff with some occupied and some genuinely empty duties
- **THEN** occupied duties show fixed state and permitted empty duties remain claimable
