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
team and all existing task fields, including optional `Unterstützung`; the expanded
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

## ADDED Requirements

### Requirement: Cake delivery appears as a compact dated task card

The overview SHALL show the date's cake-delivery block alongside its existing preparation, game and cleanup cards. The cake card SHALL start collapsed and SHALL show its label, configured delivery time or setup-needed state, requested cake quantity and staffing progress. Expanding it SHALL reveal one numbered volunteer position per requested cake and each position's assigned helper name. Each position SHALL represent one cake contribution. For a configured quantity, progress SHALL count occupied positions out of that quantity and SHALL update after successful assignments or configuration edits. A refused or stale operation SHALL not show uncommitted progress. Configured zero SHALL show that no cakes are requested without division-by-zero or invented staffing percentages.

The cake card SHALL use the established plain card treatment, a small 🍰 emoji beside its label and a labeled keyboard-operable expansion control. It SHALL have no responsible-team field. Administrators SHALL be offered controls for the date's delivery time and quantity; delivery time entry SHALL consistently use 24-hour `HH:MM` regardless of the browser locale. Other viewers SHALL see those settings without edit controls. An unconfigured block SHALL visibly require administrator setup and SHALL not offer claim controls.

#### Scenario: Configured cake block is expanded

- **WHEN** a viewer expands a cake block configured for four cakes at 10:00
- **THEN** the card shows that delivery time and quantity
- **AND** exactly four independently displayed cake positions are shown

#### Scenario: Cake block needs setup

- **WHEN** a cake block has no delivery time or requested quantity
- **THEN** it shows that administrator setup is needed
- **AND** no invented time, quantity or claimable volunteer positions are shown

#### Scenario: Administrator adjusts delivery settings

- **WHEN** an administrator saves valid new delivery settings
- **THEN** the card shows the saved time and requested quantity
- **AND** its position count and progress reflect the saved configuration

### Requirement: Cake cards follow day-block filtering and privacy

Cake cards SHALL participate in date and assigned-person filtering under the same rules as other day-task cards. A playing-team filter SHALL retain a cake card only for a date with a qualifying game. A responsible-team filter SHALL hide the cake card because it has no responsible team. A person-name filter SHALL retain the cake card only when one of its own assignments matches, and SHALL allow it to retain a qualifying date without a matching game card. Clearing filters SHALL restore cake cards in their normal upcoming or past date groups.

Guests SHALL see assigned names without roster payloads, person identifiers, contact data or assignment/configuration controls. Editable cake cards SHALL load their permitted candidates only when opened, retaining the existing conflict and candidate-failure behavior.

#### Scenario: Find a cake volunteer

- **WHEN** a person-name filter matches only a cake volunteer on a qualifying date
- **THEN** the date and cake card remain visible
- **AND** unrelated games and other day blocks are not shown merely because they share that date

#### Scenario: Responsible-team filter is selected

- **WHEN** a viewer selects a responsible-team filter
- **THEN** cake cards are hidden

#### Scenario: Guest views cake assignments

- **WHEN** a guest expands a configured cake card
- **THEN** the assigned helper names are visible
- **AND** the response contains no roster payload, person identifiers, contacts or editing controls

#### Scenario: Open an editable cake card

- **WHEN** a signed-in viewer opens an editable cake card
- **THEN** only that card loads its permitted candidates
- **AND** a failed request leaves its controls unable to submit claims based on an incomplete candidate list
