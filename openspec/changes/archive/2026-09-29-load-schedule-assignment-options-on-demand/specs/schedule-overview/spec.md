## ADDED Requirements

### Requirement: Editable cards load candidates when opened

The initial schedule response SHALL show the current assignment state without including lists of unassigned candidates in collapsed or expanded card markup. When a signed-in viewer opens an editable game or day-task card, its candidate controls SHALL become usable after the permitted candidates have loaded. Opening one card SHALL not load candidates for other cards. The summary and assigned helper names SHALL remain available before candidate loading completes.

#### Scenario: Initial overview with a growing roster

- **WHEN** active, unassigned people are added to the roster and the same schedule is requested
- **THEN** the initial response does not repeat those people across task controls
- **AND** its candidate payload remains absent regardless of roster size

#### Scenario: Open one editable game

- **WHEN** a signed-in viewer opens an editable game card
- **THEN** that card loads its permitted candidates and makes its task controls usable
- **AND** unopened cards do not load their candidates

#### Scenario: Open an editable day-task block

- **WHEN** a signed-in viewer opens a preparation or cleanup card with editable slots
- **THEN** that block loads its permitted candidates independently of games and other blocks

#### Scenario: Candidate loading fails

- **WHEN** a candidate request fails or the session has expired
- **THEN** the affected controls remain unable to submit an assignment based on an incomplete list
- **AND** the viewer sees a German error or sign-in message and can retry after recovery
- **AND** current assignments remain visible and unchanged
