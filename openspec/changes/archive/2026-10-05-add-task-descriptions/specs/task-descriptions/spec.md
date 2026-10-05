# Spec Delta

## Purpose

Explains the work associated with every supported assignment task through consistent, accessible descriptions beside schedule task labels.

## ADDED Requirements

### Requirement: Every supported task has a description

The system SHALL provide a concise German description for every supported game and day-level assignment task. Each numbered position SHALL use the description of its semantic role. Descriptions SHALL explain the task's purpose or work without inventing requirements that are not part of the club's task definition. A task introduced or renamed by another change SHALL retain complete description coverage when that change is adopted.

#### Scenario: Current tasks have coverage

- **WHEN** the current game and day-level tasks are displayed
- **THEN** Zeitnehmer, Sekretär, Verkauf, Ordnungsdienst, Kasse, Reinigung, Vorbereitung, Aufräumen and Kuchenlieferung each have a description

#### Scenario: Added or renamed duties have coverage

- **WHEN** a supported duty is introduced or renamed
- **THEN** its description is included in the actual supported role catalog
- **AND** obsolete labels are not presented as separate current tasks

#### Scenario: Numbered positions share meaning

- **WHEN** a viewer reads descriptions for Verkauf 1 and Verkauf 2
- **THEN** both positions explain the same Verkauf duty
- **AND** the same rule applies to other numbered roles

### Requirement: Description information follows each task label

Every displayed assignment-task label on the schedule SHALL be followed immediately by a small, recognizable information control. The control SHALL make the corresponding description available in a floating popover over underlying content without changing the assignment, moving other page content or expanding an unrelated card. Information controls SHALL not be confused with claim, release or task-edit actions.

#### Scenario: Expanded assignments show information

- **WHEN** a guest or signed-in viewer expands a game or day-task card
- **THEN** each visible assignment-task label is immediately followed by its information control
- **AND** activating that control reveals the matching description without submitting an assignment

#### Scenario: Help floats without shifting the schedule

- **WHEN** a viewer opens a task description by hover, focus, click or tap
- **THEN** the description floats over the underlying content within the viewport
- **AND** the assignment fields, card and other page content retain their positions
- **AND** the description is not clipped by the card container

### Requirement: Descriptions work with pointer keyboard and touch

Task descriptions SHALL be available on pointer hover and keyboard focus and through click or tap on the information control. Keyboard users SHALL be able to open and dismiss the description without changing assignments. Each information control SHALL have an accessible name identifying its task and SHALL be associated with the description it exposes. Descriptions SHALL remain readable within narrow-screen layouts.

#### Scenario: Pointer reads a description

- **WHEN** a viewer hovers over a task's information control
- **THEN** the task description becomes readable

#### Scenario: Keyboard reads and dismisses a description

- **WHEN** a keyboard user focuses a task's information control
- **THEN** its description is available to that user
- **AND** the user can dismiss an opened description with the keyboard without changing an assignment

#### Scenario: Touch reads a description

- **WHEN** a viewer taps a task's information control on a narrow screen
- **THEN** its description is shown in a readable area
- **AND** a further tap or the provided dismissal interaction closes it

#### Scenario: Long description on a small screen

- **WHEN** the description exceeds the available viewport height
- **THEN** the viewer can scroll its text inside the help bubble
- **AND** repositioning the bubble retains its reading position

### Requirement: Task descriptions are public generic information

Descriptions SHALL be available to guests and signed-in viewers alike. A description SHALL contain task guidance rather than roster data, personal contacts or information about an individual assignee. Adding descriptions SHALL preserve existing guest-response and assignment-authorization guarantees.

#### Scenario: Guest opens task help

- **WHEN** a guest reads a task description
- **THEN** the public description is visible
- **AND** no person identifiers, unassigned roster or contact data are added to the response
- **AND** no assignment mutation becomes available
