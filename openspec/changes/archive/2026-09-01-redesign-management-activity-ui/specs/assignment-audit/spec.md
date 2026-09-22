## MODIFIED Requirements

### Requirement: Admins can review the record

The system SHALL show the recorded entries to admins in the web interface, most recent
first, and SHALL let them narrow the list down to a single game or a single person. The
review interface SHALL use the application's established visual language and SHALL present
entries in a readable form on both desktop and mobile screens. It SHALL remain read-only.

#### Scenario: Admin reviews recent changes

- **WHEN** an admin opens the record
- **THEN** the most recent changes are listed first with time, actor, affected person, game
  and task
- **AND** the page uses the same section, card, filter, button and typography treatment as
  the rest of the web interface

#### Scenario: Admin investigates one game

- **WHEN** an admin filters the record by a game
- **THEN** only the changes concerning that game are listed

#### Scenario: Admin investigates one person

- **WHEN** an admin filters the record by a person
- **THEN** only the changes concerning that person as actor or affected person are listed

#### Scenario: Activity entries are readable on mobile

- **WHEN** an admin opens the record on a narrow screen
- **THEN** each entry is presented as a readable stacked item or card rather than requiring
  horizontal scrolling across the full desktop table
- **AND** all audit fields remain available

#### Scenario: Activity review remains read-only

- **WHEN** an admin reviews the record
- **THEN** no action is offered that edits or deletes an existing entry

#### Scenario: Non-admin has no access

- **WHEN** a guest, a member or an MV requests the record
- **THEN** access is refused
