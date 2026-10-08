# Spec Delta

## ADDED Requirements

### Requirement: Roster entries disclose authorized maintenance on demand

On every viewport, each roster entry SHALL initially show the person's name, complete team membership set and any status already visible to the viewer. Authorized person maintenance controls SHALL start collapsed behind "Bearbeitung" on mobile and desktop, including without JavaScript, and SHALL be independently accessible by keyboard and touch through native disclosure. An affected entry returned with field validation errors SHALL start open. A viewer with no maintenance action for an entry SHALL see no empty editing disclosure. An MV's permitted team-membership actions SHALL remain available without granting profile-edit rights. Collapsing content SHALL NOT broaden permissions or add unauthorized contacts, birth dates or other private data to the response. The team-management section SHALL retain its existing position before the roster.

#### Scenario: Admin browses many helpers on mobile

- **WHEN** an administrator opens the roster on mobile with responsive enhancement available
- **THEN** person entries show names, every team and status while maintenance controls start collapsed
- **AND** opening one entry reveals only that entry's existing authorized controls

#### Scenario: Member and MV maintenance scopes

- **WHEN** a member or MV opens another person's maintenance disclosure
- **THEN** only actions already authorized for that viewer are available
- **AND** another person's contacts and full birth date remain absent from the response
- **AND** an ordinary member sees no editing disclosure for another person

#### Scenario: Validation failure in a person form

- **WHEN** an authorized person edit is rejected with field validation errors
- **THEN** the affected entry's maintenance controls start open on the returned page, including on mobile
- **AND** its submitted values and field errors remain available for correction
- **AND** unrelated entries retain their ordinary initial defaults

#### Scenario: Desktop and no-JavaScript maintenance

- **WHEN** a viewer uses the roster on desktop or without JavaScript
- **THEN** all authorized maintenance actions remain reachable
- **AND** maintenance controls start collapsed behind "Bearbeitung" and can be opened natively
- **AND** an affected entry returned with field validation errors still starts open

### Requirement: Roster filters have a compact mobile presentation

The existing roster filter form SHALL start collapsed behind "Filter" on viewports at most 700 CSS pixels wide when responsive enhancement is available and SHALL start expanded on wider initial viewports. Applied non-default filters SHALL remain visible outside the form in a readable German summary with a reset action. The summary SHALL include only filters permitted for the viewer and SHALL show human-readable values instead of internal identifiers. Form selections, result counts, membership-set matching and visibility rules SHALL remain unchanged. Opening or closing the form SHALL NOT submit it or change results. Filtering SHALL remain usable without JavaScript.

#### Scenario: Applied roster filters are visible

- **WHEN** a viewer applies permitted name, team, status or missing-birth-date filters on mobile
- **THEN** the collapsed form's outside summary identifies the applied values and offers reset
- **AND** the form retains its selected values and the roster contains only permitted matching records

#### Scenario: Lower tier requests admin filters

- **WHEN** a member or MV supplies an administrator-only filter
- **THEN** neither the summary nor the results disclose unauthorized information
