# Spec Delta

## ADDED Requirements

### Requirement: Birth dates have restricted visibility

An existing person's full birth date SHALL be visible only to that person and administrators in authorized person-maintenance contexts. MV status alone SHALL NOT permit reading or editing another existing person's birth date. Authorized MV creation SHALL accept the new person's birth date without granting continuing access to it.

Schedule pages, assignment-candidate responses, statistics, ordinary notifications, and assignment-audit snapshots SHALL NOT expose full birth dates or exact calculated ages. Where assignment feedback requires age information, the system SHALL return the applicable threshold or a derived eligibility reason instead.

#### Scenario: Member views their own profile

- **WHEN** a signed-in member opens their own editable person entry
- **THEN** their stored birth date is visible and can be corrected within existing ownership rights

#### Scenario: Member or MV views another roster entry

- **WHEN** a member or MV who is not an administrator views another existing person
- **THEN** that person's full birth date is absent from the response
- **AND** the viewer cannot edit it

#### Scenario: Administrator maintains the roster

- **WHEN** an administrator opens authorized person maintenance
- **THEN** stored and missing birth dates are available for completion and correction

#### Scenario: Assignment candidates are loaded

- **WHEN** any viewer loads a game's candidates or assignment feedback
- **THEN** the response communicates eligibility only as necessary for assignment
- **AND** it contains no full birth date or exact calculated age

#### Scenario: Guest views assigned helper names

- **WHEN** a guest opens the public schedule
- **THEN** assigned helper names remain visible under the existing public policy
- **AND** no birth dates, exact ages, or private roster data are exposed
