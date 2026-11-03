# Schema Migrations Delta

## MODIFIED Requirements

### Requirement: Existing unversioned databases are adopted only from an exact baseline

The guarded schema-migration command SHALL adopt an unversioned database only when its schema and required invariants exactly match the reviewed baseline. It SHALL record that baseline without replaying historical creation and then apply later revisions. A database with the legacy game-identity schema SHALL be refused unchanged with a diagnostic stating that this release no longer converts legacy game identity and that an earlier release providing the specialized game-identity migration must be used first; every other unknown unversioned shape SHALL be refused.

#### Scenario: Current unversioned database matches the baseline

- **WHEN** preflight finds the exact supported pre-versioning schema and data invariants
- **THEN** the command records the baseline revision
- **AND** applies pending revisions without recreating existing application data

#### Scenario: Legacy game identity is still present

- **WHEN** preflight or normal startup finds the historical game source-key schema
- **THEN** it makes no change
- **AND** the diagnostic states that the legacy game-identity conversion is no longer part of this release and must be completed with an earlier release before retrying
- **AND** it does not name a command that this release does not provide

#### Scenario: Unversioned schema is not recognized

- **WHEN** preflight finds an unversioned database that differs from every accepted source schema
- **THEN** adoption is refused with diagnostic differences
- **AND** the database remains unversioned and unchanged
