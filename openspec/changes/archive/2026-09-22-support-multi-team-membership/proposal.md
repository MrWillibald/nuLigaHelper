## Why

People can belong to more than one team, for example by playing for one team while coaching or serving as MV for another. The current single `Person.team_id` relationship cannot represent that reality and therefore produces incomplete roster management, MV eligibility, authorization, and task-candidate suggestions; changing it safely also requires a repeatable schema-migration mechanism now that production data must be preserved.

## What Changes

- Replace each person's optional single team with a many-to-many membership between persons and the automatically managed teams.
- Treat every membership alike: membership in the team playing a game produces the existing advisory warning and final candidate category, even when the person also belongs to the responsible or Supporter team.
- Update roster display, filtering, administration, registration approval, MV appointment, MV staffing rights, statistics, assignment labels, and CLI operations to work with membership sets.
- Let self-registering users select one or more teams freely; the selections stay hidden and unusable until an admin approves the user registration, at which point all selected memberships become active together.
- Remove MV registration approval and instead let an MV add or remove active roster persons only for teams they manage, while administrators retain full membership-set control.
- Introduce Alembic as the general schema revision framework, with a reviewed baseline and a data migration that copies every legacy `Person.team_id` into the membership table before removing the column.
- Replace implicit schema mutation on normal startup with revision compatibility checks and actionable fail-closed migration instructions. Fresh database initialization remains explicit.
- Add a guarded, backup-first, offline schema-migration command that accepts only recognized source states and validates the upgraded SQLite database.
- Retain the existing specialized game-identity migration as a prerequisite for older databases rather than rewriting its historical transformation into Alembic.
- **BREAKING**: internal person/team APIs, templates, CLI behavior, and ORM access change from one `team_id`/`team` value to a collection of memberships.
- Exclude day-boundary duty blocks and synthetic before/after-game entries; those remain a separate future change.

## Capabilities

### New Capabilities

- `team-membership`: Many-to-many person/team membership, membership administration, display, filtering, lifecycle, and migration guarantees.
- `schema-migrations`: Versioned, operator-controlled, backup-first database initialization, compatibility checking, and schema upgrades.

### Modified Capabilities

- `user-accounts`: Self-registration selects multiple teams, only admins approve user registrations, and approval activates every selected membership.
- `access-control`: MV authorization, scoped roster editing, and roster visibility/filtering use membership sets rather than one team.
- `task-self-service`: Task eligibility, MV staffing scope, candidate ordering, and advisory warnings evaluate all of a person's memberships.

## Impact

- Database and ORM: `db.py`, a new person/team association table, removal of `persons.team_id`, Alembic configuration and revisions, schema-state startup checks, and the existing `PRAGMA user_version` handling.
- Operations: `manage_db.py`, explicit initialization and schema-upgrade commands, backup and validation integration, deployment documentation, and migration tests against known legacy/current states.
- Web UI and APIs: `webapp.py`, roster and registration views, multi-team administration controls, MV selectors, schedule candidate metadata, statistics, and the corresponding templates and JavaScript.
- Notifications and CLI: team-member lookup and displayed team context while preserving `Team.mv_person_id` as the authoritative MV appointment.
- Tests and documentation: database, authentication, authorization, management, schedule, notification, privacy, CLI, migration, and complete offline-suite coverage; README and operator procedure updates.
- Dependencies: Alembic becomes a runtime/operations dependency alongside SQLAlchemy.
