## Context

See `proposal.md` for motivation and the delta specs for behavioral requirements. The current ORM stores `Person.team_id`, while registration staging separately stores `desired_team_id` and each `Team` optionally stores `mv_person_id`. Team membership is read directly throughout database helpers, authorization, roster management, schedule rendering, statistics, CLI operations, privacy cleanup, and notification context.

Database startup currently combines SQLite runtime verification, `Base.metadata.create_all()`, a narrow legacy-game-schema check, `PRAGMA user_version`, foreign-key validation, and Supporter seeding. The only data migration is the reviewed, backup-first game-identity command. Production uses a single local SQLite database in WAL mode with one web process and at most one daily job, and destructive schema work must remain an offline operator action.

The current task candidate order already computes a per-game category and transports sort metadata into the browser. Multi-team membership changes the category predicate but not the compare-and-swap assignment model.

## Goals / Non-Goals

**Goals:**

- Establish one durable, source-controlled schema revision chain for all future ordinary schema evolution.
- Adopt only an exact known existing schema and preserve the project's backup-first, stopped-services, fail-closed operational posture.
- Represent person/team membership without role distinctions and expose deterministic membership sets everywhere one team is currently shown or evaluated.
- Preserve person identity, assignments, audits, contacts, account state, registration routing, and MV appointments during migration.
- Make membership and MV invariant changes atomic and keep authorization derived on every request.

**Non-Goals:**

- Reimplementing the historical game-identity transformation as an Alembic revision.
- Automatically upgrading an existing database during web or daily-job startup.
- Supporting in-place downgrade as the production rollback mechanism.
- Allowing users to create or edit teams, allowing ordinary members to edit memberships after registration, or allowing MVs to change memberships outside teams they manage.
- Distinguishing player, coach, staff, or other membership roles.
- Adding before/after-game duty blocks, synthetic games, or configurable task types.

## Decisions

### Use Alembic from a reviewed current-schema baseline

Add Alembic as the schema revision mechanism and keep its configuration inside the repository. The environment will receive the already resolved SQLite target programmatically rather than relying on a second database-path configuration source. Configure SQLite batch rendering, but manually review every generated operation and hand-author data transformations.

The first revision is a baseline marker for the exact pre-membership schema. Its purpose is adoption, not replaying the project's undocumented schema history. The next revision creates and populates person/team memberships and removes the legacy column. Later features receive later revisions.

An alternative was another standalone migration module like the game-identity conversion. That would solve this deployment but retain no ordered schema history, no standard head check, and no drift detection for the next structural change. Reconstructing all historical schemas as executable revisions was also rejected because it adds risk without improving adoption of the one known production state.

### Separate explicit creation, runtime verification, and offline upgrade

Split the current `init_db` responsibilities conceptually into three paths:

1. Explicit initialization accepts only an absent or demonstrably empty target, creates the current ORM schema, stamps the current Alembic head, seeds required data, and validates it.
2. Normal web/daily startup configures and verifies SQLite runtime properties, verifies schema head and integrity, and seeds nothing or changes no schema.
3. The guarded migration command recognizes, backs up, adopts if necessary, upgrades, and validates an existing stopped database.

Fresh creation will use the current SQLAlchemy metadata followed by an Alembic head stamp. This is simpler and faster than replaying a growing revision history, and an automated metadata-versus-head check will prove that both routes converge. `create_all()` must never run merely because an existing application process starts.

An alternative was automatic `upgrade head` at startup. It was rejected because SQLite table rebuilds, WAL sidecars, backups, and two independent application entry points require an explicit maintenance window and actionable operator control.

### Recognize source states before any migration write

The migration wrapper will classify the target using read-only SQLite inspection:

- exact unversioned current baseline: eligible for backup, baseline stamp, and upgrade;
- known Alembic revision at or behind this application's head: eligible for backup and ordered upgrade;
- already at head: report a validated no-op;
- legacy `games.source_key`: refuse and direct the existing game-identity command first;
- absent/empty target: direct explicit initialization;
- unknown, newer, divergent, corrupt, or foreign-key-invalid state: refuse with diagnostics.

The baseline fingerprint will cover expected application tables, relevant columns, constraints/indexes required to interpret data, the canonical game identity, and invariants needed by the membership copy. Merely finding `persons.team_id` is not sufficient evidence to stamp.

`PRAGMA user_version` remains historical input for recognizing the game-identity state but ceases to be the general schema version. Alembic's version table becomes authoritative after adoption.

### Wrap Alembic in existing backup and validation safeguards

Expose a project command such as `manage_db.py migrate-schema --confirm-stopped`; operators do not need to invoke raw Alembic commands in production. Before stamping or upgrading, the wrapper creates a self-contained snapshot through SQLite's backup API and validates it. It reports the source revision, planned target, and backup path.

The migration uses Alembic's SQLite batch operations where a table must be rebuilt. SQLite foreign-key enforcement is disabled only on the dedicated stopped migration connection for the required rebuild window, then restored before postflight. Postflight checks the revision, `quick_check`/integrity, foreign keys, copied membership counts, orphan absence, MV membership, and retained relationship counts.

SQLite DDL is not treated as fully transactional. If an upgrade or postflight fails, the supported recovery is to keep all users stopped and restore the reported validated snapshot. Downgrade functions are not presented as the production rollback route.

### Represent membership with one association table and no role metadata

Introduce `person_teams` with `person_id` and `team_id` foreign keys and a composite primary key. Cascading deletion of an erroneous person removes their association rows; automatic team deletion remains outside the product model. ORM relationships become `Person.teams` and `Team.persons` through this table. Shared helpers will return deterministic ordered team IDs/names instead of relying on collection load order.

Pending registrations store their freely selected teams in the same `person_teams` association. Account status determines whether those rows are inactive registration selections or active roster memberships, so a second desired-team association and per-team approval state are unnecessary. `Person.desired_team_id` is removed by the migration. `Team.mv_person_id` remains separate because MV is an appointment, not a membership subtype.

An association object with a player/staff/coach field was rejected because the clarified domain treats every reason for belonging identically. Keeping `Person.team_id` as a preferred or primary team was rejected because it would recreate ambiguous precedence and two sources of truth.

### Make membership set replacement atomic and preserve MV invariants

Central database helpers will validate a submitted set of existing team IDs before mutating the ORM collection. Registration and admin forms/APIs submit repeated `team_ids` values; public registration requires at least one, while an empty set remains valid for administrative maintenance of existing records. Invalid IDs reject the entire write.

Admins use complete-set replacement. MVs use a separate scoped add/remove operation that verifies the target person is active and the target team is in `g.mv_team_ids`, then changes only that one membership. MVs cannot change pending registrations, replace a person's set, touch another team's membership, or remove their own qualifying membership; an admin must perform a change that would alter an MV appointment.

When an admin membership removal would leave an appointed MV outside that team, the helper clears the appointment in the same transaction. Deactivation retains memberships but clears all MV appointments, matching existing lifecycle behavior. Reactivation does not recreate appointments. MV selection lists draw from active memberships.

An alternative was rejecting removal until the MV was manually replaced. Atomic clearing better matches deactivation, guarantees the invariant after every write, and avoids trapping an administrator in an invalid edit sequence.

### Store free multi-team registration selections behind one admin approval gate

Self-registration requires at least one team and accepts several. The validated selection is written to `person_teams` when the pending person is created, but the account status keeps the person and memberships out of the roster, selections, statistics, authorization, and assignments. Contact verification queues one user-level decision for admins; no MV or per-team decision exists. Admin approval changes the person to active and thereby activates the complete stored membership set atomically. Rejection activates none and follows the existing rejected-account lifecycle.

Only admins see and decide pending registrations, and the approval action does not edit the selected teams. An admin can correct memberships through ordinary roster administration after approval. MVs may then adjust only their own managed roster. Direct MV-created roster persons remain a separate administrative convenience and receive one active membership in the chosen managed team without using public registration.

Privacy cleanup uses account status and retention timestamps rather than interpreting membership presence as approval. Registration notifications target an active admin and include the selected team labels without exposing contacts.

### Evaluate candidate categories from sets with playing precedence

Expose each candidate's ordered `team_ids` and display names to schedule construction. For each game, compute exactly one category:

1. playing team occurs in `team_ids`;
2. otherwise responsible team occurs in `team_ids`;
3. otherwise Supporter occurs in `team_ids`;
4. otherwise other/no team.

The presentation order remains responsible, Supporter, other, playing, so the numeric sort groups map the first predicate to the last group. Outside warnings apply only when a responsible team exists and neither responsible nor Supporter membership exists. Server-rendered options continue to carry category/name/person metadata; browser reinsertion clones it without independently interpreting memberships.

MV authorization filters candidates by membership in that game's responsible team. Member self-service remains identity-based and is unaffected except for advisory classification. One-task-per-person and slot compare-and-swap constraints do not change.

### Normalize multi-team presentation at view-model boundaries

Replace singular `team_id`/`team_name` person dictionaries with ordered `team_ids`, structured team summaries where needed, and one deterministic combined label for templates and notification/statistics context. Roster team filtering uses membership containment. Assignment and audit identity remains `Person.id`; team labels are display data only.

CLI creation keeps a convenient single `--team` initial membership and Supporter default, while explicit admin-style set-membership and MV-style scoped add/remove operations use person IDs and existing team names/IDs. Existing-person mutation never falls back to a display-name identity.

## Risks / Trade-offs

- [Baseline stamping accepts a subtly incompatible database] -> Compare an explicit reviewed schema fingerprint and domain invariants, test near-miss schemas, and refuse rather than infer.
- [SQLite batch rebuild leaves a partially changed file] -> Require stopped services, create and validate a backup first, validate postflight, and make snapshot restore the documented rollback.
- [Normal startup still mutates an old database through `create_all()`] -> Remove schema creation from runtime paths and test that old/missing schemas fail unchanged.
- [Membership checks diverge across authorization and presentation] -> Centralize membership IDs/containment helpers and cover direct forged requests as well as rendered options.
- [An overlapping membership receives the wrong suggestion] -> Apply the explicit playing, responsible, Supporter, other precedence once on the server and preserve emitted sort metadata in JavaScript.
- [MV appointment becomes detached from membership] -> Mutate collections only through invariant-preserving helpers and verify the invariant in migration postflight and tests.
- [MV roster editing broadens into general administration] -> Authorize each add/remove against one managed team, require an active target, forbid complete-set writes and self-removal, and cover forged cross-team requests.
- [Pending team selections leak an unapproved person into the roster] -> Gate every membership consumer by active account status and test roster, MV, schedule, statistics, and notification queries with verified pending registrations.
- [Displaying several teams makes controls crowded] -> Use deterministic compact labels in cards/options and a multi-select/checklist only in the admin editing state; keep contacts out of schedule responses.
- [Active unrelated schema work changes the adoption baseline] -> Build and review the baseline against the repository schema at implementation time, coordinate revision ordering, and validate exact source states rather than relying on a hard-coded table count alone.

## Migration Plan

1. Add Alembic configuration, the reviewed baseline revision, schema-state inspection, explicit initialization, runtime head verification, and migration-wrapper tests without yet removing the old membership accessors.
2. Add the membership revision: create `person_teams`, copy legacy active-team and pending desired-team references into a deduplicated union, validate counts and MV membership, and rebuild `persons` without `team_id` or `desired_team_id`.
3. Switch ORM relationships and all application consumers to membership sets; update multi-team registration, admin-only approval, scoped MV roster editing, forms, view models, authorization, privacy cleanup, CLI, notifications, statistics, and tests.
4. Verify fresh initialization produces a database at head with no metadata drift. Verify exact-baseline adoption, already-versioned upgrades, no-op head handling, legacy-game refusal, near-miss refusal, migration failure, and snapshot recovery paths entirely offline.
5. Run the complete test suite with both debug switches false and update operator documentation.
6. For production, stop the web service and daily timer, preserve the live database set, run `migrate-schema --confirm-stopped`, retain the printed backup, start one process, and verify roster memberships, MV appointments, assignments, audit history, and schedule rendering before restoring normal operation.

Rollback after any migration failure or rejected postflight is an offline restore of the command-created validated snapshot followed by deployment of the previous application version. No in-place reverse membership migration is supported.
