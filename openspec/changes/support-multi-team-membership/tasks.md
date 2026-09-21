## 1. Schema Revision Foundation

- [ ] 1.1 Add the pinned Alembic dependency, repository configuration, programmatic resolved-database wiring, and a reviewed no-op baseline revision for the exact current schema; verify the Alembic history has one head and configuration tests never fall back to an independent database URL.
- [ ] 1.2 Implement read-only schema-state inspection for empty, exact unversioned baseline, legacy game identity, known versioned, head, newer/divergent, corrupt, and near-miss unversioned databases; verify focused synthetic-database tests classify every state without creating an `alembic_version` table or changing the file.
- [ ] 1.3 Separate explicit fresh initialization from normal runtime verification so only an absent/empty target receives `create_all`, head stamping, Supporter seeding, and validation; verify web and daily startup accept head, reject missing/behind/ahead schemas with actionable diagnostics, and leave rejected databases byte/schema unchanged.
- [ ] 1.4 Add `manage_db.py migrate-schema --confirm-stopped` around the schema preflight, SQLite-API snapshot, baseline adoption, ordered Alembic upgrade, and postflight checks; verify CLI tests cover missing confirmation, exact-baseline adoption, already-at-head no-op, legacy-game instructions, unknown-schema refusal, backup failure, upgrade failure, and reporting of the retained backup path.

## 2. Membership Schema and Domain Model

- [ ] 2.1 Add the `person_teams` association with composite uniqueness and foreign keys, replace singular ORM relationships with `Person.teams`/`Team.persons`, and add deterministic membership ID/name/label helpers; verify database tests cover zero, one, several, and duplicate submitted memberships.
- [ ] 2.2 Create the reviewed membership Alembic revision using SQLite batch operations to copy legacy active `persons.team_id` and pending `desired_team_id` references into a deduplicated membership union, remove both old columns, and validate retained IDs and relationship counts; verify upgrade tests cover active users, pending registrations, records with both/neither reference, contactless users, assignments, audits, admins, and MVs.
- [ ] 2.3 Implement an atomic membership-set replacement helper that validates the entire requested set and clears MV appointments whose qualifying membership is removed; verify invalid team IDs roll back the full change and MV removal commits membership and appointment changes together.
- [ ] 2.4 Update person deletion, deactivation, reactivation, team-member lookup, and MV appointment helpers for membership sets; verify deactivation retains all memberships while clearing every MV role, reactivation does not restore MV roles or assignments, and deletion removes only the target person's membership rows.
- [ ] 2.5 Update registration persistence, MV-created users, admin-created users, seeding conveniences, and privacy cleanup so pending registrations retain inactive selected memberships and account status alone gates roster use; verify authentication, contactless-person, cleanup, and lifecycle tests cover pending, approved, rejected, inactive, and active membership states.

## 3. Roster and Access Control

- [ ] 3.1 Replace singular person/team view-model fields with ordered membership IDs, summaries, and combined labels while preserving person-ID identity and contact visibility; verify duplicate-name and no-team rendering tests show deterministic complete labels without exposing contacts or IDs.
- [ ] 3.2 Change the admin person editor and creation form to accept atomic multi-team selections, add MV controls that add/remove active people only for one managed roster at a time, and keep MV creation limited to one managed initial team; verify management UI tests cover admin add/retain/remove-all, MV managed-team add/remove, self-removal refusal, inactive targets, invalid IDs, and forged cross-team submissions.
- [ ] 3.3 Update roster team filtering and MV selection lists to use membership containment and active-member rules; verify a person with two teams matches either filter, appears as an MV candidate for both, and a non-member/inactive MV appointment is rejected.
- [ ] 3.4 Make verified self-registration approval and rejection admin-only, remove pending-registration visibility and decision rights from MVs, and notify an active admin with the selected team labels; verify the direct refusal suite rejects every MV registration decision while admin approval atomically activates all selected memberships.
- [ ] 3.5 Change the registration form and signed challenge context to require and preserve one or more validated team IDs through both server-rendered steps without client-side scripting; verify authentication tests cover several teams, no team, unknown/forged teams, duplicate selections, contact anti-enumeration, approval, and rejection.

## 4. Task Selection and Schedule Presentation

- [ ] 4.1 Refactor per-game candidate classification to consume membership sets with playing, responsible, Supporter, other precedence and keep normalized name/person-ID tie-breaking; verify focused helper tests cover all categories, every overlap, no memberships, no responsible team, and duplicate names.
- [ ] 4.2 Update schedule slot status, outside/playing warnings, option labels, selected occupants, and emitted sort metadata to use complete memberships; verify server-rendered tests retain one-task exclusion, current-occupant visibility, complete team labels, and the playing warning when any membership matches the playing team.
- [ ] 4.3 Update the browser release/reinsertion path for the revised membership labels and server-provided category metadata without client-side recategorization; verify the DOM-level regression returns released people to the correct position with labels and warning classes intact.
- [ ] 4.4 Update MV candidate restriction and claim validation to accept any active person whose memberships contain the game's responsible team while preserving member self-service, compare-and-swap, and one-task-per-game rules; verify webapp, refusal, and concurrency tests cover multi-team candidates and stale claims/releases.

## 5. Statistics, Notifications, and CLI

- [ ] 5.1 Update per-person statistics, assignment displays, and any notification/log context that shows a person's team to use deterministic complete membership labels while leaving game responsibility and MV notification routing unchanged; verify statistics and notifier regressions cover multi-team and no-team people.
- [ ] 5.2 Keep CLI creation's single-team/Supporter convenience and add ID-safe commands for inspecting and changing an existing person's membership set; verify CLI tests cover multiple memberships, unknown teams, MV clearing, duplicate names, and no name-based mutation fallback.
- [ ] 5.3 Remove remaining production singular-membership and desired-team reads and writes from Python, templates, and JavaScript while retaining `Game.team_id`; verify a targeted source search finds no obsolete `Person.team_id`, `Person.desired_team_id`, `person.team`, `person.desired_team`, or singular person view-model usage.

## 6. Operational Documentation and Verification

- [ ] 6.1 Update `README.MD` and operator guidance for explicit initialization, revision checks, the stopped-services schema migration, legacy game-identity prerequisite, backup retention, post-migration verification, and snapshot rollback; verify every documented command matches CLI help and the duty-block feature remains out of scope.
- [ ] 6.2 Add migration drift verification against a freshly initialized head database and test that the ORM metadata requires no additional upgrade operations; verify the automated Alembic check succeeds at head and fails for an intentional synthetic model/schema mismatch.
- [ ] 6.3 Run focused migration, database, authentication, access-control, management, schedule, concurrency, notifier, privacy, and CLI tests; verify all pass offline with synthetic databases and no production configuration reads.
- [ ] 6.4 Run `test/run_tests.sh` and `openspec validate support-multi-team-membership --strict`; verify the complete suite and strict change validation pass with `DEBUG_FLAG = False` and `CHANGE_DAY = False`.
