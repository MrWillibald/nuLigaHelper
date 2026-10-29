# Initialize or migrate the application schema

Schema revisions are managed by pinned Alembic migrations, but application startup never creates or upgrades schema implicitly. `manage_db.py init` is only for an absent or empty database; web and daily startup require the configured database to be at the current revision. Existing databases are upgraded only through the guarded, stopped-services `migrate-schema` workflow below. A database with the older source-key game identity must first use the separate `migrate-game-identity` prerequisite.

## Migrate legacy game identity

Databases created before canonical game-number identity contain a non-null
`games.source_key` column. New application startup detects that schema and exits with
migration instructions; it never upgrades production data implicitly.

1. Stop public ingress; disable the daily and cleanup timers, wait for running
   oneshot jobs, stop the web process, and verify no database user or open
   SQLite handle remains.
2. Run the guarded migration against the configured database:

   ```bash
   python3 manage_db.py --db /absolute/path/nuliga_helper.db \
       migrate-game-identity --confirm-stopped
   ```

3. The command performs a read-only preflight, reports ambiguous ordinary numbers or
   conflicting Spielfest task data without changing the database, and creates a dated
   SQLite-API backup before starting its transaction.
4. A successful migration converts ordinary numbers to text, collapses all SPF matches
   with the same date and full age group into one Spielfest, preserves safe game IDs,
   assignments and audit references, rebuilds the constraint as
   `(season_year, game_nr)`, and runs row-count, integrity and foreign-key checks.
5. Record the printed backup path, start one application process, verify the plan,
   assignments and audit page, then re-enable the daily timer.

If any step fails, keep every writer stopped. The database transaction is rolled back;
retain the failed file for diagnosis and restore the printed pre-migration backup before
redeploying the previous application. Do not attempt an in-place reverse migration.

## Initialize or upgrade a database

Fresh installation is explicit and only accepts an absent or empty target:

```bash
python3 manage_db.py --db /absolute/path/nuliga_helper.db init
```

Normal web and daily startup perform read-only revision classification before opening the
database. They refuse a missing, unversioned, behind, unknown, divergent or corrupt
schema and never call `create_all` or Alembic upgrade implicitly.

To upgrade an existing database:

1. Stop public ingress; disable the daily and cleanup timers, wait for running
   oneshot jobs, stop the web service, and verify no database user remains.
2. If the command reports legacy `games.source_key`, first complete
   `migrate-game-identity --confirm-stopped` as described above. Do not reverse this order.
3. Run the guarded schema migration:

   ```bash
   python3 manage_db.py --db /absolute/path/nuliga_helper.db \
       migrate-schema --confirm-stopped
   ```

4. The command accepts only the reviewed exact unversioned baseline or a known Alembic
   revision. It creates and validates a dated SQLite-API snapshot before stamping or
   upgrading, applies revisions in order, and runs revision, integrity and foreign-key
   postflight checks, including retained person and assignment data. Unknown and
   near-miss schemas are refused without adoption. The birth-date revision adds a
   nullable date-only column: all existing people retain unknown dates, with no
   inference from team, name or duties. Person IDs, contacts, membership sets,
   account states, MV appointments, assignments and audit snapshots are retained.
   The following cake-delivery revision extends day blocks and seeds exactly one
   unconfigured cake block per existing valid season/home-game date. It preserves
   preparation/cleanup block IDs, positions, assignments and audit snapshots, plus
   all birth dates and person relationships. No cake time, quantity or volunteer
   is inferred. Revision `0006_game_duty_staffing`
   renames current Unterstützung assignments to Kasse while preserving their
   IDs, game, person and numbered position; older original Reinigung rows traverse the
   earlier Reinigung-to-Unterstützung revision first. Historical audit labels
   and snapshots remain unchanged and readable, and the two Reinigung places introduced by that historical revision begin empty. Unterstützung/Kasse collisions, unexpected new-role rows and invalid
   old positions are refused without merging or discarding them. Youth/unknown
   Kasse assignments become retained release-only assignments, as specified in
   [game-task-staffing](../../openspec/specs/game-task-staffing/spec.md). Revision `0007_binding_duty_assignments` adds
   nullable replacement tokens and saved lock timestamps to game/block assignments
   plus the `duty_cutoffs` season/date cutoff marker table. Existing occupants, positions,
   personal data and historical audit snapshots remain unchanged. Migration creates
   no replacement requests, inferred confirmations or historical reminder cutoffs;
   an existing database's first authoritative runtime checks enforce the
   fixed 09:00 Europe/Berlin deadline independently of a reminder run. All upgrades require this stopped-service migration before startup.
   Revision `0008_single_wischer` is the preceding duty migration. It renames current
   Reinigung position 0 to Wischer, preserving its assignment ID, helper, saved lock
   and replacement token. As the explicitly requested **one-time migration exception**
   to commitment preservation, it deletes **every current Reinigung position 1** across
   all seasons, dates and categories, including past, locked and replacement-seeking
   duties. It never moves the second helper into the first position. Deleted duties
   disappear from personal counts (including past counts) and ordinary reminders.
   Each deletion appends one system `migration_remove` audit with the original role,
   position, helper and game; all previous audits and snapshots stay unchanged. Original
   pre-day-block Reinigung still becomes Kasse through the historical chain. Unexpected
   Wischer target rows and unsupported source positions fail preflight; retained-data
   checks verify the exact deletion/rename set and every old plus appended audit row.
   No provider notifications are sent. Adult staffing now requires seven places.
   This exception grants no runtime release or cutoff override. Keep the printed backup:
   on failure, stop all writers and restore that snapshot and the prior application,
   which also restores deleted second duties; never downgrade in place.
   Revision `0009_schedule_visibility` is the single required current head. It adds
   exactly one unrestricted `schedule_visibility_setting` row with a nullable
   date-only boundary. It preserves every existing game, block, person, membership,
   assignment ID, lock, replacement token and historical audit. Earlier pending
   revisions retain their separately reviewed effects, including the one-time `0008`
   removal described above; `0009` removes no additional duties. Stop all web/daily
   database users, run `./venv/bin/python manage_db.py migrate-schema --confirm-stopped`,
   retain the printed validated snapshot and restart only after postflight succeeds.
   Startup refuses prior revisions without upgrading or seeding settings. Repeating
   migration at head preserves a configured boundary. Recovery requires stopped
   writers, the validated pre-upgrade snapshot and matching prior application;
   never downgrade the active database in place.
   Verification rejects unreviewed schema differences, even if a database already
   carries the current revision label. Previously recognized historical defaults
   remain unchanged, and SQLite uniqueness checks confirm reflected constraints.
5. Retain the printed `Backup:` path. Start one application process and verify startup,
   several multi-team roster entries, MV appointments, assignments, audit history and the
   schedule before deliberately re-enabling normal traffic and the daily/cleanup timers.
   Review the admin roster's missing-date indication, complete legacy dates through
   authorized person maintenance or `set-birth-date`, and review current/future
   eligibility warnings. Existing assignments with unknown dates remain stored;
   the warnings identify staffing that still needs verification or repair.
   Expand each **Kuchenlieferung** card as an admin and set its delivery time and
   quantity; newly seeded cards show setup needed and cannot accept claims.

If migration or verification fails, keep services stopped. Restore the printed snapshot
with `restore-snapshot ... --confirm-stopped`, redeploy the previous application version,
and inspect the retained failed database separately. Membership, birth-date and cake revisions
are not rolled back in place. Do not run migration or backfill against production data
as a release smoke test.


---
Back to the [README](../../README.MD) · [Operations index](README.md)
