# nuLigaHelper

A helper class to organize home games with nuLiga Handball.

Host-specific deployment configuration and operator records are maintained privately outside this repository.

# Features

- Read home games from Hallenspielplan website of nuLiga
- Games, persons and task assignments are stored in a SQLite database (`nuliga_helper.db`)
- The database is backed up to a dropbox account after every run
- Tasks (Zeitnehmer, Sekretär, Verkauf, Ordnungsdienst, Reinigung) can be assigned to persons including notification addresses
- nuLigaHelper keeps track of all game shifts
- Notifications are sent to all game judges one day prior to the game (if the script is run daily)
- Notification articles can be sent to newspapers one week before the game
- Notifications on missing referees are sent to specified persons
- Unknown new games are reported to the admin
- Notifications can be sent via Email or SMS (using Twilio account)
- Passwordless Flask web interface with registration, tiered access, self-service assignments and an audit trail

# Web interface

For local development or trusted-LAN use, create a persistent signing secret once,
then start the Flask development server:

```bash
python3 -c 'import secrets; print(secrets.token_urlsafe(48))' > .nuligahelper_secret
./run_webapp.sh
```

It is then reachable in the local network at `http://<device-ip>:8080`.
`run_webapp.sh` is development-only and must not be exposed to the public internet.
Public hosting requires a separately maintained production deployment with TLS,
restricted secrets, monitoring and a reviewed launch procedure.

- **Heimspielplan**: public, read-only schedule with assigned helper names. Use the filters **Spielende Mannschaft**, **Verantwortliche Mannschaft** and **Eingeteilte Person** together to find relevant games; the person field searches assigned names, and **Filter löschen** restores the full plan. Past game days appear muted under **Vergangene Spieltage anzeigen** and can be expanded. Signed-in members can claim and release their own future tasks. MVs can additionally staff their own team on games for which that team is responsible; admins can correct every slot and set the responsible team. Concurrent stale changes are rejected rather than overwriting another person's work.
- **Helfer verwalten**: members see names and complete team memberships but only their own contact data. Admins create, edit, deactivate/reactivate or delete erroneous entries, replace membership sets, approve user registrations and appoint team MVs. MVs can add or remove active people only in rosters they manage; they do not see or decide pending registrations.
- **Statistik**: games covered per team, jobs per person (incl. role breakdown) and a list of upcoming games with missing assignments.
- **Änderungsprotokoll**: admin-only, append-only history of assignment claims and releases with actor, tier, person, game, role, slot and timestamp.

Registration and login each use one two-step form: request a six-digit code, then enter it on the same page. During registration, E-Mail and SMS are individually optional, but at least one valid contact is required and both may be supplied. The selected route alone receives the verification code, while all supplied contacts are stored; after approval, either stored route can be selected independently for login. Registration additionally requires a name, one or more selected teams and consent to publish the name when assigned. Successful contact verification queues the person and all selected teams as one user-level approval and notifies every active administrator; it does not itself approve the account. When an administrator approves the registration, the new active user receives a welcome notification inviting them to sign in and take open Heimspielplan duties. Both automatic notifications prefer E-Mail and use SMS only when no usable e-mail address is stored. Team membership itself has no separate confirmation. E-mail addresses are syntax-checked and stored in a canonical lowercase form. Phone input combines the selected country calling code with the national number and is stored as E.164 (for example `+491701234567`). Codes are single-use and valid for 15 minutes; sessions expire after one hour of inactivity. Active accounts land on the Heimspielplan after login, while verified registrations awaiting approval see only their status.

Access is derived from database facts on every request:

- **Guest**: public read-only schedule, without roster/contact payloads.
- **Member**: roster without others' contact data, statistics, own profile and own assignments.
- **MV**: member rights plus staffing and active-roster maintenance for each team they manage.
- **Admin**: all administration and the audit trail. Bootstrap the first admin with `manage_db.py grant-admin PERSON_ID`.

Teams are fully automatic: they are derived from the scraped age classes (`ak`, e.g. "BL mD") plus one support team ("Supporter"). A person may belong to no team, one team or several teams, without player/staff role distinctions. Teams cannot be created or edited manually. The responsible team is chosen per game; it is never pre-assigned, since the team playing the game should not be pre-booked for its own support tasks. Membership in the playing team always produces the advisory "spielt selbst" warning; membership outside the responsible and Supporter teams produces an advisory outside-team warning. Both remain selectable.

Schema revisions are managed by pinned Alembic migrations, but application startup never creates or upgrades schema implicitly. `manage_db.py init` is only for an absent or empty database; web and daily startup require the configured database to be at the current revision. Existing databases are upgraded only through the guarded, stopped-services `migrate-schema` workflow below. A database with the older source-key game identity must first use the separate `migrate-game-identity` prerequisite.

Authentication does not make the Flask development server suitable for public use.
The host-specific Gunicorn, systemd, Caddy and environment configuration is
maintained outside this repository.

## SQLite production profile and backups

All application database connections enable foreign-key enforcement, wait up to five seconds for transient locks and use `synchronous=FULL`. Startup establishes WAL journal mode and refuses to continue if WAL, the connection settings or `PRAGMA foreign_key_check` cannot be verified. Existing foreign-key violations are reported for explicit operator repair or restore; they are never deleted automatically.

The supported deployment uses the database on a local netcup VPS filesystem, exactly one web process/worker, and at most one daily-job process. Multi-worker WSGI operation, network-filesystem SQLite and PostgreSQL are outside the currently verified topology. WAL creates temporary `<database>-wal` and `<database>-shm` files while connections are active. Do not copy, move or delete these files while the application is running.

The daily job takes a non-blocking OS lock at `<resolved-database>.daily.lock` before database revision verification or external work. A concurrent run for the same database fails immediately with a nonzero exit status; an unlocked file left after a restart is harmless. `run.sh` uses `exec`, so cron receives the Python process status directly.

Backups use SQLite's online backup API and therefore include one complete committed state even when committed pages still reside in WAL. The live `.db` file and its sidecars are never uploaded directly. Every snapshot must pass `PRAGMA quick_check` and `PRAGMA foreign_key_check` before publication. Dropbox receives, in this order:

1. `<database-stem>-YYYY-MM-DD<suffix>` for `common.effective_today()`;
2. the stable configured database basename as the latest backup;
3. retention cleanup of only older names that exactly match that database's dated convention.

`NULIGAHELPER_DROPBOX_DATED_RETENTION` is required in production and must be a positive integer. A same-day rerun replaces that day's recovery point. Retry-safe Dropbox operations recover from transient connection, DNS, timeout, rate-limit and server failures with three bounded retries after 1, 2 and 4 seconds; authentication, bad-input and other permanent failures fail immediately. Snapshot, validation, read, Dropbox client, upload, listing, pagination, deletion and local-cleanup failures are logged concisely by stage and make the daily job exit unsuccessfully. When the source database remains usable, due notification work is still attempted after a backup failure, but notifications are not retried automatically.

### Restore a downloaded snapshot

Restoration is an offline operation. Copying only a live WAL-mode main file is not a valid backup or restore method.

1. Stop the web process and disable or wait for the daily job. Verify that no process is using the database or holding its daily-run lock.
2. Download the selected latest or dated Dropbox object to a separate path. Never download it over the active database.
3. Run the guarded command against the configured target:

   ```bash
   python3 manage_db.py --db /absolute/path/nuliga_helper.db \
       restore-snapshot /path/to/downloaded-snapshot.db --confirm-stopped
   ```

4. The command validates the candidate before changing the target, fsyncs a same-filesystem temporary copy, preserves the current `.db`, `-wal` and `-shm` set in the printed rollback directory, and atomically installs the snapshot. If installation fails, it attempts to recover the complete preserved set. Keep all services stopped and inspect the reported paths after any failure.
5. Start exactly one application process. Normal startup re-establishes WAL and repeats the runtime and foreign-key checks. Verify the schedule, authentication and one non-destructive database query before re-enabling normal traffic and the daily timer.
6. Retain the command-created rollback directory until the restored service and the next validated Dropbox backup have been accepted.

Before deploying these settings to an existing database, stop writers, preserve the current database together with any sidecars, run the full offline suite, and start one process to perform the integrity preflight. For rollback, stop every database user first and restore a validated snapshot or the complete preserved database set; never replace only the main file while WAL users are active.

### Migrate legacy game identity

Databases created before canonical game-number identity contain a non-null
`games.source_key` column. New application startup detects that schema and exits with
migration instructions; it never upgrades production data implicitly.

1. Stop the web process and daily timer/job and verify no database writer remains.
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

### Initialize or migrate the application schema

Fresh installation is explicit and only accepts an absent or empty target:

```bash
python3 manage_db.py --db /absolute/path/nuliga_helper.db init
```

Normal web and daily startup perform read-only revision classification before opening the
database. They refuse a missing, unversioned, behind, unknown, divergent or corrupt
schema and never call `create_all` or Alembic upgrade implicitly.

To upgrade an existing database:

1. Stop the web service and daily timer/job and verify no database user remains.
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
   postflight checks. Unknown and near-miss schemas are refused without adoption.
5. Retain the printed `Backup:` path. Start one application process and verify startup,
   several multi-team roster entries, MV appointments, assignments, audit history and the
   schedule before re-enabling normal traffic and the daily timer.

If migration or verification fails, keep services stopped. Restore the printed snapshot
with `restore-snapshot ... --confirm-stopped`, redeploy the previous application version,
and inspect the retained failed database separately. Membership revisions are not rolled
back in place.

`NULIGAHELPER_SECRET` is mandatory for both the webapp and daily job. Set it in the environment or keep the raw value in the gitignored `.nuligahelper_secret` file used by both launch scripts. Keep it stable: changing it invalidates every session. `NULIGAHELPER_DB` is an optional SQLite-path override locally; production requires `NULIGAHELPER_ENV=production`, an absolute `NULIGAHELPER_DB`, and exact `NULIGAHELPER_TRUSTED_HOSTS` entries.

# Testing

The `test/` folder contains offline tests for the database, notifications, authentication, authorization, CLI, concurrent claims and web flow. Every test uses synthetic secrets and throwaway databases:

```bash
test/run_tests.sh
```

or individually via `python3 test/test_db.py` etc., or with pytest (`pytest test/ -v`). See `test/README.md` for details.

# Database schema

- `games`: one task-relevant plan entry identified by `(season_year, game_nr)`; ordinary
  numbers are decimal text, while SPF matches on one date for one full age group collapse
  into a pseudo number such as `SPF:2026-11-28:spf mini`
- `teams`: derived automatically from the scraped age classes plus one support team; each team optionally references its MV (Mannschaftsverantwortlicher)
- `persons`: stable internal identity, mutable/non-unique name, contacts, account state and admin flag
- `person_teams`: deduplicated many-to-many membership set; pending registrations use the same rows, while account status gates roster use
- `assignments`: one stable role/slot link between a game and person
- `auth_tokens`: expiring, single-use registration/login nonces and six-digit e-mail/SMS codes
- `assignment_audit`: append-only assignment-change snapshots that remain readable after deletion

The MV of a game's responsible team receives an MV notification as long as tasks of that game are still unassigned.

# Getting started

Clone the repository and open `config_template.json`. Modify the club data and notification texts, then save it as `config.json`. Email, Twilio and Dropbox settings belong in environment variables mapped in `common.PROVIDER_ENV`. Production refuses provider sections in `config.json`; local legacy files remain readable during transition. Install all libraries from `requirements.txt`; this includes the pinned `email-validator` and `phonenumbers` packages used for deterministic, offline-safe contact validation.

Initialize the database and manage persons, games and assignments via the management CLI:

```
python3 manage_db.py init
python3 manage_db.py add-person "Max Mustermann" --team "Supporter" --email max@example.com
python3 manage_db.py list-persons
python3 manage_db.py show-memberships 1
python3 manage_db.py set-memberships 1 "BL mD" "BL mC"
python3 manage_db.py contact-preflight
python3 manage_db.py search-person "Max"
python3 manage_db.py grant-admin 1
python3 manage_db.py list-games
python3 manage_db.py list-games --number 12034
python3 manage_db.py assign 42 Zeitnehmer 1
python3 manage_db.py set-jteam 42 "Damen 1"
```

Run `contact-preflight` before deploying contact normalization to an existing database; it reports invalid values, canonical changes and collisions without writing anything. Person IDs are shown by `list-persons` and `search-person`; names are display data and may repeat. Existing-person commands (`show-memberships`, `set-memberships`, `assign`, `unassign`, `set-mv`, `grant-admin`) therefore take the person ID. `set-memberships` atomically replaces the complete set and accepts no team names to clear it; CLI creation keeps the convenient single `--team` option. Canonical game numbers are unique within a season: `list-games` shows the internal game ID together with its schedule context, and `assign`, `unassign` and `set-jteam` take that game ID. `--number` accepts either an ordinary decimal number or an SPF pseudo number and only filters the list.

Games are inserted automatically on the first run after scraping. Run the daily job via `./run.sh` (or export `NULIGAHELPER_SECRET` before `python3 main.py`). nuLigaHelper will scrape all available data from the nuLiga Hallenspielplan, sync it into the SQLite database and send all due notifications. A backup of the database is uploaded to Dropbox after every run.
All actions are logged in the file `helper.log`.
Run the daily job through one scheduler only; disable any old scheduler before enabling a replacement.

Placeholder games or duty blocks before the first and after the last game of a day are
deliberately outside this change and remain future work.

## Private deployment

The application exposes a fixed `/healthz` readiness route and public
`/impressum` and `/datenschutz` routes. Legal-page content, privacy retention
settings, host monitoring, service units and launch evidence are supplied by the
private deployment. Missing or invalid legal content returns a generic 503 page.
The private deployment uses a short cleanup policy (five retention periods and
batch size), two public legal pages, and a launch record with one club officer's
decision. The launch record checks the restore rehearsal, privacy notice and
retention settings; technical host checks run separately.

Production uses a stable `NULIGAHELPER_SECRET`, an absolute `NULIGAHELPER_DB`
shared by the web app, daily job, cleanup and backup,
exact `NULIGAHELPER_TRUSTED_HOSTS`, and provider settings from environment
variables. Production `config.json` contains club data and notification texts,
not email, Twilio or Dropbox provider sections. Host setup and public launch
require separate operator review.
