# nuLigaHelper

A helper class to organize home games with nuLiga Handball.

Host-specific deployment configuration and operator records are maintained privately outside this repository.

# Features

- Read home games from Hallenspielplan website of nuLiga
- Games, persons and task assignments are stored in a SQLite database (`nuliga_helper.db`)
- The database is backed up to a dropbox account after every run
- Game tasks (Zeitnehmer, Sekretär, Verkauf, Ordnungsdienst, optional Unterstützung) can be assigned to persons including notification addresses
- Every home-game date has three-slot Vorbereitung and Aufräumen blocks, timed 90 minutes before the first and 60 minutes after the last game start; slot numbers are display positions while statistics, reminders and audits use the shared task name, just like Verkauf
- Every home-game date also has a Kuchenlieferung block; admins set its delivery time and requested quantity, with one volunteer position per cake
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

- **Heimspielplan**: public, read-only schedule with assigned helper names. Use **Datum** to show one game day or all days, together with **Spielende Mannschaft**, **Verantwortliche Mannschaft** and **Eingeteilte Person**. A person search keeps only games and day-task blocks where that person is assigned; selecting a responsible team hides day-task blocks because they have no responsible team. **Filter löschen** restores the full plan. Every game and day-task block starts compact; cake cards awaiting setup or requesting zero cakes show their status, while configured duties show a staffing bar; compact game cards also show the responsible team or **– offen –**. Expand a card to see its assignments and, for games, the editable responsible-team field when permitted. Assignment candidates load when an editable card opens; a loading error leaves its dropdowns disabled and offers a retry. Unopened cards carry no candidate lists, so registrations do not multiply the initial page size. Game progress counts the five required task slots; optional **Unterstützung** remains assignable but does not affect the bar. Each preparation and cleanup bar counts three slots. Past game days appear muted under **Vergangene Spieltage anzeigen**; selecting a past date opens that section. Signed-in members can claim and release their own future tasks. MVs can additionally staff their own team on games for which that team is responsible; admins can correct every slot and set the responsible team. Concurrent stale changes are rejected rather than overwriting another person's work.
- **Helfer verwalten**: members see names and complete team memberships but only their own contact data and birth date. Admins create, edit, deactivate/reactivate or delete erroneous entries, replace membership sets, complete missing birth dates, approve user registrations and appoint team MVs. MVs can create people with a birth date and add or remove active people only in rosters they manage; they do not see another existing person's birth date or decide pending registrations.
- **Statistik**: games covered per team, jobs per person (incl. role breakdown) and upcoming games and day blocks with missing assignments, unresolved age eligibility or cake setup still needed.
- **Änderungsprotokoll**: admin-only, append-only history of assignment claims and releases with actor, tier, person, game, role, slot and timestamp.

Registration and login each use one two-step form: request a six-digit code, then enter it on the same page. During registration, E-Mail and SMS are individually optional, but at least one valid contact is required and both may be supplied. The selected route alone receives the verification code, while all supplied contacts are stored; after approval, either stored route can be selected independently for login. Registration additionally requires a name, valid birth date, one or more selected teams and consent to publish the name when assigned. The birth date is validated before requesting a code and preserved through verification. Successful contact verification queues the person and all selected teams as one user-level approval and notifies every active administrator; it does not itself approve the account. When an administrator approves the registration, the new active user receives a welcome notification inviting them to sign in and take open Heimspielplan duties. Both automatic notifications prefer E-Mail and use SMS only when no usable e-mail address is stored. Team membership itself has no separate confirmation. E-mail addresses are syntax-checked and stored in a canonical lowercase form. Phone input combines the selected country calling code with the national number and is stored as E.164 (for example `+491701234567`). Codes are single-use and valid for 15 minutes; sessions expire after one hour of inactivity. Active accounts land on the Heimspielplan after login, while verified registrations awaiting approval see only their status.

New people need a real calendar birth date that is not after `common.effective_today()`, including contactless admin/MV creation and CLI creation. A member can complete or correct their own date; an admin can do so for any person. A known date cannot be cleared. Existing people with an unknown date can still receive unrelated profile repairs. Dates are self-reported; the application does not verify identity documents. Full dates are limited to own/admin person maintenance and are absent from schedules, candidates, statistics, ordinary notifications, diagnostic logs and assignment audits. See the [data inventory](DATA-INVENTORY.md).

## Age eligibility for game duties

Age means completed calendar years on the scheduled game date, including future games. A minimum is met on the birthday: a person turning 14 on 5 September can perform a youth timing duty on 5 September, but not on 4 September. A 29 February birthday advances on 1 March in a non-leap year.

| Duty | Youth games | Adult games |
| --- | --- | --- |
| Zeitnehmer | 14 or older | 18 or older |
| Sekretär | 14 or older | 16 or older |
| Verkauf | At least one assigned seller is 18 or older | At least one assigned seller is 18 or older |

Adult M/F and youth mA–mE/wA–wE classes are recognized even with league prefixes; SPF is youth. GE, missing and unsupported classes leave timing eligibility unresolved, so new timing claims are refused. An invalid game date also prevents a new age-restricted claim; today's date is never substituted. This change adds no age limit for Ordnungsdienst, optional Unterstützung, or day blocks, including cake delivery.

Verkauf is a group rule. A younger helper or a legacy helper with an unknown birth date may take the first sale slot. Filling the last free slot requires at least one explicitly assigned seller who reaches 18 by the game date; the other seller has no individual minimum. An adult in the responsible team or candidate list does not provide coverage without an assignment. Every claim, including admin and CLI claims or past-game corrections, is checked against saved data. There is no admin age override. Sale candidates refresh after claims and releases, and stale selections can be refused without overwriting saved occupants.

Authorized releases remain possible, including removal of an ineligible occupant or the sole adult seller. Removing that adult leaves coverage outstanding and limits the final free sale slot to helpers who restore it. Existing assignments remain visible after a birth-date correction or game date/class change. The occupancy bar continues to count physical slots; a separate eligibility indication marks invalid or unresolved staffing. To resolve it, self/admin maintenance can complete or correct the stored date, or an authorized caller can release and replace an ineligible helper with an eligible one. A full bar alone does not mean the game's staffing is complete. Statistics and the existing responsible-team MV follow-up include these deficiencies; no new fallback recipient is introduced when there is no active responsible MV.

Access is derived from database facts on every request:

- **Guest**: public read-only schedule, without roster/contact payloads.
- **Member**: roster without others' contact data, statistics, own profile and own assignments.
- **MV**: member rights plus staffing and active-roster maintenance for each team they manage.
- **Admin**: all administration and the audit trail. Bootstrap the first admin with `manage_db.py grant-admin PERSON_ID`.

Teams are fully automatic: they are derived from the scraped age classes (`ak`, e.g. "BL mD") plus one support team ("Supporter"). A person may belong to no team, one team or several teams, without player/staff role distinctions. Teams cannot be created or edited manually. The responsible team is chosen per game; it is never pre-assigned, since the team playing the game should not be pre-booked for its own support tasks. Membership in the playing team always produces the advisory "spielt selbst" warning; membership outside the responsible and Supporter teams produces an advisory outside-team warning. Both remain selectable.

Schema revisions are managed by pinned Alembic migrations, but application startup never creates or upgrades schema implicitly. `manage_db.py init` is only for an absent or empty database; web and daily startup require the configured database to be at the current revision. Existing databases are upgraded only through the guarded, stopped-services `migrate-schema` workflow below. A database with the older source-key game identity must first use the separate `migrate-game-identity` prerequisite.

Authentication does not make the Flask development server suitable for public use.
Generic Gunicorn and systemd templates are tracked in `release-assets/`;
hostnames, credentials, environment files, and installed host configuration
remain private and outside the release source.

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
5. Start exactly one application process. Normal startup re-establishes WAL and repeats the runtime and foreign-key checks. Verify the schedule, authentication and one non-destructive database query before deliberately re-enabling normal traffic and the daily/cleanup timers.
6. Retain the command-created rollback directory until the restored service and the next validated Dropbox backup have been accepted.

Before deploying these settings to an existing database, stop writers, preserve the current database together with any sidecars, run the full offline suite, and start one process to perform the integrity preflight. For rollback, stop every database user first and restore a validated snapshot or the complete preserved database set; never replace only the main file while WAL users are active.

### Migrate legacy game identity

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

### Initialize or migrate the application schema

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
   is inferred. Revision `0005_cake_delivery_blocks` is the single current head;
   it must be reached through this stopped-service migration before startup.
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
- `persons`: stable internal identity, mutable/non-unique name, contacts, date-only `birth_date` (nullable for unknown legacy dates), account state and admin flag
- `person_teams`: deduplicated many-to-many membership set; pending registrations use the same rows, while account status gates roster use
- `assignments`: one stable role/slot link between a game and person
- `day_blocks`: one preparation, cake-delivery and cleanup block per season/home-game date; nullable cake quantity and delivery time belong only to cake blocks
- `block_assignments`: one volunteer per numbered block position, with at most one position per person in each block
- `auth_tokens`: expiring, single-use registration/login nonces and six-digit e-mail/SMS codes
- `assignment_audit`: append-only assignment-change snapshots that remain readable after deletion

The active MV of a game's responsible team receives an MV notification as long as required tasks of that game are still unassigned or have unresolved/invalid age eligibility. The optional `Unterstützung` task never makes a game incomplete. Day blocks have no responsible team and therefore never create an MV reminder.

Preparation, cake delivery and cleanup are date-level duties rather than duties of an individual game. Active members and MVs may claim or release only their own block slot; admins may manage any active person, including corrections on past dates. A person may hold one task in each block independently of any game assignment on the same date. Preparation helpers receive the special one-week reminder, cleanup and cake helpers an ordinary one-week reminder, and all occupied block slots a day-before reminder. Cake reminders include the saved delivery time and ask each volunteer for one cake. Empty positions send no helper reminder; ordinary reminders prefer e-mail, fall back to SMS and skip contactless helpers. Game-level `Verkauf` helpers always use the normal game reminder flow.

## Cake delivery

The **Kuchenlieferung** card appears after preparation and before the games. It starts
collapsed and shows the saved delivery time, requested quantity and staffing progress.
Expand it with the mouse or keyboard to see **Kuchenlieferung 1**, **Kuchenlieferung 2**,
and so on. Each position is one cake contribution. Guests see assigned names; signed-in
viewers get candidates only when an editable card opens. Members and MVs can manage
only their own cake duty, independently of their preparation, game or cleanup duties.

New and migrated cake blocks have neither time nor quantity. They show administrator
setup is needed and offer no claimable positions. An admin expands the card, enters a
delivery time in `HH:MM` (24-hour format, for example `14:30`) and a nonnegative whole-number quantity, then saves the
settings. For example, `10:00` and `4` provide four positions. Quantity `0` explicitly
means no cakes are requested and exposes no positions; it differs from unknown quantity.
The saved delivery time stays attached to the home-game date when game start times change.

Increasing quantity from two to four keeps positions 1 and 2 and adds two empty
positions. A decrease cannot remove an occupied position: to reduce four to two when
position 4 is occupied, explicitly release position 4 first, then save the reduction.
An occupant of position 2 keeps that position; positions are never compacted or silently
released. Every successful claim or release remains in the assignment audit, including
after quantity changes or removal of a date's last game. Stale settings edits and stale
claims are refused without replacing another saved change; reload the saved state before
retrying.

Cake cards use the same date and assigned-person filters as the other blocks. A person
search matching only a cake volunteer retains that card and date without unrelated
games. A playing-team filter retains the cake card when the date has a matching game;
a responsible-team filter hides it. Clearing filters restores the usual upcoming and
past groups. Cake blocks have no responsible team and never generate MV gap reminders.

Each occupied cake position adds one **Kuchenlieferung** duty to the person's season
statistics. A configured four-position block with two volunteers reports two missing
cake duties. Unconfigured blocks are reported separately as needing setup, without an
invented missing-cake count or a complete staffing indicator. Configured zero has no
missing cake duties.

# Getting started

Clone the repository and open `config_template.json`. Modify the club data and notification texts, then save it as `config.json`. Email, Twilio and Dropbox settings belong in environment variables mapped in `common.PROVIDER_ENV`. Production refuses provider sections in `config.json`; local legacy files remain readable during transition. Install all libraries from `requirements.txt`; this includes the pinned `email-validator` and `phonenumbers` packages used for deterministic, offline-safe contact validation.

Initialize the database and manage persons, games and assignments via the management CLI:

```
python3 manage_db.py init
python3 manage_db.py add-person "Max Mustermann" --birth-date 1990-01-01 --team "Supporter" --email max@example.com
python3 manage_db.py list-persons
python3 manage_db.py set-birth-date 1 1990-01-02
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

Run `contact-preflight` before deploying contact normalization to an existing database; it reports invalid values, canonical changes and collisions without writing anything. Person IDs are shown by `list-persons` and `search-person`; names are display data and may repeat. Existing-person commands (`set-birth-date`, `show-memberships`, `set-memberships`, `assign`, `unassign`, `set-mv`, `grant-admin`) therefore take the person ID. `add-person --birth-date` and `set-birth-date PERSON_ID DATE` use ISO `YYYY-MM-DD` dates and shared calendar/future-date validation. Invalid creation data never inserts a partial person. Correction retains the same identity and all assignments; it reevaluates eligibility rather than deleting appointments. `set-memberships` atomically replaces the complete set and accepts no team names to clear it; CLI creation keeps the convenient single `--team` option. Canonical game numbers are unique within a season: `list-games` shows the internal game ID together with its schedule context, and `assign`, `unassign` and `set-jteam` take that game ID. `--number` accepts either an ordinary decimal number or an SPF pseudo number and only filters the list.

Games are inserted automatically on the first run after scraping. Run the daily job via `./run.sh` (or export `NULIGAHELPER_SECRET` before `python3 main.py`). nuLigaHelper will scrape all available data from the nuLiga Hallenspielplan, sync it into the SQLite database and send all due notifications. A backup of the database is uploaded to Dropbox after every run.
All actions are logged in the file `helper.log`.
Run the daily job through one scheduler only; disable any old scheduler before enabling a replacement.

Placeholder games or duty blocks before the first and after the last game of a day are
deliberately outside this change and remain future work.

## Private deployment

For repeat updates to an existing installation, follow the tracked
[deployment guide](DEPLOYMENT.md). It uses a pinned commit on protected `master`
and the installed release command; first-cutover host acceptance remains a
separate prerequisite.

The operator-triggered release workflow has tracked, non-secret
templates in `release-assets/`. Its deploy command now implements separate
preparation, guarded activation, explicit timer resumption, and conditional
code-only rollback. Activation and compatible rollback wait up to 60 seconds
for Gunicorn's loopback listener and local `/healthz` after the `Type=simple`
web service starts; Caddy and timers remain closed if readiness fails or times
out. These commands are **not yet approved for production**:
the deployer is merged into protected `master`, but host preflight and the
first-cutover operator review in the private operations runbook remain
outstanding. Do not activate a candidate or switch `current` from these
templates before that acceptance. `release-assets/README.md` explains the
intended operator command sequence and recovery gates.
The production database remains `/var/lib/nuligahelper/nuliga_helper.db`, and
`/etc/nuligahelper/web.env` remains private and shared by the services.

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
