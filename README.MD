# nuLigaHelper

A helper class to organize home games with nuLiga Handball.

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
The supported production path uses one loopback Gunicorn worker supervised by
systemd behind Caddy-managed HTTPS. Follow the complete
[Raspberry Pi/Linux production guide](deploy/PRODUCTION.md).

- **Heimspielplan**: public, read-only schedule with assigned helper names. Signed-in members can claim and release their own future tasks. MVs can additionally staff their own team on games for which that team is responsible; admins can correct every slot and set the responsible team. Concurrent stale changes are rejected rather than overwriting another person's work.
- **Helfer verwalten**: members see names and teams but only their own contact data. Admins create, edit, deactivate/reactivate or delete erroneous entries and appoint team MVs. MVs and admins approve registrations in their scope.
- **Statistik**: games covered per team, jobs per person (incl. role breakdown) and a list of upcoming games with missing assignments.
- **Änderungsprotokoll**: admin-only, append-only history of assignment claims and releases with actor, tier, person, game, role, slot and timestamp.

Registration and login each use one two-step form: request a six-digit code, then enter it on the same page. During registration, E-Mail and SMS are individually optional, but at least one valid contact is required and both may be supplied. The selected route alone receives the verification code, while all supplied contacts are stored; after approval, either stored route can be selected independently for login. Automatic notifications still prefer E-Mail and use SMS only when no usable e-mail address is stored. Registration additionally requires a name, desired team and consent to publish the name when assigned; verification queues the registration for approval by the requested team's MV (or an admin for Supporter/teams without an MV). E-mail addresses are syntax-checked and stored in a canonical lowercase form. Phone input combines the selected country calling code with the national number and is stored as E.164 (for example `+491701234567`). Codes are single-use and valid for 15 minutes; sessions expire after one hour of inactivity. Active accounts land on the Heimspielplan after login, while verified registrations awaiting approval see only their status.

Access is derived from database facts on every request:

- **Guest**: public read-only schedule, without roster/contact payloads.
- **Member**: roster without others' contact data, statistics, own profile and own assignments.
- **MV**: member rights plus staffing and registration approval for the MV's team.
- **Admin**: all administration and the audit trail. Bootstrap the first admin with `manage_db.py grant-admin PERSON_ID`.

Teams are fully automatic: they are derived from the scraped age classes (`ak`, e.g. "BL mD") plus one support team ("Supporter") for persons without an active team. Teams cannot be created or edited manually. The responsible team is chosen per game; it is never pre-assigned, since the team playing the game should not be pre-booked for its own support tasks. In the dropdowns, members of the responsible team and the support team appear normally; members of other teams and of the playing team itself are greyed out but can still be selected (with a hint).

There is no schema-migration layer. During development, a database may be recreated after a schema change; once real production data exists, preserve it and make schema upgrades only through a separately reviewed migration plan. The SQLite runtime settings below do not change the table schema and do not require recreating the database.

Authentication does not make the Flask development server suitable for public use.
The versioned Gunicorn, systemd, Caddy, and environment examples under `deploy/`
provide the supported hardened production path; site-specific secrets and hostnames
remain outside Git.

## SQLite production profile and backups

All application database connections enable foreign-key enforcement, wait up to five seconds for transient locks and use `synchronous=FULL`. Startup establishes WAL journal mode and refuses to continue if WAL, the connection settings or `PRAGMA foreign_key_check` cannot be verified. Existing foreign-key violations are reported for explicit operator repair or restore; they are never deleted automatically.

The supported deployment uses the database on a local Raspberry Pi filesystem, exactly one web process/worker, and at most one daily-job process. Multi-worker WSGI operation, network-filesystem SQLite and PostgreSQL are outside the currently verified topology. WAL creates temporary `<database>-wal` and `<database>-shm` files while connections are active. Do not copy, move or delete these files while the application is running.

The daily job takes a non-blocking OS lock at `<resolved-database>.daily.lock` before database initialization or external work. A concurrent run for the same database fails immediately with a nonzero exit status; an unlocked file left after a restart is harmless. `run.sh` uses `exec`, so cron receives the Python process status directly.

Backups use SQLite's online backup API and therefore include one complete committed state even when committed pages still reside in WAL. The live `.db` file and its sidecars are never uploaded directly. Every snapshot must pass `PRAGMA quick_check` and `PRAGMA foreign_key_check` before publication. Dropbox receives, in this order:

1. `<database-stem>-YYYY-MM-DD<suffix>` for `common.effective_today()`;
2. the stable configured database basename as the latest backup;
3. retention cleanup of only older names that exactly match that database's dated convention.

`club.dropbox.dated_retention` is optional and defaults to `14`; it must be a positive integer. A same-day rerun replaces that day's recovery point. Retry-safe Dropbox operations recover from transient connection, DNS, timeout, rate-limit and server failures with three bounded retries after 1, 2 and 4 seconds; authentication, bad-input and other permanent failures fail immediately. Snapshot, validation, read, Dropbox client, upload, listing, pagination, deletion and local-cleanup failures are logged concisely by stage and make the daily job exit unsuccessfully. When the source database remains usable, due notification work is still attempted after a backup failure, but notifications are not retried automatically.

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

`NULIGAHELPER_SECRET` is mandatory for both the webapp and daily job. Set it in the environment or keep the raw value in the gitignored `.nuligahelper_secret` file used by both launch scripts. Keep it stable: changing it invalidates every session. `NULIGAHELPER_DB` is an optional SQLite-path override locally; production requires `NULIGAHELPER_ENV=production`, an absolute `NULIGAHELPER_DB`, and exact `NULIGAHELPER_TRUSTED_HOSTS` entries as described in the [production guide](deploy/PRODUCTION.md).

# Testing

The `test/` folder contains offline tests for the database, notifications, authentication, authorization, CLI, concurrent claims and web flow. Every test uses synthetic secrets and throwaway databases:

```bash
test/run_tests.sh
```

or individually via `python3 test/test_db.py` etc., or with pytest (`pytest test/ -v`). See `test/README.md` for details.

# Database schema

- `games`: one row per scraped home game, identified within a season by a stable nuLiga source key; game numbers are display data and may repeat
- `teams`: derived automatically from the scraped age classes plus one support team; each team optionally references its MV (Mannschaftsverantwortlicher)
- `persons`: stable internal identity, mutable/non-unique name, contact/team, account state and admin flag
- `assignments`: one stable role/slot link between a game and person
- `auth_tokens`: expiring, single-use registration/login nonces and six-digit e-mail/SMS codes
- `assignment_audit`: append-only assignment-change snapshots that remain readable after deletion

The MV of a game's responsible team receives an MV notification as long as tasks of that game are still unassigned.

# Getting started

Clone the repository and open `config_template.json`. Modify all entries according to your needs. Accounts and logins have to be specified, the texts can be modified according to your needs. Save the file as `config.json` in the same folder. Install all libraries from `requirements.txt`; this includes the pinned `email-validator` and `phonenumbers` packages used for deterministic, offline-safe contact validation.

Initialize the database and manage persons, games and assignments via the management CLI:

```
python3 manage_db.py init
python3 manage_db.py add-person "Max Mustermann" --team "Supporter" --email max@example.com
python3 manage_db.py list-persons
python3 manage_db.py contact-preflight
python3 manage_db.py search-person "Max"
python3 manage_db.py grant-admin 1
python3 manage_db.py list-games
python3 manage_db.py list-games --number 12034
python3 manage_db.py assign 42 Zeitnehmer 1
python3 manage_db.py set-jteam 42 "Damen 1"
```

Run `contact-preflight` before deploying contact normalization to an existing database; it reports invalid values, canonical changes and collisions without writing anything. Person IDs are shown by `list-persons` and `search-person`; names are display data and may repeat. Existing-person commands (`assign`, `unassign`, `set-mv`, `grant-admin`) therefore take the person ID. Game numbers may also repeat: `list-games` shows the internal game ID together with its schedule context, and `assign`, `unassign` and `set-jteam` take that game ID. `--number` only filters the list.

Games are inserted automatically on the first run after scraping. Run the daily job via `./run.sh` (or export `NULIGAHELPER_SECRET` before `python3 main.py`). nuLigaHelper will scrape all available data from the nuLiga Hallenspielplan, sync it into the SQLite database and send all due notifications. A backup of the database is uploaded to Dropbox after every run.
All actions are logged in the file `helper.log`.
This script is good to run on a raspberry pi as scheduled job once a day.
If you need any assistance feel free to contact me.
