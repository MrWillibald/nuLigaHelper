# nuLigaHelper

A helper class to organize home games with nuLiga Handball.

Host-specific deployment configuration and operator records are maintained privately outside this repository.

# Features

- Read home games from Hallenspielplan website of nuLiga
- Games, persons and task assignments are stored in a SQLite database (`nuliga_helper.db`)
- The database is backed up to a dropbox account after every run
- Game tasks (Zeitnehmer, Sekretär, Verkauf, Ordnungsdienst, Kasse, Reinigung) offer eight required positions for adults and five for youth and unresolved categories
- Every home-game date has three-slot Vorbereitung and Aufräumen blocks, timed 90 minutes before the first and 60 minutes after the last game start; slot numbers are display positions while statistics, reminders and audits use the shared task name, just like Verkauf
- Every home-game date also has a Kuchenlieferung block; admins set its delivery time and requested quantity, with one volunteer position per cake
- nuLigaHelper keeps track of all game shifts
- Notifications are sent to all assigned game helpers one week and one day prior to the game (if the script is run daily), including retained assignments in removed duties
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

- **Heimspielplan**: public, read-only schedule with assigned helper names. Use **Datum** to show one game day or all days, together with **Spielende Mannschaft**, **Verantwortliche Mannschaft** and **Eingeteilte Person**. A person search keeps only games and day-task blocks where that person is assigned; selecting a responsible team hides day-task blocks because they have no responsible team. **Filter löschen** restores the full plan. Every game and day-task block starts compact; cake cards awaiting setup or requesting zero cakes show their status, while configured duties show a staffing bar; compact game cards also show the responsible team or **– offen –**. The triangle beside **Details** points right when closed and down when open, including without JavaScript. Expand a card through **Details** to see the game's offered positions and any retained assignments in removed duties and, when permitted, the editable responsible-team field. Assignment candidates load when an editable card opens; a loading error leaves its dropdowns disabled and offers a retry. Unopened cards carry no candidate lists, so registrations do not multiply the initial page size. Game progress counts eight required positions for adult games and five for youth/unknown games; retained removed duties are release-only and do not affect the bar. Each preparation and cleanup bar counts three slots. Past game days appear muted under **Vergangene Spieltage**; selecting a past date opens that section. Signed-in members can claim and release their own future tasks. MVs can additionally staff their own team on games for which that team is responsible; admins can correct every slot and set the responsible team. Concurrent stale changes are rejected rather than overwriting another person's work.
- **Helfer verwalten**: members see names and complete team memberships but only their own contact data and birth date. Admins create, edit, deactivate/reactivate or delete erroneous entries, replace membership sets, complete missing birth dates, approve user registrations and appoint team MVs. MVs can create people with a birth date and add or remove active people only in rosters they manage; they do not see another existing person's birth date or decide pending registrations.
- **Statistik**: games covered per team, jobs per person (incl. role breakdown) and upcoming games and day blocks with missing assignments, unresolved age eligibility or cake setup still needed.
- **Änderungsprotokoll**: admin-only, append-only history of assignment claims and releases with actor, tier, person, game, role, slot and timestamp.

At **700 CSS pixels wide or less**, JavaScript initially collapses the general
introductions on Heimspielplan, Helfer verwalten and Statistik behind **Hinweise**.
On wider screens they start expanded. Opening a disclosure shows the
complete explanation; validation errors, missing-birth-date follow-up, cake setup
and staffing/age warnings remain beside the content they concern.

The schedule and roster filters start collapsed on mobile behind **Filter**
and expanded on desktop. Applied filters remain readable outside the disclosure,
with **Filter löschen** on the schedule or **Zurücksetzen** on the roster. Reopening
the form retains its selections; opening or closing it does not submit filters.
Team values use their names, and roster summaries include only filters permitted
for the viewer. The existing GET submissions and reset links also work without
JavaScript.

Helper cards keep each person's name, every team and permitted status visible.
Existing authorized profile, membership and record actions sit behind **Bearbeitung**;
those controls start collapsed on mobile, desktop and without JavaScript. Cards with
no permitted action have no empty editing control. A rejected profile edit reopens
the affected card, preserving submitted values and field errors for correction.
Other cards keep their normal
defaults. The administration area still precedes the roster, and disclosure does not
change access to contacts or birth dates.

Statistics sections appear in this order: **Spiele pro Mannschaft**, **Dienste pro
Person**, then **Offene Dienste, Altersanforderungen und Einrichtung**. All three start
collapsed on enhanced mobile pages; desktop opens only the team section initially.
They open independently. Headers count displayed teams, people with duties, and
affected upcoming games/day blocks respectively. One game or block contributes once
to the last count even if it has several vacancies or age/setup issues; the count is
not a number of missing helpers. Zero-count headers remain visible and expand to an
explanation of the empty result.

Disclosure captions identify the content and stay the same when open or closed:
**Hinweise**, **Filter**, **Bearbeitung**, **Details**, **Vergangene Spieltage** and
the statistics section titles. Native disclosures support keyboard and touch.
During a page visit, resizing or rotating preserves deliberate disclosure choices,
unsaved form edits, focused fields
and validation errors. A fresh visit restores the initial defaults. Without JavaScript,
introductions and filters start expanded; authorized helper maintenance starts collapsed
and opens through its native control, except that returned validation errors start open.
The team statistics section starts expanded and the other statistics sections remain
available through their native disclosure controls.

Shared pages offer an upward-arrow **Nach oben** action. With JavaScript it floats at
the lower right after scrolling at least one viewport height and hides near the top;
without JavaScript it remains a footer link. It supports keyboard focus and a touch
target of at least 44×44 CSS pixels, respects safe areas, and yields space to active
feedback, task guidance and dialogs. Activation returns to the top with usable keyboard
focus and leaves form values and filters intact. Reduced-motion settings disable
animated scrolling.

Action results share a bottom-center feedback stack on every page, including assignments
and releases, cake settings, responsible-team/MV changes, helper/profile/membership
maintenance, registration decisions and login/logout. Each bubble includes a German
severity label. Every pill uses a flat tinted background, bold dark text and rounded
corners, without a border or shadow. Confirmation matches the green **Aktiv** badge,
information matches the navy-on-slate team badges, and errors match the pink/red
delete button. Warnings use the corresponding light amber/dark brown treatment.
All bubbles disappear automatically after five seconds of visible reading time;
hovering, keyboard focus, a hidden tab or scrolling a message outside the stack's
visible area pauses that time. Bubbles have no close button. The bounded stack
retains unread results and reveals each newest result
without moving keyboard focus. It wraps on narrow screens, follows the visible viewport
(including the on-screen keyboard), respects safe-area space and reduced-motion settings,
and shares space with task guidance. Without JavaScript, server form feedback stays
readable in the document flow, without timers or close buttons.

Action-required reloads and sign-in navigation carry only that action's plain-text
feedback once through same-tab session storage, expiring after 60 seconds. Its full
reading time starts at the destination. Blocked storage keeps the result on the current
page with an explicit refresh/sign-in link. Login/logout clear obsolete account feedback.
No form values, contact data, birth dates, authentication challenges/codes or security
tokens are copied into this transfer. Server redirects use one-use Flask flashes.

A saved assignment with an advisory says **Dienst gespeichert.** before the warning;
a release says **Dienst freigegeben.** Removing the last contact confirms the profile
save in orange and warns about future login. A replacement confirms release first and
then separately claims the new helper: a refused claim explains that the earlier duty
was released and the replacement was refused. A lost/unreadable response reports an
unconfirmed outcome and asks the user to check the saved state. Failed candidate reads
after a confirmed save retain its success and occupancy/progress; incomplete pickers
stay disabled with local retry/sign-in controls. Field validation, code-entry instructions,
awaiting-approval status and staffing deficiencies remain beside the relevant content.
Contact verification confirms only verification; the persistent status still explains
that administrative approval is pending. Approval and rejection both receive success
confirmation after a completed decision. The pre-deletion warning remains separate.

Registration and login each use one two-step form: request a six-digit code, then enter it on the same page. During registration, E-Mail and SMS are individually optional, but at least one valid contact is required and both may be supplied. The selected route alone receives the verification code, while all supplied contacts are stored; after approval, either stored route can be selected independently for login. Registration additionally requires a name, valid birth date, one or more selected teams and consent to publish the name when assigned. The birth date is validated before requesting a code and preserved through verification. Successful contact verification queues the person and all selected teams as one user-level approval and notifies every active administrator; it does not itself approve the account. When an administrator approves the registration, the new active user receives a welcome notification inviting them to sign in and take open Heimspielplan duties. Both automatic notifications prefer E-Mail and use SMS only when no usable e-mail address is stored. Team membership itself has no separate confirmation. E-mail addresses are syntax-checked and stored in a canonical lowercase form. Phone input combines the selected country calling code with the national number and is stored as E.164 (for example `+491701234567`). Codes are single-use and valid for 15 minutes; sessions expire after one hour of inactivity. Active accounts land on the Heimspielplan after login, while verified registrations awaiting approval see only their status.

New people need a real calendar birth date that is not after `common.effective_today()`, including contactless admin/MV creation and CLI creation. A member can complete or correct their own date; an admin can do so for any person. A known date cannot be cleared. Existing people with an unknown date can still receive unrelated profile repairs. Dates are self-reported; the application does not verify identity documents. Full dates are limited to own/admin person maintenance and are absent from schedules, candidates, statistics, ordinary notifications, diagnostic logs and assignment audits. See the [data inventory](DATA-INVENTORY.md).

## Task descriptions

Every assignment-task label in an expanded game or day-task card has a small **i**
information button, available to guests and signed-in viewers. Hover over it or
focus it with the keyboard to read the German task description. Click or tap to
keep the description open; click or tap again, or press **Escape**, to close it.
The description floats over the underlying content without moving the page or
changing an assignment. It wraps within the viewport on narrow screens; longer
text scrolls inside the help bubble. Numbered positions of the same task share
one description.

Developers maintain the public text in `db.TASK_DESCRIPTIONS`, keyed by semantic
game roles and day-block phases. Add or rename descriptions alongside
`db.ROLE_SLOT_COUNT` or `db.BLOCK_PHASES`; do not add numbered keys or obsolete
roles. Schedule views and dynamically refreshed cake positions reuse this catalog.
`test/test_task_descriptions.py` checks coverage against the actual supported
catalog, repeated-position wording, rendering and public-response privacy.

## Game duty positions

All offered game duties are mandatory. Position numbers distinguish occupants of
one semantic role: **Verkauf 1/2** and **Reinigung 1/2** count as Verkauf or Reinigung
in personal statistics and reminders. There is exactly one independent **Ordnungsdienst**
and one adult **Kasse** role. Kasse additionally takes the ordering/security work
originally intended for a second Ordner. A person still holds at most one task per game.
Per-game Reinigung remains separate from the day-wide Aufräumen block.

| Duty | Positions | Adult games | Youth games and unresolved categories |
| --- | --- | --- | --- |
| Zeitnehmer | 1 | Required | Required |
| Sekretär | 1 | Required | Required |
| Verkauf | 2 | Both required | Both required |
| Ordnungsdienst | 1 | Required | Required |
| Kasse | 1 | Required | Not offered |
| Reinigung | 2 | Both required | Not offered |

Adult M/F and youth mA–mE/wA–wE tokens are recognized with league prefixes; SPF is youth.
GE, missing, conflicting and unsupported classes retain the five-position baseline
and show that classification needs review. The same explicit classifier also determines
timing-duty age eligibility below.

The schedule bar, open-duty statistics and responsible-team MV reminders use the same
offered/required positions. Five baseline occupants produce **5/8** (63%) for an adult
game, with Kasse and both Reinigung places still open; an age-eligible youth roster
shows **5/5** (100%). There are no optional game duties or empty Kasse/Reinigung fields
for youth/unknown games. New claims for these removed duties are refused for all tiers,
including admin, CLI and system callers. Claims recheck the saved game category under
write serialization, so a stale adult classification cannot authorize a youth duty.

A migrated Unterstützung assignment becomes Kasse even on a youth/unknown game.
An adult game changing category can also leave retained Kasse/Reinigung occupants.
These assignments keep their IDs, original positions and history and appear as
**Bestehende Einteilung – nur Freigabe**. Existing member/MV/admin rights and past-date
rules govern release; no replacement candidate or new claim is offered. After release
the field disappears and an append-only audit records it. A retained assignment still
prevents that person from taking a second task in the same game.

All saved assignments, including retained removed duties, count once in personal
statistics and receive ordinary weekly, day-before and game-shift reminders. Each
repeated-role occupant receives one message with the shared duty name. Retained removed
duties do not create vacancies or affect progress/MV follow-up. Age deficiencies remain
separate even when the required-position bar is full. Preparation, day cleanup and cake
blocks keep their existing independent behavior.

## Age eligibility for game duties

Age means completed calendar years on the scheduled game date, including future games. A minimum is met on the birthday: a person turning 14 on 5 September can perform a youth timing duty on 5 September, but not on 4 September. A 29 February birthday advances on 1 March in a non-leap year.

| Duty | Youth games | Adult games |
| --- | --- | --- |
| Zeitnehmer | 14 or older | 18 or older |
| Sekretär | 14 or older | 16 or older |
| Verkauf | At least one assigned seller is 18 or older | At least one assigned seller is 18 or older |

Adult M/F and youth mA–mE/wA–wE classes are recognized even with league prefixes; SPF is youth. GE, missing and unsupported classes leave timing eligibility unresolved, so new timing claims are refused. An invalid game date also prevents a new age-restricted claim; today's date is never substituted. Ordnungsdienst, Kasse, Reinigung and day blocks, including cake delivery, have no additional age minimum.

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
   is inferred. Revision `0006_game_duty_staffing` is the single current head.
   It renames current Unterstützung assignments to Kasse while preserving their
   IDs, game, person and numbered position; older Reinigung rows traverse the
   earlier Reinigung-to-Unterstützung revision first. Historical audit labels
   and snapshots remain unchanged and readable, and new Reinigung places begin
   empty. Unterstützung/Kasse collisions, unexpected new-role rows and invalid
   old positions are refused without merging or discarding them. Youth/unknown
   Kasse assignments become retained release-only assignments, as described
   above. All upgrades require this stopped-service migration before startup.
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

The active MV of a game's responsible team receives an MV notification as long as required positions of that game are still unassigned or have unresolved/invalid age eligibility. Adult games require all eight offered positions; youth games and unresolved classes require the baseline five. Retained removed duties do not affect completeness. Day blocks have no responsible team and therefore never create an MV reminder.

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

Clone the repository and open `config_template.json`. Modify the club and recipient settings, then save it as `config.json`. Notification wording is maintained in `messages.py`. Email, Twilio and Dropbox settings belong in environment variables mapped in `common.PROVIDER_ENV`. Production refuses provider sections in `config.json`; local legacy files remain readable during transition. Install all libraries from `requirements.txt`; this includes the pinned `email-validator` and `phonenumbers` packages used for deterministic, offline-safe contact validation.

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
python3 manage_db.py assign 42 Kasse 2
python3 manage_db.py assign 42 Reinigung 3
python3 manage_db.py assign 42 Reinigung 4
python3 manage_db.py unassign 42 Reinigung 4
python3 manage_db.py set-jteam 42 "Damen 1"
```

Game and person arguments are internal IDs. Repeated-role `assign` uses the first
empty numbered position and prints that position; `unassign` releases that person's
saved position without moving its sibling. In the example, game 42 must be an adult
game for Kasse/Reinigung. `list-games` shows both numbered occupants. Retained removed
assignments can still be released by `unassign`; new youth/unknown Kasse/Reinigung
assignments are refused.

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
variables. Production `config.json` contains club, recipient and operational settings,
not email, Twilio or Dropbox provider sections. Host setup and public launch
require separate operator review.

# Notification wording and migration

`messages.py` is the single editing point for retained notification subjects,
e-mail/SMS bodies and fragments. Its catalog uses explicit keys for day-before
and weekly reminders, ordinary games and Spielfeste, day blocks, account actions,
operator alerts and the dormant newspaper feature. Each entry declares the
required named fields, with their meanings documented beside the catalog.
For example, `{recipient_name}`, `{task_label}`, `{game_date}` and
`{game_start_time}` identify a game duty, while `{block_meeting_time}` identifies
the meeting or delivery time of a day block. Display labels come from domain
data, so numbered positions do not become separate task names in notifications.

Developers edit the relevant catalog entry, preserve the named field contract,
then review synthetic rendered examples and run `test/run_tests.sh`. When a
contract changes, update its declaration, meanings and callers together.
`messages.render(key, **context)` returns `subject`, `email` and `sms` variants;
`messages.render_fragment(key, **context)` renders a reusable fragment.
An explicitly shared variant is declared in the catalog. Ordinary notifications
prefer e-mail and fall back to SMS; authentication uses the explicitly selected
route. Rendering rejects unknown keys, missing or extra fields, and line breaks
in subjects before dispatch. Values stay plain text, including literal braces.
Rendering diagnostics contain identifiers and field names, never values or codes.

For an existing deployment, run this **read-only** preflight from the candidate
release before activation, using the actual private configuration path:

```bash
./venv/bin/python -m message_config --config /path/to/private/config.json
```

The command opens no database or provider connection and sends no notifications.
It reports obsolete keys, their destinations and `prior-default`, `customized`
or `unrecognized` status. Comparison fingerprints cover the last public positional
example and notifier fallback defaults; other historical versions are treated as
customized for review. No message contents or recipient values are printed.
Privately inspect customized or unrecognized entries, map the old positional
values to the documented named fields, and reconcile intended wording into
`messages.py` through normal code review **before deployment**. Referee alerts
now explicitly place age class, game date, start time and notified names; the old
sample's shifted positional values must not be copied into the new contract.

Move `club.texts.mailRefCoordTargets` unchanged to
`club.notifications.referee_targets`; entries retain their `Name` and `Address`
shape. If both are present, the new key takes precedence, including an empty
list; conflicting metadata produces a value-free warning. When only the old key
exists it remains a recipient fallback. Remove `club.texts` once wording and
metadata have been reconciled, and retain a private configuration snapshot for
rollback together with the previous application release. Legacy text presence
alone does not reject startup, but produces key-only deprecation warnings and
never overrides the catalog. Provider/environment settings are unchanged.
No database migration is involved.

The [message inventory](openspec/changes/archive/2026-10-05-centralize-message-templates/message-inventory.md)
records every legacy key, channel and caller, including removed unused early-task
definitions. `send_article` remains callable with centralized newspaper content;
its daily-job call site remains disabled.

Copy review covers spelling and grammar, current task labels, greetings and
sign-offs, readable paragraphs, complete date/time information and a clear next
action. Spielfeste have no invented opponents; block meeting/delivery times are
distinct from match start times. Required welcome wording and the 15-minute code
lifetime must remain intact. Existing club-specific arrival instructions are
retained and need club/domain review before anyone changes their policy;
centralizing text introduces no new duty rules.
