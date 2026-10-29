# AGENTS.md — nuLigaHelper

Context and working rules for coding agents (and humans) working in this repo.
Read this before changing anything; it encodes decisions and pitfalls we already
paid for.

## What this project is

nuLigaHelper organizes home games of the handball department of TuS Raubling.
It runs on a Linux host (e.g. a Raspberry Pi) as a **daily scheduled job** and a web service:

1. scrapes all home games from the BHV/nuLiga Hallenspielplan,
2. syncs them into a local **SQLite** database (SQLAlchemy 2.x ORM),
3. reports shifts, missing referees ("§77") and unknown new games,
4. sends notifications by e-mail or SMS (Twilio) to assigned helpers,
5. backs up the database file to Dropbox (backup only – no Excel anymore),
6. serves a **Flask web interface** with passwordless authentication, tiered
   rights, self-service assignments and statistics. Production runs publicly
   behind Caddy (TLS) → Gunicorn → systemd; the first cutover has been accepted
   (see `DEPLOYMENT.md`). `run_webapp.sh` (port 8080) is for local development only.

Language: code/comments in English, UI texts and notification templates in German.

## Module map

Domain and data:

| File | Purpose |
|---|---|
| `db.py` | Domain layer: models (teams, persons, games, assignments, day blocks, audit, auth tokens), SQLite engine/startup checks, sync, duty state machine and all mutation helpers |
| `age_eligibility.py` | Calendar and game-category age rules without exposing private dates |
| `contact_validation.py` | Validation and canonicalization of e-mail and phone contacts |
| `task_labels.py` | Display labels for stored role identifiers (kept separate from `db` to avoid an import cycle) |
| `schema_migrations.py`, `migrations/`, `alembic.ini` | Read-only schema inspection and guarded Alembic upgrades |

Daily job and notifications:

| File | Purpose |
|---|---|
| `main.py` | Daily job entry point: scrape, sync, notify, back up |
| `daily_lock.py` | Per-database process lock against overlapping daily runs |
| `scraper.py` | Scrapes nuLiga into plain dicts (keys see `GAME_FIELDS`) |
| `notifier.py` | All mail/SMS notifications; reads assignments from the DB |
| `messages.py` | Authoritative German notification catalog with named fields and validated rendering |
| `message_config.py` | Read-only legacy-message preflight and referee-recipient compatibility |
| `common.py` | VERSION, `effective_today()` and the `NULIGAHELPER_TODAY`/`NULIGAHELPER_DRY_RUN` overrides, season-year helper, `load_config()` |

Web interface:

| File | Purpose |
|---|---|
| `webapp.py` | Flask app factory: production boundary, auth/registration, tier checks, pages and JSON APIs |
| `auth_abuse.py` | Persistent, privacy-safe authentication abuse controls |
| `templates/`, `static/` | Jinja templates and CSS/JS; visual language copied from www.handball-raubling.de |

Operations:

| File | Purpose |
|---|---|
| `manage_db.py` | Operator CLI for initialization, migration, restore, persons, memberships and assignments |
| `backup.py` | Online SQLite snapshots, Dropbox publication/retention and offline restore |
| `production.py` | Validation of deployment data and non-sensitive operation state |
| `production_logging.py` | Production journal formatting without raw payloads |
| `operations.py` | Production preflight, launch check, monitoring and alerts |
| `monitor_watchdog.py` | Independent root cron watchdog for the monitor |
| `runtime_check.py` | systemd `ExecStartPre` configuration check |
| `privacy.py` | Bounded retention cleanup with aggregate-only reporting |
| `release-assets/` | Reviewed systemd, Caddy, Gunicorn and deploy-command templates |

Other: `test/` offline test suite, `docs/` manuals and operations/developer
documentation, `openspec/` specifications and changes.

## Architecture

Single-process Python, no framework beyond Flask and the SQLAlchemy 2.x ORM over one
SQLite file. Two independent entry points share `db.py`:

- **`main.py`** – daily job (systemd timer in production): `scraper.fetch_home_games()`
  returns plain dicts (`scraper.GAME_FIELDS`) → `db.sync_games()` merges them and returns
  `SyncEvents` (`new_games`, `shifts`, `referee_alerts`, `removed_games`) → `Notifier`
  turns those events plus date-window checks into mails/SMS → a validated snapshot is
  published to Dropbox. All timing goes through `common.effective_today()`.
- **`webapp.py`** – `create_app()` factory with one session per request
  (`get_session()` / teardown) and a default-deny route guard. Pages: schedule (`/`),
  statistics (`/statistik`), persons (`/personen`), duty form (`/dienst`), audit
  (`/audit`), login/registration and legal pages. Mutations go through small JSON APIs
  called from `static/app.js`: `/api/assignment/{claim,release,request-replacement,
  cancel-replacement,takeover}`, the same set under `/api/block-assignment/`,
  `/api/games/<id>/team`, `/api/teams/<id>/mv`, `/api/blocks/<id>/cake-settings` and the
  candidate reads `/api/games/<id>/candidates` and `/api/blocks/<id>/candidates`.
  `build_schedule()` pre-computes per-game view dicts (including greyed-out options and
  their reasons) so templates stay simple.

Business rules belong in `db.py`; `webapp.py` and `manage_db.py` are thin callers and
the tests target `db.py` directly. `notifier.py` reads assignments from the DB and
renders German messages from `messages.py`; `Notifier._dispatch()` centralizes the
channel choice and its return counts are asserted by `test/test_notifier.py`.

No linter or formatter is configured; match the surrounding style (module banner
comments, English code/comments, German UI and notification strings). `config.json`
(gitignored) must exist for `main.py`/`notifier.py`; tests never read it or
`nuliga_helper.db`.

`openspec/` holds the behavior specifications (`openspec/specs/`) and planned or
archived changes (`openspec/changes/`). Use it for behavior changes and larger
refactorings.

## Environment & commands

```bash
python3 -m venv venv && ./venv/bin/pip install -r requirements.txt   # pinned prod set + test tools
./venv/bin/python test/test_webapp.py     # single test file (also: pytest)
./venv/bin/python -m pytest test/test_db.py -v -k sync   # single test / pattern
test/run_tests.sh                          # whole suite, must stay green
./run.sh                                   # daily job
./run_webapp.sh                            # web UI on http://<ip>:8080
./venv/bin/python manage_db.py --help      # operator CLI (see docs/betrieb/cli.md)
```

- `config.json` holds club, recipient and operational settings; it is gitignored. Email, Twilio and Dropbox
  settings come from environment variables mapped in `common.PROVIDER_ENV`.
  Production rejects provider sections in JSON. All retained notification wording
  lives in `messages.py`, with descriptive named placeholders and documented meanings.
  Use its validated renderer; change declared fields and callers together, never
  return to positional formatting or caller-local text. Legacy `club.texts` is
  deprecated and cannot override wording. Run `python -m message_config --config
  CONFIG_PATH` before deployment to reconcile customized legacy templates privately.
  Referee recipients live in `club.notifications.referee_targets`; the old
  `club.texts.mailRefCoordTargets` is only a fallback when the new key is absent.
- `NULIGAHELPER_SECRET` is mandatory for the webapp and daily job. Store a
  persistent random value in the environment or the gitignored
  `.nuligahelper_secret` file; rotating it invalidates all sessions.
- `NULIGAHELPER_DB` optionally overrides the SQLite path.
- Two independent local development overrides, read from the environment on every
  use (nothing to reset before committing):
  - `NULIGAHELPER_TODAY=YYYY-MM-DD` pins `common.effective_today()`; messages are still
    sent.
  - `NULIGAHELPER_DRY_RUN=1` makes `send_Mail`/`send_SMS` return early; dispatch counts
    are unchanged and the date is not pinned. Set both for a dry run on a fixed day.
  Invalid values stop with a German error naming the variable. With
  `NULIGAHELPER_ENV=production`, `common.refuse_overrides_in_production()` refuses either
  variable in web startup, the daily job, `runtime_check.py` and the operations
  preflight/launch check. Everything date-dependent must go through
  `common.effective_today()`, never `datetime.date.today()`, or the date override stops
  working; check `common.dry_run()` at send time, never at import time.

## Domain rules (do not break silently)

- **Identity is `Person.id`, never the display name.** Names are mutable and may
  repeat. APIs and CLI mutation commands take internal IDs; person lists and
  pickers show every team beside the name. `get_or_create_person()` is only a
  seeding/test convenience.
- **Birth dates are date-only values.** `Person.birth_date` is nullable only for
  honestly unknown legacy data; registration, admin/MV creation and CLI creation
  require a validated date no later than `common.effective_today()`. Self/admin
  maintenance may complete or correct it but cannot clear a known date. Unknown
  legacy dates do not block unrelated maintenance. Never infer dates from teams,
  names, account state or assignments, and never persist a derived age.
- **Access tiers are derived on every request**: guest (no session), member
  (active account), MV (active and referenced by `Team.mv_person_id`) and admin
  (active and `Person.is_admin`). Admin and MV rights form a union. Verified but
  unapproved registrations see only their status; inactive persons cannot log in.
- Guest access is limited to a read-only schedule. The guest response includes
  assigned helper names but no roster payload, person IDs or contact data.
- **Teams are fully automatic**: derived from the scraped age classes (`ak`,
  e.g. "BL mD") plus exactly one seeded support team ("Supporter"). Users cannot
  create/edit/delete teams; the CLI only resolves existing ones.
- **Membership is a set** through `person_teams`: a person may belong to zero,
  one or several teams, with no player/staff distinction. Pending registrations
  store their selected teams in the same association, while account status gates
  roster use. Only admins approve/reject registrations. Admins replace complete
  sets; an MV may add/remove active people only for a team they manage and may
  not remove their own qualifying membership.
- **Scraped game identity is `(season_year, game_nr)`.** Ordinary game numbers are
  canonical decimal text; nuLiga meeting links and descriptive fields never participate
  in identity. SPF rows are the sole reuse exception: the scraper collapses every date
  plus full normalized SPF age group into one task-relevant game with a textual pseudo
  number such as `SPF:2026-11-28:spf mini`, earliest start time and no invented matchup.
  A moved Spielfest intentionally has a new identity. Internal mutations and events use
  `Game.id`; duplicate ordinary numbers and inconsistent SPF groups abort before sync.
- **Games are never pre-assigned** to their own age-class team – that team is
  busy playing. The responsible team ("Verantwortlich") is chosen per game
  and may stay empty.
- Dropdown highlighting per game uses the complete membership set: membership
  in the *playing* team → greyed + "spielt selbst" hint; otherwise people with
  neither responsible nor Supporter membership → greyed when a responsible team
  is set. Category precedence is playing, responsible, Supporter, other (rendered
  in reverse suitability order: responsible, Supporter, other, playing). Membership
  warnings remain advisory; selectable candidates must still pass age eligibility.
  Duplicates within a role are rejected server-side.
- **One task per person per game**: a person already assigned to any task of a
  game cannot be assigned to another task of the same game. Such persons are
  removed from the dropdowns of the other tasks of that game (server-rendered
  and kept in sync by `static/app.js`); enforced by the unique constraint
  `(game_id, person_id)` on `assignments`, by `db.assign_person` (raises
  `ValueError`) and by the web/CLI entry points.
- **Assignments are binding; ordinary withdrawal never creates a vacancy.** Before
  any UI claim or takeover, show German `Bestätigen` / `Abbrechen` confirmation
  explaining continued responsibility and the day-before lock. Selection, opening,
  Cancel, Escape and dismissal perform no assignment/request write, successful audit
  or notification. Only explicit confirmation attempts the current-state-validated
  write; preserve keyboard/focus access and the non-JavaScript confirmation path.
  Never save first and compensate with a release on cancellation.
- **Assignment writes are per-slot compare-and-swap operations.** Claim expects
  an empty slot. Before lock, `db.request_replacement()` or
  `db.request_block_replacement()` marks an offered occupied duty `Ersatz gesucht`
  while its helper remains assigned and responsible.
  `db.cancel_replacement()` / `db.cancel_block_replacement()` let the current active
  helper end their own open search before cutoff without releasing the duty;
  existing admin/MV scope is retained. Closure compares the expected occupant and
  request token under serialization, records one actual-actor `request_close`
  transition and leaves responsibility/occupancy unchanged. Reopening uses a fresh
  token; stale closure and takeover requests cannot affect the new generation.
  `db.takeover_slot()` /
  `db.takeover_block_slot()` atomically transfer to an eligible replacement, comparing
  both expected occupant and replacement token; a failed or stale transfer keeps the
  saved responsibility. Enforce current tier scope, offeredness/capacity, one duty per
  container and resulting-group Verkauf eligibility under serialization and after
  retries. Use these helpers and matching JSON endpoints, never browser release-then-
  claim or whole-role read-modify-write. `db.release_slot()` is only authorized
  pre-cutoff retirement of a retained removed game duty; block duties have no ordinary
  release operation.
- **Occupied duties lock at 09:00 Europe/Berlin on the day before their date.**
  The fixed cutoff applies even if cron is missed or no reminder can be delivered.
  Use the shared aware `db.duty_now()` clock (effective date from
  `common.effective_today()`), recheck inside every writer after serialization and
  retries, and retain any saved lock after game date/category changes. Due lock
  persistence records the season/date marker, closes open requests and audits first
  locks idempotently; `db.lock_due_duties()` runs as an explicit preflight transaction
  before schedule/API mutations and reminder dispatch, never a hidden commit of
  unrelated pending work. Read-only pages and candidate requests derive due locked
  state from the same clock without writing; opening or cancelling confirmation
  must remain free of mutations. `db.lock_day_before_batch()` must not advance 09:00 on an
  early reminder call. Missing contacts, provider failure, restart and retry do not
  reopen commitments; weekly reminders do not lock. Empty offered duties stay
  claimable within existing rights and lock immediately when claimed after cutoff.
  No tier, including admin or CLI, may release, request replacement for or change a
  locked occupied duty. There is no 24-hour window, automatic return, periodic
  rollback or replacement broadcast. A saved cutoff is not proof of delivery.

- **Day blocks are dated, team-independent containers.** Preparation and cleanup
  keep three positions each; cake delivery uses `db.block_capacity()` and starts
  with quantity/time unset. Only admins configure cakes through
  `db.configure_cake_block()`, with expected saved settings. Zero is an explicit
  request for no cakes. Configuration and assignment writes share serialization;
  reductions must refuse occupied removed positions without force release or
  compaction; ordinary replacement leaves the position occupied. One person may hold
  one position in each block independently of game duties. New block audit snapshots preserve season, date
  and original numbered position; existing snapshots stay unchanged.
- **Game age eligibility uses the scheduled game date** and completed calendar
  years. Youth Zeitnehmer/Sekretär require 14; adult Zeitnehmer and Sekretär
  require 16. Ordner requires a known birth date and age 18 for every
  category, including unresolved categories. Adult M/F, youth mA–mE/wA–wE and
  SPF use the shared explicit classifier; GE, missing/unsupported classes or an
  invalid game date cannot authorize a new timing claim. An invalid date also
  refuses Ordner claims. Thresholds are inclusive on the birthday;
  29 February advances on 1 March in a non-leap year.
- **Verkauf at adult and unresolved-category games requires collective adult coverage.**
  Recognized youth/SPF Verkauf has no individual minimum, birth-date requirement
  or collective adult coverage; its age check needs no valid game date, while general
  assignment/date and cutoff guards remain unchanged. GE, missing and unsupported
  classes receive no youth exemption. For non-youth games, a first younger/unknown-date
  seller is permitted while another sale slot remains free. A claim filling the
  last slot must leave an explicitly assigned seller at least 18 on the game
  date. The other seller has no individual minimum. Invalid game dates refuse
  non-youth Verkauf claims even while staffing is incomplete. Revalidate saved game,
  person and post-claim group state under write serialization and after retries;
  every tier and CLI/system callers obey the rule, with no admin override.
  A replacement request retains the sole adult seller and their coverage; atomic
  takeover must validate the resulting non-youth group without counting the outgoing
  helper; youth/SPF takeovers may remove the sole adult seller. Deficiencies never
  authorize ordinary release or a lock override.
- **Vacancies and age deficiencies are separate.** Existing assignments survive
  migration, date corrections and game changes; reevaluate current/future staffing
  without rewriting history or audits. Occupancy counts remain physical, while
  schedule, statistics and existing MV follow-up include unresolved/invalid age
  eligibility. Current occupants remain visible and may seek an eligible replacement
  before cutoff; a locked deficient assignment stays assigned. Refresh affected
  candidates after saved claims/takeovers. This adds no age rules for day blocks.
- Every successful assignment mutation and first request/lock transition writes
  append-only audit evidence. An atomic takeover records the old and new responsible
  helpers together; repeated lock enforcement creates no duplicate transition.
  **Destructive maintenance may not cancel binding commitments.** Deactivation is
  refused atomically while a current/future game or block assignment would be
  removed; retain account status, MV appointments, assignments and request state.
  Past-only deactivation remains possible and preserves past duties/history.
  Erroneous-record deletion refuses any assigned commitment, including past ones;
  deleting an empty mistaken record preserves historical audit snapshots. Removal of
  an occupied game/day/block, including the last game of a date, is refused without
  partial destructive synchronization or successful-removal audits/notifications.
  These restrictions apply before and after cutoff and have no administrator/system
  override. Successful deactivation without blocking duties clears MV appointments;
  reactivation restores no historically released assignment.

- Current task names use `db.role_label()` / `db.position_label()`. **Ordner** retains
  the stable assignment/API key `Ordnungsdienst`; never use its display label in
  assignment lookups or mutation payloads. New audit entries capture the current
  label in their own `role` field; render historical audit labels literally and
  never rewrite them through the current label mapping.
- Roles: `Zeitnehmer`, `Sekretär`, `Verkauf` (**2 slots**), `Ordner`
  (**1 slot**), `Kasse`, `Wischer` (**1 slot**) — see `db.ROLE_SLOT_COUNT`.
  Adults offer all seven positions; youth/SPF and unknown classes offer the
  baseline five without Kasse/Reinigung. All offered positions are required; use
  `db.offered_positions()` / `db.required_positions()` for per-game calculations.
  Retained assignments outside that set stay visible, count in personal
  statistics/reminders and prevent a second duty. Existing authorized release-only
  retirement is available before the actual cutoff; afterward ordinary release is
  locked. They cannot accept claims, replacement requests or takeovers and never
  affect required progress. Any aggregation over roles must iterate
  roles once (e.g. `ROLE_SLOT_COUNT.items()`), not `SLOT_LABELS` in `webapp.py`
  (which repeats Verkauf for the two UI dropdowns).
- **Team MV**: each team can have exactly one Mannschaftsverantwortlicher
  (`Team.mv_person`), who must be an active member of that team; assign via web UI
  ("Helfer verwalten" → Mannschaften) or CLI (`manage_db.py set-mv`). The MV of
  a game's responsible team receives the MV notification only while the game
  still has open task slots or age deficiencies (`db.staffing_status`). MV is not a per-game
  assignment role anymore.
- **Contact data (mail/phone) must never be rendered on the schedule overview.**
  Validate and canonicalize every contact at authentication and person-write
  boundaries: e-mail uses the shared offline-safe normalizer; phone uses the
  selected calling code and is stored as E.164. Run `manage_db.py contact-preflight`
  before normalizing existing data; collisions are reported, never auto-merged.
- **Full birth dates are private to self/admin person maintenance.** MV creation
  accepts the new date but grants no continuing access to another person's date.
  Schedule/candidate responses, statistics, ordinary notifications, diagnostics
  and assignment audits contain no full birth dates or exact personal ages;
  return necessary thresholds or derived eligibility reasons instead. Database
  backups contain dates and retain their existing protected access/retention.
- Ordinary notifications prefer e-mail, fall back to phone and skip with warning
  if neither exists. Authentication is different: the explicitly selected E-Mail
  or SMS route receives a six-digit code even when both contacts exist.
- Passwordless e-mail and SMS challenges are purpose-bound, signed/single-use and
  expire after 15 minutes. Unknown or ineligible contacts receive a same-shaped
  dummy challenge without a message or session. Sessions use a one-hour sliding
  lifetime. State-changing forms and JSON endpoints require CSRF tokens; the route
  guard is default-deny.
- Dates/times are stored exactly as scraped (German `dd.mm.yyyy` strings).
  **Never `ORDER BY` the date column** – use `db.game_sort_key()`.
  Normalizing to real DATE columns is a deliberate future task.
- `§77` inside `score` means "no referee assigned". Transitions (absent → present)
  trigger referee-coordinator notifications; shifts trigger helper notifications.
  Both can fire for the same game in one sync.
- Seasons run July–June (`common.season_year_for`). New scraped games whose `ak`
  != "GE" trigger one admin info-mail (tournament numbers churn weekly).
- **Schema revisions use Alembic and never run implicitly.** `manage_db.py init`
  accepts only an absent/empty target; web and daily startup require the current
  head. Operators stop all database users and run `manage_db.py migrate-schema
  --confirm-stopped`, which fingerprints, snapshots, upgrades and verifies. A
  legacy `games.source_key` database is refused unchanged; its one-time converter
  was removed after all databases were migrated. Unknown/near-miss schemas fail closed; rollback restores the
  printed snapshot rather than downgrading membership in place.
  The birth-date revision joins the single existing head, preserves unknown
  legacy dates and all person relationships, assignments and audit snapshots,
  and verifies retained data after upgrading. Backfill through self/admin
  maintenance or `set-birth-date PERSON_ID YYYY-MM-DD`; rollback restores the
  validated pre-migration snapshot and previous application with writers stopped.
  The sticky-duty revision preserves assignments and old audit snapshots while
  adding replacement tokens, saved locks and a dated cutoff marker. It must not
  fabricate prior confirmation, replacement requests or historical reminder batches.
- Revision `0008_single_wischer` is the one-time, explicitly requested exception to
  commitment preservation: current Reinigung slot 0 becomes Wischer with its ID,
  helper, lock and replacement state intact; every current Reinigung slot 1 is deleted
  across all seasons, dates and categories, including past, locked and open-request
  duties. Never promote its helper into slot 0. Deleted duties cease counting in
  personal statistics (including past counts) and reminders. Append exactly one system
  `migration_remove` audit per deletion, preserving every old audit and unrelated row.
  Pre-day-block original Reinigung still migrates to Kasse. Use only the stopped-writers,
  backup-first guarded migration; recovery restores the printed snapshot and prior
  application, including deleted assignments. This grants no runtime release, destructive
  maintenance or cutoff bypass. Adult games now require seven positions.
- The newspaper-article feature lives in `Notifier.send_article` but its call
  site in `main.py` is commented out.

## Pitfalls already solved (don't reintroduce)

- **pandas 3**: `read_html` rejects raw bytes (decode to `str` first);
  indexing with a *tuple* is one key (use lists for multi-column selection);
  string columns are NA-aware — clean values via the `_clean` pattern in
  `scraper.py` so empty cells become `""`, never `NA`/"nan".
- **SQLAlchemy collections**: `session.delete(obj)` leaves stale entries inside
  loaded relationships. Mutate assignments through the slot helpers (which use
  collection removal and flush), otherwise duplicate
  checks hit ghost rows and inserts silently vanish.
- Dispatch counting: a person without mail *and* phone counts as skipped (return 0),
  every valid contact returns 1 — tests rely on these exact counts.
- Webapp tests build on each other top-to-bottom (a click-through scenario);
  keep them order-tolerant or update the scenario consciously.
- `test/test_auth.py`, `test/test_refusals.py` and `test/test_concurrency.py`
  cover authentication, tier failures and stale claims. Every test database and
  secret must remain synthetic; tests must never read `config.json`.

## Conventions

- Task descriptions shown behind the **i** buttons live in `db.TASK_DESCRIPTIONS`,
  keyed by semantic game roles and day-block phases. Add or rename them together with
  `db.ROLE_SLOT_COUNT` or `db.BLOCK_PHASES`; never add numbered keys or obsolete roles.
  `test/test_task_descriptions.py` checks coverage.
- Tests: offline only, synthetic sample data from `test/helpers.py`; descriptive
  `test_*` names with failure messages that explain the intent. Suite must pass
  before committing.
- Commit messages: short imperative summary line (+ optional body), matching the
  existing history style.
- Run `test/run_tests.sh` after every change; keep `README.MD` and `docs/` in sync when
  features/commands change; behavior belongs in `openspec/specs/`, operator
  procedures in `docs/betrieb/`.
