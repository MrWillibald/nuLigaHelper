# AGENTS.md — nuLigaHelper

Context and working rules for coding agents (and humans) working in this repo.
Read this before changing anything; it encodes decisions and pitfalls we already
paid for.

## What this project is

nuLigaHelper organizes home games of the handball department of TuS Raubling.
It runs as a **daily cron job on a Raspberry Pi** and:

1. scrapes all home games from the BHV/nuLiga Hallenspielplan,
2. syncs them into a local **SQLite** database (SQLAlchemy 2.x ORM),
3. reports shifts, missing referees ("§77") and unknown new games,
4. sends notifications by e-mail or SMS (Twilio) to assigned helpers,
5. backs up the database file to Dropbox (backup only – no Excel anymore),
6. serves a **Flask web interface** (local network, port 8080) with passwordless
   authentication, tiered rights, self-service assignments and statistics.
   Production serving, TLS and public exposure remain separate work.

Language: code/comments in English, UI texts and notification templates in German.

## Module map

| File                | Purpose                                                                 |
|---------------------|-------------------------------------------------------------------------|
| `common.py`         | VERSION, `DEBUG_FLAG`/`CHANGE_DAY`/`DEBUG_TODAY`, `effective_today()`, season-year helper, `load_config()` |
| `db.py`             | Models incl. accounts, assignment slots/auth tokens/audit, sync and domain helpers |
| `scraper.py`        | Scrapes nuLiga into plain dicts (keys see `GAME_FIELDS`)                 |
| `notifier.py`       | All mail/SMS notifications; reads assignments from the DB               |
| `webapp.py`         | Flask app factory: auth/registration, tier checks, pages and JSON APIs   |
| `manage_db.py`      | CLI to manage persons/games until the web UI covers everything          |
| `main.py`           | Daily job entry point incl. Dropbox DB backup                           |
| `templates/`, `static/` | Jinja templates + CSS/JS; visual language copied from www.handball-raubling.de |
| `test/`             | Offline test suite (`test/run_tests.sh` or `pytest test/ -v`)           |

## Environment & commands

```bash
python3 -m venv venv && ./venv/bin/pip install -r requirements.txt
./venv/bin/python test/test_webapp.py     # single test file (also: pytest)
test/run_tests.sh                          # whole suite, must stay green
./run.sh                                   # daily job
./run_webapp.sh                            # web UI on http://<ip>:8080
```

- `config.json` holds credentials/texts; it is gitignored. `config_template.json`
  lists every key the code reads. Notification texts are `str.format` templates —
  placeholder **order and count are part of the contract**, don't reorder lightly.
- `NULIGAHELPER_SECRET` is mandatory for the webapp and daily job. Store a
  persistent random value in the environment or the gitignored
  `.nuligahelper_secret` file; rotating it invalidates all sessions.
- `NULIGAHELPER_DB` optionally overrides the SQLite path.
- Two independent debug switches in `common.py`, both `False` in the committed state:
  - `DEBUG_FLAG = True` disables all outbound mail/SMS (`send_Mail`/`send_SMS` return
    early) **and** pins "today" to `DEBUG_TODAY`.
  - `CHANGE_DAY = True` only pins "today" to `DEBUG_TODAY` — messages are still sent,
    so combine it with `DEBUG_FLAG` unless you really want real notifications.
  Everything date-dependent must go through `common.effective_today()`, never
  `datetime.date.today()`, or these switches stop working. Reset both to `False`
  before committing.

## Domain rules (do not break silently)

- **Identity is `Person.id`, never the display name.** Names are mutable and may
  repeat. APIs and CLI mutation commands take internal IDs; person lists and
  pickers show the team beside the name. `get_or_create_person()` is only a
  seeding/test convenience.
- **Access tiers are derived on every request**: guest (no session), member
  (active account), MV (active and referenced by `Team.mv_person_id`) and admin
  (active and `Person.is_admin`). Admin and MV rights form a union. Verified but
  unapproved registrations see only their status; inactive persons cannot log in.
- Guest access is limited to a read-only schedule. The guest response includes
  assigned helper names but no roster payload, person IDs or contact data.
- **Teams are fully automatic**: derived from the scraped age classes (`ak`,
  e.g. "BL mD") plus exactly one seeded support team ("Supporter"). Users cannot
  create/edit/delete teams; the CLI only resolves existing ones.
- **Games are never pre-assigned** to their own age-class team – that team is
  busy playing. The responsible team ("Verantwortlich") is chosen per game
  and may stay empty.
- Dropdown highlighting per game: members of the *playing* team → greyed +
  "spielt selbst" hint; members of other teams → greyed when a responsible team
  is set; responsible/support members unmarked. Everything stays selectable;
  duplicates within a role are rejected server-side.
- **One task per person per game**: a person already assigned to any task of a
  game cannot be assigned to another task of the same game. Such persons are
  removed from the dropdowns of the other tasks of that game (server-rendered
  and kept in sync by `static/app.js`); enforced by the unique constraint
  `(game_id, person_id)` on `assignments`, by `db.assign_person` (raises
  `ValueError`) and by the web/CLI entry points.
- **Assignment writes are per-slot compare-and-swap operations.** Claim expects
  an empty slot; release names the expected occupant. A stale expectation returns
  a conflict and current occupant without overwriting the winner. Use
  `db.claim_slot()` / `db.release_slot()` and the matching JSON endpoints, not a
  read-modify-write of a complete role.
- Every successful assignment mutation writes an append-only audit snapshot.
  Deactivation keeps past assignments, releases and audits future assignments,
  clears MV records and does not restore freed slots on reactivation. Deletion is
  reserved for erroneous records and must preserve readable audit entries.
- Roles: `Zeitnehmer`, `Sekretär`, `Verkauf` (**2 slots**), `Ordnungsdienst`,
  `Reinigung` — see `db.ROLE_SLOT_COUNT`. Any aggregation over roles must iterate
  roles once (e.g. `ROLE_SLOT_COUNT.items()`), not `SLOT_LABELS` in `webapp.py`
  (which repeats Verkauf for the two UI dropdowns).
- **Team MV**: each team can have exactly one Mannschaftsverantwortlicher
  (`Team.mv_person`), who must be a member of that team; assign via web UI
  ("Helfer verwalten" → Mannschaften) or CLI (`manage_db.py set-mv`). The MV of
  a game's responsible team receives the MV notification only while the game
  still has open task slots (`db.missing_slots`). MV is not a per-game
  assignment role anymore.
- **Contact data (mail/phone) must never be rendered on the schedule overview.**
  Channel choice: prefer e-mail, fall back to phone; skip with warning if neither.
- Passwordless e-mail links and SMS codes are signed/single-use and expire after
  15 minutes. Sessions use a one-hour sliding lifetime. State-changing forms and
  JSON endpoints require CSRF tokens; the route guard is default-deny.
- Dates/times are stored exactly as scraped (German `dd.mm.yyyy` strings).
  **Never `ORDER BY` the date column** – use `db.game_sort_key()`.
  Normalizing to real DATE columns is a deliberate future task.
- `§77` inside `score` means "no referee assigned". Transitions (absent → present)
  trigger referee-coordinator notifications; shifts trigger helper notifications.
  Both can fire for the same game in one sync.
- Seasons run July–June (`common.season_year_for`). New scraped games whose `ak`
  != "GE" trigger one admin info-mail (tournament numbers churn weekly).
- **No DB migration layer by intent**: after schema changes delete the `.db`
  file and let the next run recreate it. The owner decides when persistence
  across changes is needed – ask before adding migration machinery.
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

- Tests: offline only, synthetic sample data from `test/helpers.py`; descriptive
  `test_*` names with failure messages that explain the intent. Suite must pass
  before committing.
- Commit messages: short imperative summary line (+ optional body), matching the
  existing history style.
- Run `test/run_tests.sh` after every change; keep `README.MD` in sync when
  features/commands change.
