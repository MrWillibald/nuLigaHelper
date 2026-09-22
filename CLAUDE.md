# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

`AGENTS.md` is the canonical, hand-maintained rulebook for this repo (domain rules,
solved pitfalls, conventions). **Read it before changing anything**; this file only adds
orientation and does not replace it. When behaviour documented there changes, update
`AGENTS.md` and `README.MD` in the same change.

## Commands

```bash
python3 -m venv venv && ./venv/bin/pip install -r requirements.txt

test/run_tests.sh                            # full suite (standalone runner) — must stay green
./venv/bin/python test/test_webapp.py        # one test file
./venv/bin/python -m pytest test/ -v         # pytest variant
./venv/bin/python -m pytest test/test_db.py -v -k sync   # single test / pattern

./run.sh                                     # daily job (scrape + sync + notify + backup)
./run_webapp.sh                              # web UI on http://<ip>:8080
./venv/bin/python manage_db.py --help        # CLI for persons/assignments/MV
```

There is no linter or formatter configured; match the surrounding style (module banner
comments, English code/comments, German UI and notification strings).

`config.json` (gitignored) must exist for `main.py`/`notifier.py`; `config_template.json`
lists every key that is read (`club.{info,email,dropbox,twilio,database,texts}`). Tests
never touch `config.json` or `nuliga_helper.db` — they use `config_template.json` texts
and temp databases from `test/helpers.py`.

## Architecture

Single-process Python, no framework beyond Flask + SQLAlchemy 2.x ORM over one SQLite
file. Two independent entry points share `db.py`:

- **`main.py`** — daily cron job on a Raspberry Pi: `scraper.fetch_home_games()` returns
  plain dicts (`scraper.GAME_FIELDS`) → `db.sync_games()` merges them and returns a
  `SyncEvents` dataclass (`new_games`, `shifts`, `referee_alerts`, `removed_games`) →
  `Notifier` turns those events plus date-window checks into mails/SMS → the `.db` file
  is uploaded to Dropbox. All timing decisions go through `common.effective_today()`, so
  `DEBUG_FLAG`/`CHANGE_DAY` in `common.py` can fake "today" and suppress outbound
  messages.
- **`webapp.py`** — `create_app()` factory, one session per request (`get_session()` /
  teardown). Pages: schedule (`/`), statistics (`/statistik`), persons (`/personen`);
  mutations happen through small JSON APIs (`/api/assignment`, `/api/game/<id>/team`,
  `/api/team/<id>/mv`) called from `static/app.js`, which also keeps dropdown state in
  sync client-side after each save. `build_schedule()` pre-computes per-game view dicts
  (including which person options are greyed out and why) so templates stay dumb.

`db.py` is the whole domain layer: models `Team`/`Person`/`Game`/`Assignment`, role
constants + `ROLE_SLOT_COUNT`, and every mutation helper (`assign_person`,
`set_role_assignments`, `set_team_mv`, `missing_slots`, `game_sort_key`). Business rules
belong here, not in `webapp.py` or `manage_db.py` — both are thin callers, and the tests
target `db.py` directly.

`notifier.py` reads assignments straight from the DB and formats German `str.format`
templates from `config["texts"]`; placeholder order and count are part of the contract.
`_dispatch()` centralizes channel choice (e-mail preferred, phone fallback, skip with a
warning if neither) and its return counts are asserted by `test_notifier.py`.

### Invariants worth knowing before editing

Full list in `AGENTS.md`; these bite most often:

- Teams are derived automatically from scraped age classes (`ak`) plus the seeded
  `Supporter` team — never user-created. A game is never pre-assigned to its own
  age-class team.
- One task per person per game: enforced by the `uq_game_person` constraint,
  `db.assign_person` (raises `ValueError`), and the dropdown filtering in webapp/JS.
- `Verkauf` has 2 slots. Aggregate over `db.ROLE_SLOT_COUNT`, not over
  `webapp.SLOT_LABELS` (which lists Verkauf twice for the UI).
- Dates/times are stored as scraped German `dd.mm.yyyy` strings — never `ORDER BY` the
  date column, sort with `db.game_sort_key()`.
- No migration layer by intent: after a schema change delete the `.db` file. Ask before
  adding migration machinery.
- Contact data must never be rendered on the schedule overview.

## Workflow tooling

`openspec/` plus the `opsx:*` / `openspec-*` skills are set up (currently no specs or
changes tracked). Use them when the user asks for spec-driven change management;
otherwise work directly.
