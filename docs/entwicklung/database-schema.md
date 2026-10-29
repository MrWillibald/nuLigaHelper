# Database schema

Tables of the SQLite database (SQLAlchemy models in `db.py`). Dates and times are stored as scraped German `dd.mm.yyyy` strings; sort with `db.game_sort_key()`.

- `games`: one task-relevant plan entry identified by `(season_year, game_nr)`; ordinary
  numbers are decimal text, while SPF matches on one date for one full age group collapse
  into a pseudo number such as `SPF:2026-11-28:spf mini`
- `teams`: derived automatically from the scraped age classes plus one support team; each team optionally references its MV (Mannschaftsverantwortlicher)
- `persons`: stable internal identity, mutable/non-unique name, contacts, date-only `birth_date` (nullable for unknown legacy dates), account state and admin flag
- `person_teams`: deduplicated many-to-many membership set; pending registrations use the same rows, while account status gates roster use
- `assignments`: one stable role/slot link between a game and person, with replacement-request token and saved lock state
- `day_blocks`: one preparation, cake-delivery and cleanup block per season/home-game date; nullable cake quantity and delivery time belong only to cake blocks
- `block_assignments`: one volunteer per numbered block position, with at most one position per person in each block, replacement-request token and saved lock state
- `schedule_visibility_setting`: constrained singleton with the optional global first-hidden date, initialized unrestricted
- `duty_cutoffs`: durable season/date cutoff markers, including positions that were empty at 09:00 Europe/Berlin on the day before the duty
- `auth_tokens`: expiring, single-use registration/login nonces and six-digit e-mail/SMS codes
- `assignment_audit`: append-only assignment-change snapshots that remain readable after deletion

---
Back to the [README](../../README.MD)
