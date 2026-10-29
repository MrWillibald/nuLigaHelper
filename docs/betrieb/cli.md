# Management CLI (`manage_db.py`)

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
python3 manage_db.py assign 42 Ordner 4
python3 manage_db.py assign 42 Kasse 2
python3 manage_db.py assign 42 Wischer 3
python3 manage_db.py unassign 42 Wischer 3
python3 manage_db.py set-jteam 42 "Damen 1"
```

Use `Ordner` as the CLI task name; `Ordnungsdienst` remains an input alias for
existing scripts and resolves to the same duty.
Game and person arguments are internal IDs. Repeated-role `assign` uses the first
empty numbered position and prints that position. `unassign` refuses ordinary
withdrawal from an offered duty; use the schedule's replacement request and eligible
atomic takeover before lock. It can still retire a retained removed-role assignment
under existing authority before cutoff, without moving its sibling. In the example,
game 42 must be an adult game for Kasse/Wischer; the shown `unassign` is refused
while that Wischer duty is offered. `list-games` shows the single unnumbered Wischer
and both numbered Verkauf occupants.
New youth/unknown Kasse/Wischer assignments and every post-lock ordinary release
are refused.

Run `contact-preflight` before deploying contact normalization to an existing database; it reports invalid values, canonical changes and collisions without writing anything. Person IDs are shown by `list-persons` and `search-person`; names are display data and may repeat. Existing-person commands (`set-birth-date`, `show-memberships`, `set-memberships`, `assign`, `unassign`, `set-mv`, `grant-admin`) therefore take the person ID. `add-person --birth-date` and `set-birth-date PERSON_ID DATE` use ISO `YYYY-MM-DD` dates and shared calendar/future-date validation. Invalid creation data never inserts a partial person. Correction retains the same identity and all assignments; it reevaluates eligibility rather than deleting appointments. `set-memberships` atomically replaces the complete set and accepts no team names to clear it; CLI creation keeps the convenient single `--team` option. Canonical game numbers are unique within a season: `list-games` shows the internal game ID together with its schedule context, and `assign`, `unassign` and `set-jteam` take that game ID. `--number` accepts either an ordinary decimal number or an SPF pseudo number and only filters the list.

Games are inserted automatically on the first run after scraping. Run the daily job via `./run.sh` (or export `NULIGAHELPER_SECRET` before `python3 main.py`). nuLigaHelper will scrape all available data from the nuLiga Hallenspielplan, sync it into the SQLite database and send all due notifications. A backup of the database is uploaded to Dropbox after every run.
All actions are logged in the file `helper.log`.
Run the daily job through one scheduler only; disable any old scheduler before enabling a replacement.

Preparation, cake delivery and cleanup use dated day blocks
([specification](../../openspec/specs/game-day-task-blocks/spec.md));


---
Back to the [README](../../README.MD) · [Operations index](README.md)
