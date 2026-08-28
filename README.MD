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
- Web interface for viewing the home game schedule and managing helpers (Flask)

# Web interface

Start the web interface with `run_webapp.sh` (or `python3 webapp.py`). It is then reachable in the local network at `http://<device-ip>:8080`.

- **Heimspielplan**: all home games of the current season grouped by day. Each game gets a responsible team ("Verantwortlich") and tasks (Zeitnehmer, Sekretär, Verkauf, Ordnungsdienst, Reinigung) assigned via dropdown directly in the schedule; changes are saved instantly. If a team is assigned to a game, helpers of other teams are greyed out in the dropdowns but can still be chosen (with a hint). A person can hold at most one task per game: once assigned, the person is removed from the dropdowns of the other tasks of that game. Contact data is intentionally not shown on the overview.
- **Helfer verwalten**: add, edit and delete persons incl. e-mail, phone and team. Every person belongs to a team; persons without an active team belong to the support team (created automatically). Each team can have exactly one Mannschaftsverantwortlicher ("MV"), chosen among its members.
- **Statistik**: games covered per team, jobs per person (incl. role breakdown) and a list of upcoming games with missing assignments.

Teams are fully automatic: they are derived from the scraped age classes (`ak`, e.g. "BL mD") plus one support team ("Supporter") for persons without an active team. Teams cannot be created or edited manually. The responsible team is chosen per game; it is never pre-assigned, since the team playing the game should not be pre-booked for its own support tasks. In the dropdowns, members of the responsible team and the support team appear normally; members of other teams and of the playing team itself are greyed out but can still be selected (with a hint).

There is no database migration: after schema changes simply delete `nuliga_helper.db` (or your configured database file) and let the next run recreate it.

The interface is currently open (no authentication) and intended for local network use. User authentication and user rights are planned for a future step.

Optional environment variables: `NULIGAHELPER_DB` (path to the SQLite file) and `NULIGAHELPER_SECRET` (Flask secret key).

# Testing

The `test/` folder contains automated tests for the database, notification and web layers. They run completely offline against throwaway databases:

```bash
test/run_tests.sh
```

or individually via `python3 test/test_db.py` etc., or with pytest (`pytest test/ -v`). See `test/README.md` for details.

# Database schema

- `games`: one row per scraped home game (number, date, time, hall, age class, teams, score)
- `teams`: derived automatically from the scraped age classes plus one support team; each team optionally references its MV (Mannschaftsverantwortlicher)
- `persons`: name plus e-mail address and/or phone number and team, stored once per person
- `assignments`: link between a game and a person including the role (Zeitnehmer, Sekretär, Verkauf, Ordnungsdienst, Reinigung)

The MV of a game's responsible team receives an MV notification as long as tasks of that game are still unassigned.

The database is prepared to be used by a future web interface.

# Getting started

Clone the repository and open `config_template.json`. Modify all entries according to your needs. Accounts and logins have to be specified, the texts can be modified according to your needs. Save the file as `config.json` in the same folder. Make sure all necessary libraries are installed (see `requirements.txt`).

Initialize the database and manage persons, games and assignments via the management CLI:

```
python3 manage_db.py init
python3 manage_db.py add-person "Max Mustermann" --email max@example.com --phone +4917012345678
python3 manage_db.py list-games
python3 manage_db.py assign 12034 Zeitnehmer "Max Mustermann"
python3 manage_db.py set-jteam 12034 "Damen 1"
```

Games are inserted automatically on the first run after scraping. Then run the daily job via `run.sh` or `python3 main.py`. nuLigaHelper will scrape all available data from the nuLiga Hallenspielplan, sync it into the SQLite database and send all due notifications. A backup of the database is uploaded to Dropbox after every run.
All actions are logged in the file `helper.log`.
This script is good to run on a raspberry pi as scheduled job once a day.
If you need any assistance feel free to contact me.
