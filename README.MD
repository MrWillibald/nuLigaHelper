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

# Database schema

- `games`: one row per scraped home game (number, date, time, hall, teams, score, judge team)
- `persons`: name plus e-mail address and/or phone number, stored once per person
- `assignments`: link between a game and a person including the role (MV, Zeitnehmer, Sekretär, Verkauf, Ordnungsdienst, Reinigung)

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
