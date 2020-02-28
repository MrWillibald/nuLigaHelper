# nuLigaHelper

A helper class to organize home games with nuLiga Handball.

# Features

- Read home games from Hallenspielplan website of nuLiga
- Create a *.xlsx file with home game plan
- Game judges including notification addresses can be added to the game plan
- Additional tasks can be assigned to persons including notification addresses
- Home game plan is stored and updated on a dropbox account
- nuLigaHelper keeps track of all game shifts
- Notifications are sent to all game judges one day prior to the game (if the script is run daily)
- Notification articles can be sent to newspapers one week before the game
- Notifications on missing referees are sent to specified persons
- Notifications can be sent via Email or SMS (using Twilio account)

# Getting started

Clone the repository and open `config_template.json`. Modify all entries according to your needs. Accounts and logins have to be specified, the texts can be modified according to your needs. Save the file as `config.json` in the same folder. Make sure all necessary libraries are installed. Then run the script via `python3 nuLigaHelperclass.py`. nuLigaHelper will then scrape all available data from the nuLiga Hallenspielplan and save the game plan to your Dropbox. All actions are logged in the file `helper.log`.
This script is good to run on a raspberry pi as scheduled job once a day.
If you need any assistance feel free to contact me.