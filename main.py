# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Main entry point: daily scheduled job
#
# - Scrapes home games from nuLiga and syncs them into SQLite DB
# - Reports shifts, missing referees and unknown new games
# - Sends notifications via e-mail/SMS based on DB assignments
# - Backs up the database file to Dropbox
# ---------------------------------------------------------------

import datetime
import logging
import os

import dropbox

import common
import db
from notifier import Notifier
from scraper import fetch_home_games


def backup_to_dropbox(token: str, folder: str, local_path: str):
    """Upload the SQLite database file as backup to Dropbox."""
    try:
        dbc = dropbox.Dropbox(token)
        remote_path = f"/{folder}/{os.path.basename(local_path)}"
        with open(local_path, "rb") as f:
            dbc.files_upload(
                f.read(), remote_path, mode=dropbox.files.WriteMode.overwrite
            )
        logging.info("Database backup successfully uploaded to Dropbox")
    except dropbox.exceptions.ApiError:
        logging.warning("Database backup to Dropbox failed")


def main():
    # Initialize logger
    logging.basicConfig(
        format="%(asctime)s - %(levelname)s - %(message)s",
        filename="helper.log",
        level=logging.DEBUG,
    )
    logging.getLogger().addHandler(logging.StreamHandler())
    logging.getLogger("twilio.http_client").setLevel(logging.WARNING)
    logging.getLogger("requests").setLevel(logging.WARNING)
    logging.getLogger("urllib3").setLevel(logging.WARNING)

    logging.info("#################################################")
    logging.info("nuLiga Helper start, version " + common.VERSION)
    logging.info("-------------------------------------------------")

    config = common.load_config()
    club_cfg = config["club"]

    today = common.effective_today()
    season_year = common.season_year_for(today)

    # Initialize database
    db_path = club_cfg.get("database", {}).get("path", db.DEFAULT_DB_PATH)
    engine = db.make_engine(db_path)
    db.init_db(engine)

    with db.Session(engine) as session:
        # Scrape current home games and merge into database
        scraped = fetch_home_games(club_cfg["info"], season_year)
        events = db.sync_games(session, scraped, season_year)

        # Backup database to Dropbox
        backup_to_dropbox(
            club_cfg["dropbox"]["dropbox_token"],
            club_cfg["dropbox"]["dropbox_folder"],
            db.resolve_db_path(db_path),
        )
        logging.info("-------------------------------------------------")

        notifier = Notifier(club_cfg, session, season_year)

        # Report shifted games
        cnt = notifier.notify_shifts(events.shifts)
        logging.info(f"Number of sent shift notifications: {cnt}")
        logging.info("-------------------------------------------------")

        # Report newly missing referees detected during sync
        for event in events.referee_alerts:
            notifier.notify_referee_alert(event)
        logging.info(
            f"Referee alerts during sync: {len(events.referee_alerts)}"
        )
        logging.info("-------------------------------------------------")

        # Report unknown new games to admin (tournament games excluded)
        ak_by_nr = {rec["game_nr"]: rec["ak"] for rec in scraped}
        new_nrs = [nr for nr in events.new_games if ak_by_nr.get(nr) != "GE"]
        notifier.notify_new_games(new_nrs)

        """
        # Check if newspaper article has to be sent
        gameDateSa      = today + datetime.timedelta(days=9)
        strGameDateSa   = gameDateSa.strftime("%d.%m.%Y")
        strGameDaySa    = gameDateSa.strftime("%A")
        gameDateSo      = today + datetime.timedelta(days=10)
        strGameDateSo   = gameDateSo.strftime("%d.%m.%Y")
        strGameDaySo    = gameDateSo.strftime("%A")

        # Send newspaper article for Saturday
        if strGameDaySa == "Saturday" and db.get_games_on_date(session, strGameDateSa):
            articleDate = gameDateSa + datetime.timedelta(days=-1)
            cnt         = notifier.send_article(strGameDateSa, "Samstag", articleDate.strftime("%d.%m.%Y"))

        # Send newspaper article for Sunday
        elif strGameDaySo == "Sunday" and db.get_games_on_date(session, strGameDateSo):
            articleDate = gameDateSo + datetime.timedelta(days=-2)
            cnt         = notifier.send_article(strGameDateSo, "Sonntag", articleDate.strftime("%d.%m.%Y"))
        """

        # Check if judge notifications have to be sent
        tomorrow = (today + datetime.timedelta(days=1)).strftime("%d.%m.%Y")
        if db.get_games_on_date(session, tomorrow):
            cnt = notifier.notify_game_day(tomorrow)
            logging.info(f"Number of sent service notifications: {cnt}")
            logging.info("-------------------------------------------------")

            # Check if referee notifications have to be sent
            cnt = notifier.notify_referees_for_date(tomorrow)
            logging.info(f"Number of required home referees: {cnt}")
            logging.info("-------------------------------------------------")

        # Check if early catering notifications have to be sent
        next_week = (today + datetime.timedelta(days=7)).strftime("%d.%m.%Y")
        if db.get_games_on_date(session, next_week):
            cnt = notifier.notify_service_early(next_week)
            cnt += notifier.notify_pre(next_week)
            logging.info(f"Number of sent service notifications: {cnt}")
            logging.info("-------------------------------------------------")

    logging.info("nuLiga Helper finished")
    logging.info("#################################################")


if __name__ == "__main__":
    main()
