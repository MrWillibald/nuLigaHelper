# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Main entry point: daily scheduled job
#
# - Scrapes home games from nuLiga and syncs them into SQLite DB
# - Reports shifts, missing referees and unknown new games
# - Sends notifications via e-mail/SMS based on DB assignments
# - Publishes a validated online SQLite snapshot to Dropbox
# ---------------------------------------------------------------

import datetime
import logging
import os
from typing import Any

import dropbox

import backup
import common
import daily_lock
import db
from notifier import Notifier
from scraper import fetch_home_games


def reportable_new_games(events: db.SyncEvents) -> list[db.GameEvent]:
    """Exclude tournament age class GE without collapsing duplicate game numbers."""
    return [event for event in events.new_games if event.ak != "GE"]


class DailyJobError(RuntimeError):
    """Raised after all safe daily work has run and one or more stages failed."""

    failures: tuple[tuple[str, Exception], ...]

    def __init__(self, failures: list[tuple[str, Exception]]) -> None:
        self.failures = tuple(failures)
        stages = ", ".join(stage for stage, _error in failures)
        super().__init__(f"Daily job failed in: {stages}")


def _backup_database(
    club_cfg: dict[str, Any],
    database_path: str,
    today: datetime.date,
) -> backup.BackupResult:
    dropbox_cfg: dict[str, Any] = club_cfg["dropbox"]
    client_factory = lambda: dropbox.Dropbox(
        dropbox_cfg["dropbox_token"], timeout=30
    )
    if "dated_retention" in dropbox_cfg:
        result = backup.backup_database_to_dropbox(
            database_path,
            dropbox_cfg["dropbox_folder"],
            client_factory=client_factory,
            retention=dropbox_cfg["dated_retention"],
            backup_date=today,
        )
    else:
        result = backup.backup_database_to_dropbox(
            database_path,
            dropbox_cfg["dropbox_folder"],
            client_factory=client_factory,
            backup_date=today,
        )
    logging.info(
        "Database backup successfully uploaded to Dropbox (%s bytes, latest: %s)",
        result.byte_count,
        result.paths.latest,
    )
    return result


def _log_backup_error(error: backup.BackupError) -> None:
    completed = ", ".join(stage.value for stage in error.completed_stages) or "none"
    logging.error(
        "Database backup failed at stage %s (completed: %s): %s",
        error.stage.value,
        completed,
        error,
    )
    for secondary in error.secondary_failures:
        logging.error("Secondary backup failure: %s", secondary)


def _send_notifications(
    club_cfg: dict[str, Any],
    engine: Any,
    season_year: int,
    today: datetime.date,
    events: db.SyncEvents,
) -> None:
    """Run the unchanged notification sequence in a fresh, non-writing session."""
    with db.Session(engine) as session:
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
        notifier.notify_new_games(reportable_new_games(events))

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


def main():
    if not os.environ.get("NULIGAHELPER_SECRET"):
        raise RuntimeError("NULIGAHELPER_SECRET muss gesetzt sein.")
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
    db_path = club_cfg.get("database", {}).get("path", db.DEFAULT_DB_PATH)
    resolved_db_path = db.resolve_db_path(db_path)

    today = common.effective_today()
    season_year = common.season_year_for(today)

    try:
        with daily_lock.daily_run_lock(resolved_db_path):
            engine = db.make_engine(resolved_db_path)
            db.init_db(engine)

            # Scraping is network I/O and therefore happens before a DB session opens.
            scraped = fetch_home_games(club_cfg["info"], season_year)
            with db.Session(engine) as session:
                events = db.sync_games(session, scraped, season_year)
                # sync_games currently commits itself; this keeps orchestration safe
                # if that implementation detail changes.
                session.commit()

            failures: list[tuple[str, Exception]] = []
            try:
                _backup_database(club_cfg, resolved_db_path, today)
            except backup.BackupError as error:
                _log_backup_error(error)
                failures.append((f"backup:{error.stage.value}", error))
            logging.info("-------------------------------------------------")

            try:
                _send_notifications(club_cfg, engine, season_year, today, events)
            except Exception as error:
                logging.exception("Notification delivery failed: %s", error)
                failures.append(("notifications", error))

            if failures:
                logging.error(
                    "nuLiga Helper finished unsuccessfully after %d fatal stage(s)",
                    len(failures),
                )
                raise DailyJobError(failures) from None

            logging.info("nuLiga Helper finished")
            logging.info("#################################################")
    except daily_lock.DailyRunLockError as error:
        logging.error("Daily run lock failed: %s", error)
        raise


if __name__ == "__main__":
    try:
        main()
    except (DailyJobError, daily_lock.DailyRunLockError):
        raise SystemExit(1) from None
