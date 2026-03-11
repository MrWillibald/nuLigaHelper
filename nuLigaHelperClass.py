# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# A python tool in planning home games for Handball clubs
# Functions:
# - Read home game plan for sports stadium from nuLiga
# - Update game plan Excel document with judge scheduling
# - Send notifications for tasks
# - Send notifications to team leaders
# - Send notification to referee planner
# - Send newspaper article to local newspaper
# ---------------------------------------------------------------
# Created by: MrWillibald, assisted by Claude Sonnet 4.6
# Version 0.28
# Info: Refactored DataFrame access and helper methods
# Date: 11.03.2026
# ---------------------------------------------------------------

# scraping libs
import requests
import datetime

# data table libs
import pandas as pd
import numpy as np

# messaging libs
import smtplib
from email.utils import formataddr
from email.message import EmailMessage
from twilio.rest import Client

# file handling libs
import dropbox
import os

# additional libs
import json
import logging

# Version string
VERSION = "0.28"
# Debug flag
DEBUG_FLAG = False
# Change day flag
CHANGE_DAY = False


class nuLigaHomeGames:
    """nuLigaHelper class for home game planning and scheduled job notifications"""

    # ---------------------------------------------------------------------------
    # Helper: DataFrame row lookup
    # ---------------------------------------------------------------------------

    def _get_field(self, table: pd.DataFrame, game_nr: int, col: str):
        """Return the first value of `col` in the row where _colNr == game_nr."""
        return table.loc[table[self._colNr] == game_nr, col].values[0]

    def _build_receiver(
        self,
        row: pd.Series,
        name_col: str,
        mail_col: str,
        task_col: str,
        partner_col: str | None = None,
    ) -> dict:
        """Build a receiver dict from a single DataFrame row."""
        receiver = {
            "name": row[name_col],
            "mail": row[mail_col],
            "task": task_col,
        }
        if partner_col is not None:
            receiver["partner"] = row[partner_col]
        return receiver

    def _get_game_row(self, table: pd.DataFrame, game_nr: int) -> pd.Series:
        """Return the single row for game_nr as a Series."""
        return table.loc[table[self._colNr] == game_nr].iloc[0]

    def _build_receivers_for_game(
        self, row: pd.Series, include_shop: bool = True
    ) -> list[dict]:
        """
        Build the standard list of receivers (judges, shop, security, cleaning)
        from a single game row. Set include_shop=False to omit shop roles.
        """
        roles = [
            (self._colJudge1, self._colMailJudge1),
            (self._colJudge2, self._colMailJudge2),
        ]
        if include_shop:
            roles += [
                (self._colShop1, self._colMailShop1),
                (self._colShop2, self._colMailShop2),
            ]
        roles += [
            (self._colSecurity, self._colMailSecurity),
            (self._colCleaning, self._colMailCleaning),
        ]
        return [
            self._build_receiver(row, name_col, mail_col, name_col)
            for name_col, mail_col in roles
        ]

    # ---------------------------------------------------------------------------
    # Helper: Notification dispatch
    # ---------------------------------------------------------------------------

    def _dispatch_notification(
        self,
        receiver: dict,
        subject: str,
        mail_body: str,
        sms_body: str,
        game_nr: int,
        mail_id: str | None = None,
        mail_password: str | None = None,
    ) -> int:
        """
        Send an e-mail or SMS to a single receiver depending on their contact info.
        Returns 1 if a message was sent, 0 otherwise.
        """
        mail_id = mail_id or self.mail_ID
        mail_password = mail_password or self.mail_password
        contact = receiver["mail"]

        if not isinstance(contact, str):
            return 0

        if "@" in contact:
            msg = EmailMessage()
            msg["From"] = formataddr((self.mail_name, mail_id))
            msg["Subject"] = subject
            msg["To"] = formataddr((receiver["name"], contact))
            msg.set_content(mail_body)
            self.send_Mail(msg, mail_id, mail_password)
            logging.info(
                f"E-Mail sent to {receiver['name']}, {receiver.get('task', '')}, {contact}"
            )
            return 1

        if "+" in contact:
            self.send_SMS(contact, sms_body)
            logging.info(
                f"SMS sent to {receiver['name']}, {receiver.get('task', '')}, {contact}"
            )
            return 1

        logging.warning(
            f"No valid phone number or email address available at game {game_nr}"
        )
        return 0

    # ---------------------------------------------------------------------------
    # Date / Season helpers
    # ---------------------------------------------------------------------------

    def set_today(self, today: datetime.date):
        """Set today's date. Intended for debugging."""
        self._today = today
        self._year = today.year if today.month >= 7 else today.year - 1
        self.__dictSeason = {
            "part1": str(self._year),
            "part2": str(self._year + 1),
            "part3": str(self._year + 1)[-2:],
        }
        self.__season = "{part1}%2F{part3}".format(**self.__dictSeason)
        self.file = "Heimspielplan_{part1}_{part3}.xlsx".format(**self.__dictSeason)

    def get_today(self) -> datetime.date:
        """Return today's date."""
        return self._today

    # ---------------------------------------------------------------------------
    # Initialisation
    # ---------------------------------------------------------------------------

    def __init__(self, *args, **kwargs):
        """Initialize class instance with necessary strings."""
        self.__version = VERSION

        logging.info("Initialization of nuLiga Helper")
        logging.info("Version: " + self.__version)

        self.set_today(datetime.date.today())
        if DEBUG_FLAG or CHANGE_DAY:
            self.set_today(datetime.date(2025, 11, 21))

        config_path = os.path.join(os.path.dirname(__file__), "config.json")
        with open(config_path, encoding="utf-8") as json_config_file:
            config = json.load(json_config_file)

        club = config["club"]
        self.__dict__.update(club["info"])
        self.__dict__.update(club["email"])
        self.__dict__.update(club["dropbox"])
        self.dbc = dropbox.Dropbox(self.dropbox_token)
        self.__dict__.update(club["twilio"])
        self.__dict__.update(club["columns"])
        self.__dict__.update(club["texts"])

        logging.info("Initialization completed")

    # ---------------------------------------------------------------------------
    # Communication
    # ---------------------------------------------------------------------------

    def send_Mail(self, msg: EmailMessage, ID: str, password: str):
        """Send E-Mail via specified SMTP server."""
        if DEBUG_FLAG:
            return 0
        with smtplib.SMTP_SSL(self.smtpserver) as server:
            server.login(ID, password)
            server.send_message(msg)
            server.quit()

    def send_SMS(self, toaddr: str, text: str):
        """Send SMS via specified Twilio account."""
        if DEBUG_FLAG:
            return 0
        client = Client(self.twilio_sid, self.twilio_token)
        message = client.messages.create(
            messaging_service_sid=self.twilio_service_ID, body=text, to=toaddr
        )
        return message

    # ---------------------------------------------------------------------------
    # Dropbox
    # ---------------------------------------------------------------------------

    def get_fromDropbox(self):
        """Download file from specified Dropbox account."""
        remote_path = f"/{self.dropbox_folder}/{self.file}"
        try:
            self.dbc.files_download_to_file(self.file, remote_path)
        except dropbox.exceptions.ApiError:
            logging.warning(
                "Error while loading judge schedule from Dropbox, new schedule is created"
            )
        logging.info("Judge schedule loaded and saved successfully from Dropbox")

    def upload_toDropbox(self):
        """Upload file to specified Dropbox account."""
        remote_path = f"/{self.dropbox_folder}/{self.file}"
        with open(self.file, "rb") as f:
            data = f.read()
            self.dbc.files_upload(
                data, remote_path, mode=dropbox.files.WriteMode.overwrite
            )
        try:
            rmf = os.remove(self.file)
        except OSError:
            logging.warning(rmf)
        logging.info(
            "Judge schedule successfully uploaded to Dropbox and cleaned locally"
        )

    # ---------------------------------------------------------------------------
    # Game tables
    # ---------------------------------------------------------------------------

    def get_onlineTable(self):
        """Scrape Hallenspielplan for specified sports hall from BHV website."""
        logging.info("Read current home game plan from BHV Hallenspielplan website")

        parameters = {
            "club": self.clubId,
            "searchType": "1",
            "searchTimeRangeFrom": "01.09." + self.__dictSeason["part1"],
            "searchTimeRangeTo": "01.07." + self.__dictSeason["part2"],
            "onlyHomeMeetings": "false",
        }
        result = requests.post(
            "https://bhv-handball.liga.nu/cgi-bin/WebObjects/nuLigaHBDE.woa/wa/clubMeetings",
            data=parameters,
        )

        table = pd.read_html(result.content, header=0, attrs={"class": "result-set"})[0]

        # Drop obsolete columns and rename
        table.drop(table.columns[[9, 10, 11]], axis=1, inplace=True)
        table.columns = [
            self._colDay, self._colDate, self._colTime, self._colHall,
            self._colNr, self._colAK, self._colHome, self._colGuest, self._colScore,
        ]

        table[self._colHall] = table[self._colHall].astype(str)
        table[[self._colDay, self._colDate]] = table[[self._colDay, self._colDate]].ffill()

        # Keep only games in own halls
        mask = table[self._colHall].apply(
            lambda game: any(hall in game for hall in self.hallIds)
        )
        table = table[mask]

        # Drop "spielfrei" rows (no game number)
        table = table[table[self._colNr].notna()]

        # Convert column types
        table[self._colDate] = table[self._colDate].astype(str)
        table[self._colTime] = table[self._colTime].astype(str)
        table[self._colHall] = table[self._colHall].astype(int)
        table[self._colNr] = table[self._colNr].astype(int)
        table[self._colScore] = table[self._colScore].astype(str) + "\t"

        # Add empty scheduling columns
        extra_cols = [
            self._colJTeam, self._colJMV, self._colMailJMV,
            self._colJudge1, self._colMailJudge1,
            self._colJudge2, self._colMailJudge2,
            self._colShop1, self._colMailShop1,
            self._colShop2, self._colMailShop2,
            self._colSecurity, self._colMailSecurity,
            self._colCleaning, self._colMailCleaning,
        ]
        for col in extra_cols:
            table[col] = np.empty(len(table), dtype=str)

        self.onlineTable = table.reset_index(drop=True)
        logging.info("Current home game plan loaded")

    def get_gameTable(self):
        """Read scheduled jobs from downloaded or local file."""
        logging.info("Read local judge schedule")
        try:
            self.gameTable = pd.read_excel(self.file, index_col=0, header=0).T
            self.gameTable = self.gameTable.astype({
                self._colDate: str,
                self._colTime: str,
                self._colMailJMV: str,
                self._colMailJudge1: str,
                self._colMailJudge2: str,
                self._colMailShop1: str,
                self._colMailShop2: str,
                self._colMailSecurity: str,
                self._colMailCleaning: str,
            })
            logging.info("Judge schedule available")
        except OSError:
            self.gameTable = self.onlineTable
            logging.warning("Judge schedule not available, empty schedule created")
        return 0

    # ---------------------------------------------------------------------------
    # Change handlers
    # ---------------------------------------------------------------------------

    def datum_shift_handler(
        self, game: int, oldDate: str, oldTime: str, newDate: str, newTime: str
    ):
        if (newDate != oldDate) or (newTime != oldTime):
            logging.info(
                f"Game {game} is shifted! "
                f"Old date: {oldDate} {oldTime} — New date: {newDate} {newTime}"
            )
            self.send_ShfitNotification(game, oldDate, oldTime, newDate, newTime)
        return 0

    def no_referee_handler(
        self, game: int, date: str, time: str, oldRef: str, newRef: str
    ):
        if (newRef != oldRef) and ("§77" in newRef):
            logging.info(
                f"Game {game} on {date} {time} will not get a scheduled referee!"
            )
            self.send_RefNotification(game, date, time)
        return 0

    # ---------------------------------------------------------------------------
    # Table merge
    # ---------------------------------------------------------------------------

    def merge_tables(self):
        """Merge up-to-date Hallenspielplan with scheduled jobs file."""
        logging.info("Update home game plan")
        sendError = False

        for game in self.onlineTable[self._colNr]:
            try:
                old_row = self._get_game_row(self.gameTable, game)
                oldDate, oldTime, oldRef = (
                    old_row[self._colDate],
                    old_row[self._colTime],
                    old_row[self._colScore],
                )

                # Copy scheduling columns from existing gameTable into onlineTable
                judges = self.gameTable.loc[
                    self.gameTable[self._colNr] == game, self._colJTeam:
                ]
                self.onlineTable.loc[
                    self.onlineTable[self._colNr] == game, self._colJTeam:
                ] = judges.values[0]
                logging.info(f"Game {game} merged with schedule")

                new_row = self._get_game_row(self.onlineTable, game)
                newDate, newTime, newRef = (
                    new_row[self._colDate],
                    new_row[self._colTime],
                    new_row[self._colScore],
                )

                self.datum_shift_handler(game, oldDate, oldTime, newDate, newTime)
                self.no_referee_handler(game, newDate, newTime, oldRef, newRef)

            except IndexError:
                ak = self._get_field(self.onlineTable, game, self._colAK)
                if ak != "GE":
                    sendError = True
                    logging.warning(
                        "Spielnummer not contained in home schedule, please correct manually!"
                    )

        if sendError:
            msg = EmailMessage()
            msg["From"] = formataddr((self.mail_name, self.mail_ID))
            msg["Subject"] = self.mailErrorSubject
            msg["To"] = formataddr(("Manu", self.mail_ID))
            msg.set_content(self.mailError)
            self.send_Mail(msg, self.mail_ID, self.mail_password)

        self.gameTable = self.onlineTable
        logging.info("Update home game plan completed")
        return 0

    # ---------------------------------------------------------------------------
    # Excel output
    # ---------------------------------------------------------------------------

    def write_toXlsx(self):
        """Write updated job scheduling to local *.xlsx file."""
        writer = pd.ExcelWriter(self.file, engine="xlsxwriter")
        self.writeTable = self.gameTable.transpose()
        self.writeTable.to_excel(writer, sheet_name="Heimspielplan")

        worksheet = writer.sheets["Heimspielplan"]
        workbook = writer.book
        workbook.encoding = self._enc
        formatText = workbook.add_format({"num_format": "@"})

        worksheet.set_column(0, 100, 35)
        for row_idx in range(11, 24, 2):
            worksheet.set_row(row_idx, None, formatText)
        worksheet.freeze_panes(0, 1)

        writer.close()
        logging.info("Judge schedule saved locally")

    # ---------------------------------------------------------------------------
    # Notifications: Newspaper article
    # ---------------------------------------------------------------------------

    def send_Article(self, date: str, day: str, articleDate: str) -> int:
        """Send notification article to local newspaper."""
        cnt = 0
        tournamentMI = False
        tournamentGE = False

        noteTable = self.gameTable[
            (self.gameTable[self._colDate] == date)
            & (self.gameTable[self._colGuest] != "spielfrei")
        ].dropna(how="all")

        schedule = ""
        for _, row in noteTable.iterrows():
            team_raw = row[self._colAK]
            team = {"F": "Damen", "M": "Herren"}.get(team_raw, team_raw)
            home = row[self._colHome]
            guest = row[self._colGuest]
            time_str = row[self._colTime].strip(" v").strip(" t")

            if team == "MI" and not tournamentMI:
                schedule += f"Ab {time_str} Spielfest der Minis\n"
                tournamentMI = True
                cnt += 1
            elif team == "GE" and not tournamentGE:
                schedule += f"Ab {time_str} Turnier der gemischten E-Jugend\n"
                tournamentGE = True
                cnt += 1
            elif team not in ("GE", "MI"):
                schedule += f"{time_str} {team} {home} - {guest}\n"
                cnt += 1

        logging.info(f"Send newspaper article to {self.mailAddrNewspaper}")
        msg = EmailMessage()
        msg["From"] = formataddr((self.mail_name, self.mail_ID))
        msg["To"] = self.mailAddrNewspaper
        msg["Subject"] = self.mailNewspaperSubject
        msg.set_content(self.mailNewspaper.format(articleDate, day, date, schedule))
        self.send_Mail(msg, self.mail_ID, self.mail_password)

        logging.info(
            f"Newspaper article for {cnt} games at {date} sent to {self.mailAddrNewspaper}"
        )
        return cnt

    # ---------------------------------------------------------------------------
    # Notifications: Game-day judge / service notifications
    # ---------------------------------------------------------------------------

    def send_Notifications(self, date: str) -> int:
        """Send notifications to game judges via SMS or E-Mail."""
        cnt = 0
        noteTable = self.gameTable[
            (self.gameTable[self._colDate] == date)
            & (self.gameTable[self._colGuest] != "spielfrei")
        ].dropna(how="all")

        for _, row in noteTable.iterrows():
            game = row[self._colNr]
            ak = row[self._colAK]
            mv = row[self._colJMV]
            mailMV = row[self._colMailJMV]
            home = row[self._colHome]
            guest = row[self._colGuest]
            strTime = row[self._colTime]

            receivers = self._build_receivers_for_game(row)

            for receiver in receivers:
                cnt += self._dispatch_notification(
                    receiver,
                    subject=f"Benachrichtigung Dienst {receiver['task']}",
                    mail_body=self.mailTask.format(
                        receiver["name"], date, receiver["task"], ak, home, guest, strTime
                    ),
                    sms_body=self.textTask.format(
                        receiver["name"], date, receiver["task"], ak, strTime
                    ),
                    game_nr=game,
                )

            # Notification to MV
            mv_receiver = {"name": mv, "mail": mailMV}
            cnt += self._dispatch_notification(
                mv_receiver,
                subject=self.mailMVSubject,
                mail_body=self.mailMV.format(
                    mv, row[self._colJTeam], date,
                    receivers[0]["name"], receivers[1]["name"],
                    ak, home, guest, strTime,
                ),
                sms_body=self.textMV.format(
                    mv, row[self._colJTeam], date,
                    receivers[0]["name"], receivers[1]["name"],
                    ak, strTime,
                ),
                game_nr=game,
            )

        return cnt

    # ---------------------------------------------------------------------------
    # Notifications: Referee coordinator
    # ---------------------------------------------------------------------------

    def send_RefNotification(self, game: int, date: str, time: str) -> int:
        """Send referee notification to referee coordinator."""
        cnt = 0
        row = self._get_game_row(self.onlineTable, game)
        ak = row[self._colAK]
        mv = row[self._colJMV]
        mailMv = row[self._colMailJMV]

        receivers = list(self.mailRefCoordTargets)
        if isinstance(receivers, list) and pd.notna(mv):
            receivers.append({"Name": str(mv), "Address": str(mailMv)})

        all_names = ", ".join(r["Name"] for r in receivers)

        for receiver in receivers:
            text = self.mailRefCoord.format(
                receiver["Name"], ak, date, time, all_names
            )
            cnt += self._dispatch_notification(
                {"name": receiver["Name"], "mail": receiver["Address"]},
                subject=self.mailRefCoordSubject,
                mail_body=text,
                sms_body=text,
                game_nr=game,
            )

        return cnt

    # ---------------------------------------------------------------------------
    # Notifications: Early service notifications
    # ---------------------------------------------------------------------------

    def send_ServiceNotifications(self, date: str) -> int:
        """Send early notifications to service via SMS or E-Mail."""
        cnt = 0
        noteTable = self.gameTable[
            (self.gameTable[self._colDate] == date)
            & (self.gameTable[self._colGuest] != "spielfrei")
        ].dropna(how="all")

        # Only the first game is relevant for early service notification
        row = noteTable.iloc[0]
        game = row[self._colNr]
        ak = row[self._colAK]
        home = row[self._colHome]
        guest = row[self._colGuest]
        strTime = row[self._colTime]

        receivers = [
            self._build_receiver(
                row, self._colShop1, self._colMailShop1, self._colShop1,
                partner_col=self._colShop2
            ),
            self._build_receiver(
                row, self._colShop2, self._colMailShop2, self._colShop2,
                partner_col=self._colShop1
            ),
        ]

        for receiver in receivers:
            cnt += self._dispatch_notification(
                receiver,
                subject=f"Vorbereitung Dienst {receiver['task']}",
                mail_body=self.mailEarlyTask.format(
                    receiver["name"], date, receiver["task"], ak, home, guest,
                    receiver["partner"], strTime, receiver["partner"],
                ),
                sms_body=self.textEarlyTask.format(
                    receiver["name"], date, receiver["task"], ak,
                    receiver["partner"], strTime, receiver["partner"],
                ),
                game_nr=game,
                mail_id=self.mail_saleID,
                mail_password=self.mail_salePassword,
            )

        return cnt

    # ---------------------------------------------------------------------------
    # Notifications: Pre-notifications (one week ahead)
    # ---------------------------------------------------------------------------

    def send_PreNotifications(self, date: str) -> int:
        """Send early notifications to game judges via SMS or E-Mail."""
        cnt = 0
        noteTable = self.gameTable[
            (self.gameTable[self._colDate] == date)
            & (self.gameTable[self._colGuest] != "spielfrei")
        ].dropna(how="all")

        for gameNr, (_, row) in enumerate(noteTable.iterrows()):
            game = row[self._colNr]
            ak = row[self._colAK]
            home = row[self._colHome]
            guest = row[self._colGuest]
            strTime = row[self._colTime]

            # Shop roles are excluded for the first game (notified via send_ServiceNotifications)
            receivers = self._build_receivers_for_game(row, include_shop=(gameNr > 0))

            for receiver in receivers:
                cnt += self._dispatch_notification(
                    receiver,
                    subject=f"Benachrichtigung Dienst {receiver['task']}",
                    mail_body=self.mailPreTask.format(
                        receiver["name"], date, receiver["task"], ak, home, guest, strTime
                    ),
                    sms_body=self.textPreTask.format(
                        receiver["name"], date, receiver["task"], ak, strTime
                    ),
                    game_nr=game,
                )

        return cnt

    # ---------------------------------------------------------------------------
    # Notifications: Date-shift notifications
    # ---------------------------------------------------------------------------

    def send_ShfitNotification(
        self, game: int, oldDate: str, oldTime: str, newDate: str, newTime: str
    ) -> int:
        """Send notifications on shifted datum to game judges via SMS or E-Mail."""
        cnt = 0
        noteTable = self.gameTable[
            self.gameTable[self._colGuest] != "spielfrei"
        ].dropna(how="all")

        row = self._get_game_row(noteTable, game)
        ak = row[self._colAK]
        home = row[self._colHome]
        guest = row[self._colGuest]

        receivers = self._build_receivers_for_game(row)

        for receiver in receivers:
            cnt += self._dispatch_notification(
                receiver,
                subject=f"Benachrichtigung Verschiebung Dienst {receiver['task']}",
                mail_body=self.mailShifted.format(
                    receiver["name"], receiver["task"], ak, home, guest,
                    oldDate, oldTime, newDate, newTime,
                ),
                sms_body=self.textShifted.format(
                    receiver["name"], receiver["task"],
                    oldDate, oldTime, newDate, newTime,
                ),
                game_nr=game,
            )

        return cnt


# ---------------------------------------------------------------
#  Main program
# ---------------------------------------------------------------
if __name__ == "__main__":

    # Initialize logger
    logging.basicConfig(
        format="%(asctime)s - %(levelname)s - %(message)s",
        filename="helper.log",
        level=logging.DEBUG,
    )
    logging.getLogger().addHandler(logging.StreamHandler())
    # Limit lib logging to warnings
    logging.getLogger("twilio.http_client").setLevel(logging.WARNING)
    logging.getLogger("requests").setLevel(logging.WARNING)
    logging.getLogger("urllib3").setLevel(logging.WARNING)

    # Initialize class instance
    logging.info("#################################################")
    nlh = nuLigaHomeGames()
    logging.info("-------------------------------------------------")

    # Download Heimspielplan from Dropbox
    nlh.get_fromDropbox()

    # Check nuLiga Hallenplan and update Heimspielplan
    nlh.get_onlineTable()
    nlh.get_gameTable()
    logging.info("-------------------------------------------------")
    nlh.merge_tables()
    logging.info("-------------------------------------------------")

    # Save Heimspielplan as Excel-File
    nlh.write_toXlsx()

    # Upload Heimspielplan to Dropbox
    nlh.upload_toDropbox()
    logging.info("-------------------------------------------------")

    """
    # Check if newspaper article has to be sent
    gameDateSa      = nlh.get_today() + datetime.timedelta(days=9)
    strGameDateSa   = gameDateSa.strftime("%d.%m.%Y")
    strGameDaySa    = gameDateSa.strftime("%A")
    gameDateSo      = nlh.get_today() + datetime.timedelta(days=10)
    strGameDateSo   = gameDateSo.strftime("%d.%m.%Y")
    strGameDaySo    = gameDateSo.strftime("%A")

    # Send newspaper article for Saturday
    if nlh.gameTable[nlh._colDate].str.contains(strGameDateSa).any() & (strGameDaySa == "Saturday"):
        articleDate     = gameDateSa + datetime.timedelta(days=-1)
        cnt             = nlh.send_Article(strGameDateSa, "Samstag", articleDate.strftime("%d.%m.%Y"))

    # Send newspaper article for Sunday
    elif nlh.gameTable[nlh._colDate].str.contains(strGameDateSo).any() & (strGameDaySo == "Sunday"):
        articleDate     = gameDateSo + datetime.timedelta(days=-2)
        cnt             = nlh.send_Article(strGameDateSo, "Sonntag", articleDate.strftime("%d.%m.%Y"))
    """

    # Check if judge notifications have to be send
    tomorrow = nlh.get_today() + datetime.timedelta(days=1)
    strTomorrow = tomorrow.strftime("%d.%m.%Y")

    if nlh.gameTable[nlh._colDate].str.contains(strTomorrow).any():
        cnt = nlh.send_Notifications(strTomorrow)
        logging.info(f"Number of sent service notifications: {cnt}")
        logging.info("-------------------------------------------------")

    # Check if referee notifications have to be send
    if not nlh.gameTable[
        nlh.gameTable[nlh._colDate].str.contains(strTomorrow)
        & nlh.gameTable[nlh._colScore].str.contains("§77")
    ].empty:
        cnt = nlh.send_RefNotification(strTomorrow)
        logging.info(f"Number of required home referees: {cnt}")
        logging.info("-------------------------------------------------")

    # Check if early catering notifications have to be send
    nextWeek = nlh.get_today() + datetime.timedelta(days=7)
    strNextWeek = nextWeek.strftime("%d.%m.%Y")

    if nlh.gameTable[nlh._colDate].str.contains(strNextWeek).any():
        cnt = nlh.send_ServiceNotifications(strNextWeek)
        cnt += nlh.send_PreNotifications(strNextWeek)
        logging.info(f"Number of sent service notifications: {cnt}")
        logging.info("-------------------------------------------------")

    logging.info("nuLiga Helper finished")
    logging.info("#################################################")