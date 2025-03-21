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
# Created by: MrWillibald
# Version 0.26
# Info: Send referee notice on change of scheduled referee
# Date: 17.03.2025
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

# additonal libs
import json
import logging

# Version string
VERSION = "0.26"
# Debug flag
DEBUG_FLAG = False
# Change day flag
CHANGE_DAY = False


class nuLigaHomeGames:
    """nuLigaHelper class for home game planning and scheduled job notifications"""

    def set_today(self, today):
        """Set todays date. Intended for debugging."""
        self._today = today
        if self._today.month >= 7:
            self._year = self._today.year
        else:
            self._year = today.year - 1
        self.__dictSeason = {
            "part1": str(self._year),
            "part2": str(self._year + 1),
            "part3": str(self._year + 1)[-2:],
        }
        self.__season = "{part1}%2F{part3}".format(**self.__dictSeason)
        self.file = "Heimspielplan_{part1}_{part3}.xlsx".format(**self.__dictSeason)

    def get_today(self):
        """Return todays date"""
        return self._today

    def __init__(self, *args, **kwargs):
        """Initialize class instance with necessary strings"""
        self.__version = VERSION

        logging.info("Initialization of nuLiga Helper")
        logging.info("Version: " + self.__version)

        # Set up dates and strings
        self.set_today(datetime.date.today())
        if DEBUG_FLAG or CHANGE_DAY:
            self.set_today(datetime.date(2024, 8, 20))

        # New config workflow
        with open(
            os.path.join(os.path.dirname(__file__), "config.json"), encoding="utf-8"
        ) as json_config_file:
            config = json.load(json_config_file)

        # Club properties
        self.__dict__.update(config["club"]["info"])
        # Email account properties
        self.__dict__.update(config["club"]["email"])
        # Dropbox account properties
        self.__dict__.update(config["club"]["dropbox"])
        self.dbc = dropbox.Dropbox(self.dropbox_token)
        # Twilio account properties
        self.__dict__.update(config["club"]["twilio"])
        # Column names
        self.__dict__.update(config["club"]["columns"])
        # Texts
        self.__dict__.update(config["club"]["texts"])

        logging.info("Initialization completed")

    def send_Mail(self, msg, ID, password):
        """Send E-Mail via specified SMTP server"""
        if True == DEBUG_FLAG:
            return 0
        with smtplib.SMTP_SSL(self.smtpserver) as server:
            # server.set_debuglevel(True)
            server.login(ID, password)
            server.send_message(msg)
            server.quit()

    def send_SMS(self, toaddr, text):
        """Send SMS via specified Twilio account"""
        if True == DEBUG_FLAG:
            return 0
        client = Client(self.twilio_sid, self.twilio_token)
        message = client.messages.create(
            messaging_service_sid=self.twilio_service_ID, body=text, to=toaddr
        )
        return message

    def get_fromDropbox(self):
        """Download file from specified Dropbox account"""
        try:
            self.dbc.files_download_to_file(
                self.file, "/" + self.dropbox_folder + "/" + self.file
            )
        except dropbox.exceptions.ApiError:
            logging.warning(
                "Error while loading judge schedule from Dropbox, new schedule is created"
            )
        logging.info("Judge schedule loaded and saved successfully from Dropbox")

    def upload_toDropbox(self):
        """Upload file to specified Dropbox account"""
        with open(self.file, "rb") as f:
            data = f.read()
            self.dbc.files_upload(
                data,
                "/" + self.dropbox_folder + "/" + self.file,
                mode=dropbox.files.WriteMode.overwrite,
            )
        try:
            rmf = os.remove(self.file)
        except OSError:
            logging.warning(rmf)
        logging.info(
            "Judge schedule successfully uploaded to Dropbox and cleaned locally"
        )

    def get_onlineTable(self):
        """Scrape Hallenspielplan for specified sports hall from BHV website"""
        logging.info("Read current home game plan from BHV Hallenspielplan website")
        lGames = list()
        # read home games of season (http request)
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
        resultTable = pd.read_html(
            result.content, header=0, attrs={"class": "result-set"}
        )
        table = resultTable[0]
        # drop obsolete columns and rename
        table.drop(table.columns[[9, 10, 11]], axis=1, inplace=True)
        table.columns = [
            self._colDay,
            self._colDate,
            self._colTime,
            self._colHall,
            self._colNr,
            self._colAK,
            self._colHome,
            self._colGuest,
            self._colScore,
        ]
        # convert column 3 to str
        table[self._colHall] = table[self._colHall].apply(str)
        # fill dates
        table[[self._colDay, self._colDate]] = table[
            [self._colDay, self._colDate]
        ].ffill()
        # find games in own halls and only keep them
        mask = np.array(
            [
                any(hall in game for hall in self.hallIds)
                for game in table[self._colHall]
            ]
        )
        table.drop(table[np.invert(mask)].index, inplace=True)
        # drop spielfrei
        mask = np.array([np.isnan(gamenr) for gamenr in table[self._colNr]])
        table.drop(table[mask].index, inplace=True)
        # convert columns
        table[self._colDate] = table[self._colDate].apply(str)
        table[self._colTime] = table[self._colTime].apply(str)
        table[self._colHall] = table[self._colHall].apply(int)
        table[self._colNr] = table[self._colNr].apply(int)
        lGames.append(table)
        self.onlineTable = pd.concat(lGames)
        self.onlineTable.index = range(len(self.onlineTable[self._colDay]))
        self.onlineTable[self._colScore] = (
            self.onlineTable[self._colScore].astype(str) + "\t"
        )
        # add additional columns
        kwargs = {
            self._colJTeam: np.empty(len(self.onlineTable[self._colDay]), dtype=str)
        }
        self.onlineTable = self.onlineTable.assign(**kwargs)
        kwargs = {
            self._colJMV: np.empty(len(self.onlineTable[self._colDay]), dtype=str)
        }
        self.onlineTable = self.onlineTable.assign(**kwargs)
        kwargs = {
            self._colMailJMV: np.empty(len(self.onlineTable[self._colDay]), dtype=str)
        }
        self.onlineTable = self.onlineTable.assign(**kwargs)
        kwargs = {
            self._colJudge1: np.empty(len(self.onlineTable[self._colDay]), dtype=str)
        }
        self.onlineTable = self.onlineTable.assign(**kwargs)
        kwargs = {
            self._colMailJudge1: np.empty(
                len(self.onlineTable[self._colDay]), dtype=str
            )
        }
        self.onlineTable = self.onlineTable.assign(**kwargs)
        kwargs = {
            self._colJudge2: np.empty(len(self.onlineTable[self._colDay]), dtype=str)
        }
        self.onlineTable = self.onlineTable.assign(**kwargs)
        kwargs = {
            self._colMailJudge2: np.empty(
                len(self.onlineTable[self._colDay]), dtype=str
            )
        }
        self.onlineTable = self.onlineTable.assign(**kwargs)
        kwargs = {
            self._colShop1: np.empty(len(self.onlineTable[self._colDay]), dtype=str)
        }
        self.onlineTable = self.onlineTable.assign(**kwargs)
        kwargs = {
            self._colMailShop1: np.empty(len(self.onlineTable[self._colDay]), dtype=str)
        }
        self.onlineTable = self.onlineTable.assign(**kwargs)
        kwargs = {
            self._colShop2: np.empty(len(self.onlineTable[self._colDay]), dtype=str)
        }
        self.onlineTable = self.onlineTable.assign(**kwargs)
        kwargs = {
            self._colMailShop2: np.empty(len(self.onlineTable[self._colDay]), dtype=str)
        }
        self.onlineTable = self.onlineTable.assign(**kwargs)
        kwargs = {
            self._colSecurity: np.empty(len(self.onlineTable[self._colDay]), dtype=str)
        }
        self.onlineTable = self.onlineTable.assign(**kwargs)
        kwargs = {
            self._colMailSecurity: np.empty(
                len(self.onlineTable[self._colDay]), dtype=str
            )
        }
        self.onlineTable = self.onlineTable.assign(**kwargs)
        kwargs = {
            self._colCleaning: np.empty(len(self.onlineTable[self._colDay]), dtype=str)
        }
        self.onlineTable = self.onlineTable.assign(**kwargs)
        kwargs = {
            self._colMailCleaning: np.empty(
                len(self.onlineTable[self._colDay]), dtype=str
            )
        }
        self.onlineTable = self.onlineTable.assign(**kwargs)
        logging.info("Current home game plan loaded")

    def get_gameTable(self):
        """Read scheduled jobs from downloaded or local file"""
        logging.info("Read local judge schedule")
        try:
            # Transpose table on reading
            self.gameTable = pd.read_excel(self.file, index_col=0, header=0).T
            self.gameTable = self.gameTable.astype(
                {
                    self._colDate: str,
                    self._colTime: str,
                    self._colMailJMV: str,
                    self._colMailJudge1: str,
                    self._colMailJudge2: str,
                    self._colMailShop1: str,
                    self._colMailShop2: str,
                    self._colMailSecurity: str,
                    self._colMailCleaning: str,
                }
            )
            logging.info("Judge schedule available")
        except OSError:
            self.gameTable = self.onlineTable
            logging.warning("Judge schedule not available, empty schedule created")
        return 0
    
    def datum_shift_handler(self, game, oldDate, oldTime, newDate, newTime):
        if (newDate != oldDate) or (newTime != oldTime):
            logging.info(
                "Game "
                + str(game)
                + " is shifted! Old date: "
                + oldDate
                + " "
                + oldTime
                + " New date: "
                + newDate
                + " "
                + newTime
            )
            self.send_ShfitNotification(
                game, oldDate, oldTime, newDate, newTime
            )
        return 0
    
    def no_referee_handler(self, game, date, time, oldRef, newRef):
        if (newRef != oldRef) and ("Heim" in newRef):
            logging.info(
                "Game "
                + str(game)
                + " on "
                + date
                + " "
                + time
                + " will not get a scheduled referee!"
            )
            self.send_RefNotification(game, date, time)
        return 0

    def merge_tables(self):
        """Merge up-to-date Hallenspielplan with scheduled jobs file"""
        logging.info("Update home game plan")
        sendError = False
        for game in self.onlineTable[self._colNr]:
            try:
                # save old values
                oldDate = self.gameTable.loc[
                    self.gameTable[self._colNr] == game, self._colDate
                ].values[0]
                oldTime = self.gameTable.loc[
                    self.gameTable[self._colNr] == game, self._colTime
                ].values[0]
                oldRef = self.gameTable.loc[
                    self.gameTable[self._colNr] == game, self._colScore
                ].values[0]

                # merge tables
                judges = self.gameTable.loc[
                    self.gameTable[self._colNr] == game, self._colJTeam :
                ]
                self.onlineTable.loc[
                    self.onlineTable[self._colNr] == game, self._colJTeam :
                ] = judges.values[0]
                logging.info("Game " + str(game) + " merged with schedule")

                # get new values
                newDate = self.onlineTable.loc[
                    self.onlineTable[self._colNr] == game, self._colDate
                ].values[0]
                newTime = self.onlineTable.loc[
                    self.onlineTable[self._colNr] == game, self._colTime
                ].values[0]
                newRef = self.onlineTable.loc[
                    self.onlineTable[self._colNr] == game, self._colScore
                ].values[0]

                # handle changes
                self.datum_shift_handler(game, oldDate, oldTime, newDate, newTime)
                self.no_referee_handler(game, newDate, newTime, oldRef, newRef)
                
            except IndexError:
                # oTable = pTable
                if (
                    self.onlineTable.loc[
                        self.onlineTable[self._colNr] == game, self._colAK
                    ].values[0]
                    != "GE"
                ):
                    # send Error Notification
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

        # make game and online table identical if merging was successful
        self.gameTable = self.onlineTable
        logging.info("Update home game plan completed")

        return 0

    def write_toXlsx(self):
        """Write updated job scheduling to local *.xlsx file"""
        writer = pd.ExcelWriter(self.file, engine="xlsxwriter")
        # transpose table before writing for better reading
        self.writeTable = self.gameTable.transpose()
        self.writeTable.to_excel(writer, sheet_name="Heimspielplan")
        worksheet = writer.sheets["Heimspielplan"]
        workbook = writer.book
        workbook.encoding = self._enc
        formatText = workbook.add_format({"num_format": "@"})
        worksheet.set_column(0, 100, 35)
        worksheet.set_row(11, None, formatText)
        worksheet.set_row(13, None, formatText)
        worksheet.set_row(15, None, formatText)
        worksheet.set_row(17, None, formatText)
        worksheet.set_row(19, None, formatText)
        worksheet.set_row(21, None, formatText)
        worksheet.set_row(23, None, formatText)
        # freeze first column
        worksheet.freeze_panes(0, 1)
        # worksheet.autofilter(
        #    0, 0, self.writeTable.shape[0], self.writeTable.shape[1])
        writer.close()
        logging.info("Judge schedule saved locally")

    def send_Article(self, date, day, articleDate):
        """Send notification article to local newspaper"""
        cnt = 0
        team = []
        home = []
        guest = []
        strTime = []
        tournamentMI = False
        tournamentGE = False
        noteTable = self.gameTable[
            (self.gameTable[self._colDate] == date)
            & (self.gameTable[self._colGuest] != "spielfrei")
        ].dropna(how="all")
        for game in noteTable[self._colNr]:
            teamRaw = noteTable.loc[noteTable[self._colNr] == game, self._colAK].values[
                0
            ]
            if teamRaw == "F":
                team.append("Damen")
            elif teamRaw == "M":
                team.append("Herren")
            else:
                team.append(teamRaw)
            home.append(
                noteTable.loc[noteTable[self._colNr] == game, self._colHome].values[0]
            )
            guest.append(
                noteTable.loc[noteTable[self._colNr] == game, self._colGuest].values[0]
            )
            strTime.append(
                noteTable.loc[noteTable[self._colNr] == game, self._colTime].values[0]
            )
        logging.info("Send newspaper article to " + self.mailAddrNewspaper)
        msg = EmailMessage()
        msg["From"] = formataddr((self.mail_name, self.mail_ID))
        msg["To"] = self.mailAddrNewspaper
        msg["Subject"] = self.mailNewspaperSubject

        # Create schedule from game plan
        schedule = ""
        for i in range(0, len(home)):
            # single tournament information for MI and GE
            if ("MI" == team[i]) and (False == tournamentMI):
                schedule = (
                    schedule
                    + "Ab "
                    + strTime[i].strip(" v").strip(" t")
                    + " Spielfest der Minis\n"
                )
                tournamentMI = True
                cnt +=1
            elif ("GE" == team[i]) and (False == tournamentGE):
                schedule = (
                    schedule
                    + "Ab "
                    + strTime[i].strip(" v").strip(" t")
                    + " Turnier der gemischten E-Jugend\n"
                )
                tournamentGE = True
                cnt +=1
            elif ("GE" or "MI") != team[i]:
                schedule = (
                    schedule
                    + strTime[i].strip(" v").strip(" t")
                    + " "
                    + team[i]
                    + " "
                    + home[i]
                    + " - "
                    + guest[i]
                    + "\n"
                )
                cnt +=1

        # Message body created from mail text stored in config
        msg.set_content(self.mailNewspaper.format(articleDate, day, date, schedule))
        self.send_Mail(msg, self.mail_ID, self.mail_password)
        logging.info(
            "Newspaper article for "
            + str(cnt)
            + " games at "
            + date
            + " sent to "
            + self.mailAddrNewspaper
        )
        return cnt

    def send_Notifications(self, date):
        """Send notifications to game judges via SMS or E-Mail"""
        cnt = 0
        noteTable = self.gameTable[
            (self.gameTable[self._colDate] == date)
            & (self.gameTable[self._colGuest] != "spielfrei")
        ].dropna(how="all")
        for game in noteTable[self._colNr]:
            ak = noteTable.loc[noteTable[self._colNr] == game, self._colAK].values[0]
            team = noteTable.loc[noteTable[self._colNr] == game, self._colJTeam].values[
                0
            ]
            mv = noteTable.loc[noteTable[self._colNr] == game, self._colJMV].values[0]
            mailMV = noteTable.loc[
                noteTable[self._colNr] == game, self._colMailJMV
            ].values[0]
            receivers = []
            receivers.append(
                {
                    "name": noteTable.loc[
                        noteTable[self._colNr] == game, self._colJudge1
                    ].values[0],
                    "mail": noteTable.loc[
                        noteTable[self._colNr] == game, self._colMailJudge1
                    ].values[0],
                    "task": self._colJudge1,
                }
            )
            receivers.append(
                {
                    "name": noteTable.loc[
                        noteTable[self._colNr] == game, self._colJudge2
                    ].values[0],
                    "mail": noteTable.loc[
                        noteTable[self._colNr] == game, self._colMailJudge2
                    ].values[0],
                    "task": self._colJudge2,
                }
            )
            receivers.append(
                {
                    "name": noteTable.loc[
                        noteTable[self._colNr] == game, self._colShop1
                    ].values[0],
                    "mail": noteTable.loc[
                        noteTable[self._colNr] == game, self._colMailShop1
                    ].values[0],
                    "task": self._colShop1,
                }
            )
            receivers.append(
                {
                    "name": noteTable.loc[
                        noteTable[self._colNr] == game, self._colShop2
                    ].values[0],
                    "mail": noteTable.loc[
                        noteTable[self._colNr] == game, self._colMailShop2
                    ].values[0],
                    "task": self._colShop2,
                }
            )
            receivers.append(
                {
                    "name": noteTable.loc[
                        noteTable[self._colNr] == game, self._colSecurity
                    ].values[0],
                    "mail": noteTable.loc[
                        noteTable[self._colNr] == game, self._colMailSecurity
                    ].values[0],
                    "task": self._colSecurity,
                }
            )
            receivers.append(
                {
                    "name": noteTable.loc[
                        noteTable[self._colNr] == game, self._colCleaning
                    ].values[0],
                    "mail": noteTable.loc[
                        noteTable[self._colNr] == game, self._colMailCleaning
                    ].values[0],
                    "task": self._colCleaning,
                }
            )
            home = noteTable.loc[noteTable[self._colNr] == game, self._colHome].values[
                0
            ]
            guest = noteTable.loc[
                noteTable[self._colNr] == game, self._colGuest
            ].values[0]
            strTime = noteTable.loc[
                noteTable[self._colNr] == game, self._colTime
            ].values[0]
            # Notifications for jobs
            for receiver in receivers:
                if type(receiver["mail"]) == str:
                    if "@" in receiver["mail"]:
                        # send Email
                        msg = EmailMessage()
                        msg["From"] = formataddr((self.mail_name, self.mail_ID))
                        msg["Subject"] = "Benachrichtigung Dienst " + receiver["task"]
                        msg["To"] = formataddr((receiver["name"], receiver["mail"]))
                        # Message body created from mail text stored in config
                        msg.set_content(
                            self.mailTask.format(
                                receiver["name"],
                                date,
                                receiver["task"],
                                ak,
                                home,
                                guest,
                                strTime,
                            )
                        )
                        self.send_Mail(msg, self.mail_ID, self.mail_password)
                        logging.info(
                            "E-Mail sent to "
                            + receiver["name"]
                            + ", "
                            + receiver["task"]
                            + ", "
                            + str(receiver["mail"])
                        )
                        cnt +=1
                    elif "+" in receiver["mail"]:
                        # send SMS
                        # Message text created from text stored in config
                        text = self.textTask.format(
                            receiver["name"], date, receiver["task"], ak, strTime
                        )
                        self.send_SMS(receiver["mail"], text)
                        logging.info(
                            "SMS sent to "
                            + receiver["name"]
                            + ", "
                            + receiver["task"]
                            + ", "
                            + str(receiver["mail"])
                        )
                        cnt +=1
                    else:
                        logging.warning(
                            "No valid phone number or email adress available at game "
                            + str(game)
                        )
            # Notification to MV
            if type(mailMV) == str:
                if "@" in mailMV:
                    # send Email
                    msg = EmailMessage()
                    msg["From"] = formataddr((self.mail_name, self.mail_ID))
                    msg["Subject"] = self.mailMVSubject
                    msg["To"] = formataddr((mv, mailMV))
                    # Message body created from mail text stored in config
                    msg.set_content(
                        self.mailMV.format(
                            mv,
                            team,
                            date,
                            receivers[0]["name"],
                            receivers[1]["name"],
                            ak,
                            home,
                            guest,
                            strTime,
                        )
                    )
                    self.send_Mail(msg, self.mail_ID, self.mail_password)
                    logging.info("E-Mail sent to " + mv + ", MV, " + str(mailMV))
                    cnt +=1
                elif "+" in mailMV:
                    # Message text created from text stored in config
                    text = self.textMV.format(
                        mv,
                        team,
                        date,
                        receivers[0]["name"],
                        receivers[1]["name"],
                        ak,
                        strTime,
                    )
                    self.send_SMS(mailMV, text)
                    logging.info("SMS sent to " + mv + ", MV, " + str(mailMV))
                    cnt +=1
                else:
                    logging.warning(
                        "No valid phone number or email adress available at game "
                        + str(game)
                    )
        return cnt

    def send_RefNotification(self, game, date, time):
        """Send referee notification to referee coordinator"""
        cnt = 0
        ak = self.onlineTable.loc[self.onlineTable[self._colNr] == game, self._colAK].values[0]

        # Send notifications to all specified receivers
        for receiver in self.mailRefCoordTargets:
            text = self.mailRefCoord.format(
                receiver["Name"],
                ak,
                date,
                time,
                ", ".join(rec["Name"] for rec in self.mailRefCoordTargets)
            )
            if type(receiver["Address"]) == str:
                if "@" in receiver["Address"]:
                    # send Email
                    msg = EmailMessage()
                    msg["From"] = formataddr((self.mail_name, self.mail_ID))
                    msg["Subject"] = self.mailRefCoordSubject
                    msg["To"] = formataddr((receiver["name"], receiver["Address"]))
                    # Message body created from mail text stored in config
                    msg.set_content(text)
                    self.send_Mail(msg, self.mail_ID, self.mail_password)
                    logging.info(
                        "E-Mail sent to "
                        + str(receiver["Name"])
                        + ", "
                        + str(receiver["Address"])
                    )
                    cnt += 1
                elif "+" in receiver["Address"]:
                    # send SMS
                    # Message text created from text stored in config
                    self.send_SMS(receiver["Address"], text)
                    logging.info(
                        "SMS sent to "
                        + str(receiver["Name"])
                        + ", "
                        + str(receiver["Address"])
                    )
                    cnt += 1
                else:
                    logging.warning(
                        "No valid phone number or email address available for "
                        + str(receiver["Name"])
                    )
        return cnt

    def send_ServiceNotifications(self, date):
        """Send early notifications to service via SMS or E-Mail"""
        cnt = 0
        noteTable = self.gameTable[
            (self.gameTable[self._colDate] == date)
            & (self.gameTable[self._colGuest] != "spielfrei")
        ].dropna(how="all")
        # Only first game is relevant
        game = noteTable[self._colNr].values[0]
        ak = noteTable.loc[noteTable[self._colNr] == game, self._colAK].values[0]
        receivers = []
        receivers.append(
            {
                "name": noteTable.loc[
                    noteTable[self._colNr] == game, self._colShop1
                ].values[0],
                "mail": noteTable.loc[
                    noteTable[self._colNr] == game, self._colMailShop1
                ].values[0],
                "task": self._colShop1,
                "partner": noteTable.loc[
                    noteTable[self._colNr] == game, self._colShop2
                ].values[0],
            }
        )
        receivers.append(
            {
                "name": noteTable.loc[
                    noteTable[self._colNr] == game, self._colShop2
                ].values[0],
                "mail": noteTable.loc[
                    noteTable[self._colNr] == game, self._colMailShop2
                ].values[0],
                "task": self._colShop2,
                "partner": noteTable.loc[
                    noteTable[self._colNr] == game, self._colShop1
                ].values[0],
            }
        )
        home = noteTable.loc[noteTable[self._colNr] == game, self._colHome].values[0]
        guest = noteTable.loc[noteTable[self._colNr] == game, self._colGuest].values[0]
        strTime = noteTable.loc[noteTable[self._colNr] == game, self._colTime].values[0]
        # Early notification for service
        for receiver in receivers:
            if type(receiver["mail"]) == str:
                if "@" in receiver["mail"]:
                    # send Email
                    msg = EmailMessage()
                    msg["From"] = formataddr((self.mail_saleName, self.mail_saleID))
                    msg["Subject"] = "Vorbereitung Dienst " + receiver["task"]
                    msg["To"] = formataddr((receiver["name"], receiver["mail"]))
                    # Message body created from mail text stored in config
                    msg.set_content(
                        self.mailEarlyTask.format(
                            receiver["name"],
                            date,
                            receiver["task"],
                            ak,
                            home,
                            guest,
                            receiver["partner"],
                            strTime,
                            receiver["partner"],
                        )
                    )
                    self.send_Mail(msg, self.mail_saleID, self.mail_salePassword)
                    logging.info(
                        "E-Mail sent to "
                        + receiver["name"]
                        + ", "
                        + receiver["task"]
                        + ", "
                        + str(receiver["mail"])
                    )
                    cnt +=1
                elif "+" in receiver["mail"]:
                    # send SMS
                    # Message text created from text stored in config
                    text = self.textEarlyTask.format(
                        receiver["name"],
                        date,
                        receiver["task"],
                        ak,
                        receiver["partner"],
                        strTime,
                        receiver["partner"],
                    )
                    self.send_SMS(receiver["mail"], text)
                    logging.info(
                        "SMS sent to "
                        + receiver["name"]
                        + ", "
                        + receiver["task"]
                        + ", "
                        + str(receiver["mail"])
                    )
                    cnt +=1
                else:
                    logging.warning(
                        "No valid phone number or email adress available at game "
                        + str(game)
                    )
        return cnt

    def send_PreNotifications(self, date):
        """Send early notifications to game judges via SMS or E-Mail"""
        cnt = 0
        noteTable = self.gameTable[
            (self.gameTable[self._colDate] == date)
            & (self.gameTable[self._colGuest] != "spielfrei")
        ].dropna(how="all")
        for gameNr, game in enumerate(noteTable[self._colNr]):
            ak = noteTable.loc[noteTable[self._colNr] == game, self._colAK].values[0]
            receivers = []
            receivers.append(
                {
                    "name": noteTable.loc[
                        noteTable[self._colNr] == game, self._colJudge1
                    ].values[0],
                    "mail": noteTable.loc[
                        noteTable[self._colNr] == game, self._colMailJudge1
                    ].values[0],
                    "task": self._colJudge1,
                }
            )
            receivers.append(
                {
                    "name": noteTable.loc[
                        noteTable[self._colNr] == game, self._colJudge2
                    ].values[0],
                    "mail": noteTable.loc[
                        noteTable[self._colNr] == game, self._colMailJudge2
                    ].values[0],
                    "task": self._colJudge2,
                }
            )
            # exclude first shop tasks, are notified separately
            if gameNr > 0:
                receivers.append(
                    {
                        "name": noteTable.loc[
                            noteTable[self._colNr] == game, self._colShop1
                        ].values[0],
                        "mail": noteTable.loc[
                            noteTable[self._colNr] == game, self._colMailShop1
                        ].values[0],
                        "task": self._colShop1,
                    }
                )
                receivers.append(
                    {
                        "name": noteTable.loc[
                            noteTable[self._colNr] == game, self._colShop2
                        ].values[0],
                        "mail": noteTable.loc[
                            noteTable[self._colNr] == game, self._colMailShop2
                        ].values[0],
                        "task": self._colShop2,
                    }
                )
            receivers.append(
                {
                    "name": noteTable.loc[
                        noteTable[self._colNr] == game, self._colSecurity
                    ].values[0],
                    "mail": noteTable.loc[
                        noteTable[self._colNr] == game, self._colMailSecurity
                    ].values[0],
                    "task": self._colSecurity,
                }
            )
            receivers.append(
                {
                    "name": noteTable.loc[
                        noteTable[self._colNr] == game, self._colCleaning
                    ].values[0],
                    "mail": noteTable.loc[
                        noteTable[self._colNr] == game, self._colMailCleaning
                    ].values[0],
                    "task": self._colCleaning,
                }
            )
            home = noteTable.loc[noteTable[self._colNr] == game, self._colHome].values[
                0
            ]
            guest = noteTable.loc[
                noteTable[self._colNr] == game, self._colGuest
            ].values[0]
            strTime = noteTable.loc[
                noteTable[self._colNr] == game, self._colTime
            ].values[0]
            # Notifications for jobs
            for receiver in receivers:
                if type(receiver["mail"]) == str:
                    if "@" in receiver["mail"]:
                        # send Email
                        msg = EmailMessage()
                        msg["From"] = formataddr((self.mail_name, self.mail_ID))
                        msg["Subject"] = "Benachrichtigung Dienst " + receiver["task"]
                        msg["To"] = formataddr((receiver["name"], receiver["mail"]))
                        # Message body created from mail text stored in config
                        msg.set_content(
                            self.mailPreTask.format(
                                receiver["name"],
                                date,
                                receiver["task"],
                                ak,
                                home,
                                guest,
                                strTime,
                            )
                        )
                        self.send_Mail(msg, self.mail_ID, self.mail_password)
                        logging.info(
                            "E-Mail sent to "
                            + receiver["name"]
                            + ", "
                            + receiver["task"]
                            + ", "
                            + str(receiver["mail"])
                        )
                        cnt +=1
                    elif "+" in receiver["mail"]:
                        # send SMS
                        # Message text created from text stored in config
                        text = self.textPreTask.format(
                            receiver["name"], date, receiver["task"], ak, strTime
                        )
                        self.send_SMS(receiver["mail"], text)
                        logging.info(
                            "SMS sent to "
                            + receiver["name"]
                            + ", "
                            + receiver["task"]
                            + ", "
                            + str(receiver["mail"])
                        )
                        cnt +=1
                    else:
                        logging.warning(
                            "No valid phone number or email adress available at game "
                            + str(game)
                        )
        return cnt

    def send_ShfitNotification(self, game, oldDate, oldTime, newDate, newTime):
        """Send notifications on shifted datum to game judges via SMS or E-Mail"""
        cnt = 0
        noteTable = self.gameTable[
            (self.gameTable[self._colGuest] != "spielfrei")
        ].dropna(how="all")
        ak = noteTable.loc[noteTable[self._colNr] == game, self._colAK].values[0]
        receivers = []
        receivers.append(
            {
                "name": noteTable.loc[
                    noteTable[self._colNr] == game, self._colJudge1
                ].values[0],
                "mail": noteTable.loc[
                    noteTable[self._colNr] == game, self._colMailJudge1
                ].values[0],
                "task": self._colJudge1,
            }
        )
        receivers.append(
            {
                "name": noteTable.loc[
                    noteTable[self._colNr] == game, self._colJudge2
                ].values[0],
                "mail": noteTable.loc[
                    noteTable[self._colNr] == game, self._colMailJudge2
                ].values[0],
                "task": self._colJudge2,
            }
        )
        receivers.append(
            {
                "name": noteTable.loc[
                    noteTable[self._colNr] == game, self._colShop1
                ].values[0],
                "mail": noteTable.loc[
                    noteTable[self._colNr] == game, self._colMailShop1
                ].values[0],
                "task": self._colShop1,
            }
        )
        receivers.append(
            {
                "name": noteTable.loc[
                    noteTable[self._colNr] == game, self._colShop2
                ].values[0],
                "mail": noteTable.loc[
                    noteTable[self._colNr] == game, self._colMailShop2
                ].values[0],
                "task": self._colShop2,
            }
        )
        receivers.append(
            {
                "name": noteTable.loc[
                    noteTable[self._colNr] == game, self._colSecurity
                ].values[0],
                "mail": noteTable.loc[
                    noteTable[self._colNr] == game, self._colMailSecurity
                ].values[0],
                "task": self._colSecurity,
            }
        )
        receivers.append(
            {
                "name": noteTable.loc[
                    noteTable[self._colNr] == game, self._colCleaning
                ].values[0],
                "mail": noteTable.loc[
                    noteTable[self._colNr] == game, self._colMailCleaning
                ].values[0],
                "task": self._colCleaning,
            }
        )
        home = noteTable.loc[noteTable[self._colNr] == game, self._colHome].values[0]
        guest = noteTable.loc[noteTable[self._colNr] == game, self._colGuest].values[0]
        # Notifications for jobs
        for receiver in receivers:
            if type(receiver["mail"]) == str:
                if "@" in receiver["mail"]:
                    # send Email
                    msg = EmailMessage()
                    msg["From"] = formataddr((self.mail_name, self.mail_ID))
                    msg["Subject"] = (
                        "Benachrichtigung Verschiebung Dienst " + receiver["task"]
                    )
                    msg["To"] = formataddr((receiver["name"], receiver["mail"]))
                    # Message body created from mail text stored in config
                    msg.set_content(
                        self.mailShifted.format(
                            receiver["name"],
                            receiver["task"],
                            ak,
                            home,
                            guest,
                            oldDate,
                            oldTime,
                            newDate,
                            newTime,
                        )
                    )
                    self.send_Mail(msg, self.mail_ID, self.mail_password)
                    logging.info(
                        "E-Mail sent to "
                        + receiver["name"]
                        + ", "
                        + receiver["task"]
                        + ", "
                        + str(receiver["mail"])
                    )
                    cnt +=1
                elif "+" in receiver["mail"]:
                    # send SMS
                    # Message text created from text stored in config
                    text = self.textShifted.format(
                        receiver["name"],
                        receiver["task"],
                        oldDate,
                        oldTime,
                        newDate,
                        newTime,
                    )
                    self.send_SMS(receiver["mail"], text)
                    logging.info(
                        "SMS sent to "
                        + receiver["name"]
                        + ", "
                        + receiver["task"]
                        + ", "
                        + str(receiver["mail"])
                    )
                    cnt +=1
                else:
                    logging.warning(
                        "No valid phone number or email adress available at game "
                        + str(game)
                    )
        return cnt


#  main program
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
    # Check if newspaper article has to be send
    gameDateSa      = nlh.get_today() + datetime.timedelta(days=9)
    strGameDateSa   = gameDateSa.strftime("%d.%m.%Y")
    strGameDaySa    = gameDateSa.strftime("%A")
    gameDateSo      = nlh.get_today() + datetime.timedelta(days=10)
    strGameDateSo   = gameDateSo.strftime("%d.%m.%Y")
    strGameDaySo    = gameDateSo.strftime("%A")

    # Send newspaper article for Saturday
    if nlh.gameTable[nlh._colDate].str.contains(strGameDateSa).any() & (strGameDaySa == "Saturday"):
        articleDate     = gameDateSa + datetime.timedelta(days=-1)
        strGameDate     = strGameDateSa
        strGameDay      = "Samstag"
        strArticleDate  = articleDate.strftime("%d.%m.%Y")
        cnt             = nlh.send_Article(strGameDate, strGameDay, strArticleDate)

    # Send newspaper article for Sunday
    elif nlh.gameTable[nlh._colDate].str.contains(strGameDateSo).any() & (strGameDaySo == "Sunday"):
        articleDate     = gameDateSo + datetime.timedelta(days=-2)
        strGameDate     = strGameDateSo
        strGameDay      = "Sonntag"
        strArticleDate  = articleDate.strftime("%d.%m.%Y")
        cnt             = nlh.send_Article(strGameDate, strGameDay, strArticleDate)
    """

    # Check if judge notifications have to be send
    tomorrow = nlh.get_today() + datetime.timedelta(days=1)
    strTomorrow = tomorrow.strftime("%d.%m.%Y")
    if nlh.gameTable[nlh._colDate].str.contains(strTomorrow).any():
        cnt = 0
        cnt = nlh.send_Notifications(strTomorrow)
        logging.info("Number of sent service notifications: " + str(cnt))
        logging.info("-------------------------------------------------")

    # Check if referee notifications have to be send
    if not nlh.gameTable[
        nlh.gameTable[nlh._colDate].str.contains(strTomorrow)
        & nlh.gameTable[nlh._colScore].str.contains("Heim")
    ].empty:
        cnt = 0
        cnt = nlh.send_RefNotification(strTomorrow)
        logging.info("Number of required home referees: " + str(cnt))
        logging.info("-------------------------------------------------")

    # Check if early catering notifications have to be send
    nextWeek = nlh.get_today() + datetime.timedelta(days=7)
    strNextWeek = nextWeek.strftime("%d.%m.%Y")
    if nlh.gameTable[nlh._colDate].str.contains(strNextWeek).any():
        cnt = 0
        cnt = nlh.send_ServiceNotifications(strNextWeek)
        cnt = nlh.send_PreNotifications(strNextWeek)
        logging.info("Number of sent service notifications: " + str(cnt))
        logging.info("-------------------------------------------------")

    logging.info("nuLiga Helper finished")
    logging.info("#################################################")
