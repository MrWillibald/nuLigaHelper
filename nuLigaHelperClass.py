# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# A python tool in planning home games for Handball clubs
# Functions:
# - Read home game plan for sports stadium from nuLiga
# - Update game plan Excel document with judge scheduling
# - Send notifications to game judges
# - Send notifications to team leaders
# - Send notification to referee planner
# - Send newspaper article to local newspaper
# ---------------------------------------------------------------
# Created by: MrWillibald
# Version 0.1
# Date: 16.04.2018
# ---------------------------------------------------------------
# Created by: MrWillibald
# Version 0.2
# Info: Added JudgeMV, Added SMS capabilities
# Date: 25.09.2018
# ---------------------------------------------------------------
# Created by: MrWillibald
# Version 0.2.1
# Info: Fixed issue with 'GE' games, added autofilter
#       Watch emails and SMS with empty judges!
# Date: 10.12.2018
# ---------------------------------------------------------------
# Created by: MrWillibald
# Version 0.2.2
# Info: Only tournament information for MI and GE
# Date: 11.03.2019
# ---------------------------------------------------------------
# Created by: MrWillibald
# Version 0.3
# Info: Restructured to object oriented class,
#       Added home referee notification
# Date: 12.03.2019
# ---------------------------------------------------------------
# Created by: MrWillibald
# Version 0.4
# Info: Moved config to external json file, added logger
# Date: 23.07.2019
# ---------------------------------------------------------------
# Created by: MrWillibald
# Version 0.5
# Info: Added columns for Kuchen/Verkauf job
# Date: 07.09.2019
# ---------------------------------------------------------------
# Created by: MrWillibald
# Version 0.6
# Info: Moved message strings to external config file
# Date: 10.10.2019
# ---------------------------------------------------------------
# Created by: MrWillibald
# Version 0.6.1
# Info: Fixed string conversion in warnings
# Date: 08.11.2019
# ---------------------------------------------------------------
# Created by: MrWillibald
# Version 0.7
# Info: Reworked home referee notifications with config file
# Date: 28.02.2020
# ---------------------------------------------------------------

import requests
import time
import datetime
import pandas as pd
import numpy as np
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import dropbox
import os
from twilio.rest import Client
import json
import logging

# Version string
version = '0.7'
# Debug flag
debug = False


class nuLigaHomeGames:
    """nuLigaHelper class for home game planning and scheduled job notifications"""

    def set_today(self, today):
        """Set todays date. Intended for debugging."""
        self._today = today
        if self._today.month >= 8:
            self._year = self._today.year
        else:
            self._year = today.year - 1
        self.__dictSeason = {'part1': str(self._year),
                             'part2': str(self._year + 1),
                             'part3': str(self._year + 1)[-2:]}
        self.__season = '{part1}%2F{part3}'.format(**self.__dictSeason)
        self.file = 'Heimspielplan_{part1}_{part3}.xlsx'.format(**self.__dictSeason)

    def get_today(self):
        """Return todays date"""
        return self._today

    def __init__(self, *args, **kwargs):
        """Initialize class instance with necessary strings"""
        self.__version = version

        logging.info("Initialization of nuLiga Helper")
        logging.info("Version: " + self.__version)

        # Set up dates and strings
        self.set_today(datetime.date.today())
        #self.set_today(datetime.date(2019, 10, 3))

        # New config workflow
        with open(os.path.join(os.path.dirname(__file__), 'config.json'),
                  encoding='utf-8') as json_config_file:
            config = json.load(json_config_file)

        # Club properties
        self.__dict__.update(config['club']['info'])
        # Email account properties
        self.__dict__.update(config['club']['email'])
        # Dropbox account properties
        self.__dict__.update(config['club']['dropbox'])
        self.dbc = dropbox.Dropbox(self.dropbox_token)
        # Twilio account properties
        self.__dict__.update(config['club']['twilio'])
        # Column names
        self.__dict__.update(config['club']['columns'])
        # Texts
        self.__dict__.update(config['club']['texts'])

        logging.info("Initialization completed")

    def send_Mail(self, fromaddr, toaddr, text):
        """Send E-Mail via specified SMTP server"""
        if True == debug:
            return 0
        server = smtplib.SMTP_SSL(self.smtpserver)
        # server.set_debuglevel(True)
        server.login(self.mail_ID, self.mail_password)
        server.sendmail(fromaddr, toaddr, text)
        server.quit()

    def send_SMS(self, fromaddr, toaddr, text):
        """Send SMS via specified Twilio account"""
        if True == debug:
            return  0
        client = Client(self.twilio_sid, self.twilio_token)
        message = client.messages.create(body=text, from_=fromaddr, to=toaddr)
        return message

    def get_fromDropbox(self):
        """Download file from specified Dropbox account"""
        try:
            self.dbc.files_download_to_file(
                self.file, '/' + self.dropbox_folder + '/' + self.file)
        except dropbox.exceptions.ApiError:
            logging.warning("Error while loading judge schedule from Dropbox, \
                             new schedule is created")
        logging.info("Judge schedule loaded and saved successfully from Dropbox")

    def upload_toDropbox(self):
        """Upload file to specified Dropbox account"""
        with open(self.file, 'rb') as f:
            data = f.read()
            self.dbc.files_upload(data,
                                  '/' + self.dropbox_folder + '/' + self.file,
                                  mode=dropbox.files.WriteMode.overwrite)
        try:
            rmf = os.remove(self.file)
        except OSError:
            logging.warning(rmf)
        logging.info("Judge schedule successfully uploaded to Dropbox and cleaned locally")

    def get_onlineTable(self):
        """Scrape Hallenspielan for specified sports hall from BHV website"""
        logging.info("Read current home game plan from BHV Hallenspielplan website")
        lMonths = [#'September+{part1}'.format(**self.__dictSeason),
                   'Oktober+{part1}'.format(**self.__dictSeason),
                   'November+{part1}'.format(**self.__dictSeason),
                   'Dezember+{part1}'.format(**self.__dictSeason),
                   'Januar+{part2}'.format(**self.__dictSeason),
                   'Februar+{part2}'.format(**self.__dictSeason),
                   'März+{part2}'.format(**self.__dictSeason)]#,
                   #'April+{part2}'.format(**self.__dictSeason)]
        lGames = list()
        # read all pages in the season (http requests)
        for month in lMonths:
            # read contents from gym plan
            page = requests.get('https://bhv-handball.liga.nu/cgi-bin/WebObjects/nuLigaHBDE.woa/wa/\
                                courtInfo?month={}&federation=BHV&championship=OB+{}&location={}'
                                .format(month, self.__season, self.location))
            table = pd.read_html(page.content, header=0, attrs={"class": "result-set"})
            table[0].drop(table[0].columns[[9, 10, 11]], axis=1, inplace=True)
            table[0].columns = ([self._colDay,
                                 self._colDate,
                                 self._colTime,
                                 self._colNr,
                                 self._colAK,
                                 self._colStaffel,
                                 self._colHome,
                                 self._colGuest,
                                 self._colScore])
            lGames.append(table[0])
        # modify existing columns
        self.onlineTable                                = pd.concat(lGames)
        self.onlineTable.index                          = range(len(self.onlineTable[self._colDay]))
        self.onlineTable[[self._colDay, self._colDate]] = self.onlineTable[[self._colDay, self._colDate]].fillna(method='ffill')
        self.onlineTable[self._colScore]                = self.onlineTable[self._colScore].astype(str) + "\t"
        # add additional columns
        kwargs              = {self._colJTeam: np.empty(len(self.onlineTable[self._colDay]), dtype=str)}
        self.onlineTable    = self.onlineTable.assign(**kwargs)
        kwargs              = {self._colJMV: np.empty(len(self.onlineTable[self._colDay]), dtype=str)}
        self.onlineTable    = self.onlineTable.assign(**kwargs)
        kwargs              = {self._colMailJMV: np.empty(len(self.onlineTable[self._colDay]), dtype=str)}
        self.onlineTable    = self.onlineTable.assign(**kwargs)
        kwargs              = {self._colJudge1: np.empty(len(self.onlineTable[self._colDay]), dtype=str)}
        self.onlineTable    = self.onlineTable.assign(**kwargs)
        kwargs              = {self._colMailJudge1: np.empty(len(self.onlineTable[self._colDay]), dtype=str)}
        self.onlineTable    = self.onlineTable.assign(**kwargs)
        kwargs              = {self._colJudge2: np.empty(len(self.onlineTable[self._colDay]), dtype=str)}
        self.onlineTable    = self.onlineTable.assign(**kwargs)
        kwargs              = {self._colMailJudge2: np.empty(len(self.onlineTable[self._colDay]), dtype=str)}
        self.onlineTable    = self.onlineTable.assign(**kwargs)
        kwargs              = {self._colCake: np.empty(len(self.onlineTable[self._colDay]), dtype=str)}
        self.onlineTable    = self.onlineTable.assign(**kwargs)
        kwargs              = {self._colMailCake: np.empty(len(self.onlineTable[self._colDay]), dtype=str)}
        self.onlineTable    = self.onlineTable.assign(**kwargs)
        logging.info("Current home game plan loaded")

    def get_gameTable(self):
        """Read scheduled jobs from downloaded or local file"""
        logging.info("Read local judge schedule")
        try:
            self.gameTable = pd.read_excel(self.file, dtype={self._colMailJMV: str, 
                                                             self._colMailJudge1: str, 
                                                             self._colMailJudge2: str, 
                                                             self._colMailCake: str})
            logging.info("Judge schedule available")
        except OSError:
            self.gameTable = self.onlineTable
            logging.warning("Judge schedule not available, empty schedule created")
        return 0

    def merge_tables(self):
        """Merge up-to-date Hallenspielplan with scheduled jobs file"""
        logging.info("Update home game plan")
        for game in self.onlineTable[self._colNr]:
            try:
                judges = self.gameTable.loc[self.gameTable[self._colNr] == game, 
                                               [self._colJTeam, 
                                                self._colJMV, 
                                                self._colMailJMV, 
                                                self._colJudge1, 
                                                self._colMailJudge1, 
                                                self._colJudge2, 
                                                self._colMailJudge2, 
                                                self._colCake, 
                                                self._colMailCake]]
                self.onlineTable.loc[self.onlineTable[self._colNr] == game, 
                   [self._colJTeam, 
                    self._colJMV, 
                    self._colMailJMV, 
                    self._colJudge1, 
                    self._colMailJudge1, 
                    self._colJudge2, 
                    self._colMailJudge2, 
                    self._colCake, 
                    self._colMailCake]] = judges.values[0]
                logging.info("Game " + str(game) + " merged with schedule")
            except IndexError:
                # oTable = pTable
                if self.onlineTable.loc[self.onlineTable[self._colNr] == game, [self._colAK]].values[0] != "GE":
                    # send Error Notification
                    fromaddr        = self.mail_ID
                    toaddr          = fromaddr
                    msg             = MIMEMultipart()
                    msg['From']     = fromaddr
                    msg['Subject']  = self.mailErrorSubject
                    msg['To']       = toaddr
                    body            = self.mailError
                    msg.attach(MIMEText(body, 'plain'))
                    text = msg.as_string()
                    self.send_Mail(fromaddr, toaddr, text)
                    logging.warning("Spielnummer not contained in home schedule, please correct manually!")
        # make game and online table identical if merging was successful
        self.gameTable = self.onlineTable
        logging.info("Update home game plan completed")
        return 0

    def write_toXlsx(self):
        """Write updated job scheduling to local *.xlsx file"""
        writer = pd.ExcelWriter(self.file, engine='xlsxwriter')
        self.gameTable.to_excel(writer, sheet_name='Heimspielplan', encoding=self._enc)
        worksheet   = writer.sheets['Heimspielplan']
        workbook    = writer.book
        formatText  = workbook.add_format({'num_format': '@'})
        worksheet.set_column('C:C', 10)
        worksheet.set_column('G:H', 20)
        worksheet.set_column('I:I', 28)
        worksheet.set_column('J:J', 13)
        worksheet.set_column('K:K', 18)
        worksheet.set_column('L:L', 30)
        worksheet.set_column('M:M', 35, formatText)
        worksheet.set_column('N:N', 30)
        worksheet.set_column('O:O', 35, formatText)
        worksheet.set_column('P:P', 30)
        worksheet.set_column('Q:Q', 35, formatText)
        worksheet.set_column('R:R', 30)
        worksheet.set_column('S:S', 35, formatText)
        worksheet.autofilter(0, 0, self.gameTable.shape[0], self.gameTable.shape[1])
        writer.save()
        logging.info("Judge schedule saved locally")

    def send_Article(self, date, day, articleDate):
        """Send notification article to local newspaper"""
        cnt             = 0
        team            = []
        home            = []
        guest           = []
        strTime         = []
        tournamentMI    = False
        tournamentGE    = False
        noteTable       = self.gameTable[self.gameTable[self._colDate] == date].dropna(how='all')
        for game in noteTable[self._colNr]:
            teamRaw = noteTable.loc[noteTable[self._colNr] == game, self._colAK].values[0]
            if teamRaw == 'F':
                team.append('Damen')
            elif teamRaw == 'M':
                team.append('Herren')
            else:
                team.append(teamRaw)
            home.append(noteTable.loc[noteTable[self._colNr] == game, self._colHome].values[0])
            guest.append(noteTable.loc[noteTable[self._colNr] == game, self._colGuest].values[0])
            strTime.append(noteTable.loc[noteTable[self._colNr] == game, self._colTime].values[0])
        fromaddr        = self.mail_ID
        toaddr          = self.mailAddrNewspaper
        logging.info("Send newspaper article to " + toaddr)
        msg             = MIMEMultipart()
        msg['From']     = fromaddr
        msg['To']       = toaddr
        msg['Subject']  = self.mailNewspaperSubject

        # Create schedule from game plan
        schedule = ""
        for i in range(0, len(home)):
            # single tournament information for MI and GE
            if ('MI' == team[i]) and (False == tournamentMI):
                schedule        = schedule + 'Ab ' + strTime[i].strip(' v').strip(' t') + ' Spielfest der Minis\n'
                tournamentMI    = True
                cnt             = cnt + 1
            elif ('GE' == team[i]) and (False == tournamentGE):
                schedule        = schedule + 'Ab ' + strTime[i].strip(' v').strip(' t') + ' Turnier der gemischten E-Jugend\n'
                tournamentGE    = True
                cnt             = cnt + 1
            elif ('GE' or 'MI') != team[i]:
                schedule        = schedule + strTime[i].strip(' v').strip(' t') + " " + team[i] + " " + home[i] + " - " + guest[i] + "\n"
                cnt             = cnt + 1

        # Message body created from mail text stored in config
        body = self.mailNewspaper.format(articleDate, day, date, schedule)
        msg.attach(MIMEText(body, 'plain'))
        text = msg.as_string()
        self.send_Mail(fromaddr, toaddr, text)
        logging.info("Newspaper article for " + str(cnt) + " games at " + date + " sent to " + toaddr)
        return cnt

    def send_Notifications(self, date):
        """Send notifications to game judges via SMS or E-Mail"""
        cnt         = 0
        judge       = ['Du', 'Du']
        mailJudge   = ['', '']
        noteTable   = self.gameTable[self.gameTable[self._colDate] == date].dropna(how="all")
        for game in noteTable[self._colNr]:
            AK              = noteTable.loc[noteTable[self._colNr] == game, self._colAK].values[0]
            team            = noteTable.loc[noteTable[self._colNr] == game, self._colJTeam].values[0]
            MV              = noteTable.loc[noteTable[self._colNr] == game, self._colJMV].values[0]
            mailMV          = noteTable.loc[noteTable[self._colNr] == game, self._colMailJMV].values[0]
            judge[0]        = noteTable.loc[noteTable[self._colNr] == game, self._colJudge1].values[0]
            mailJudge[0]    = noteTable.loc[noteTable[self._colNr] == game, self._colMailJudge1].values[0]
            judge[1]        = noteTable.loc[noteTable[self._colNr] == game, self._colJudge2].values[0]
            mailJudge[1]    = noteTable.loc[noteTable[self._colNr] == game, self._colMailJudge2].values[0]
            home            = noteTable.loc[noteTable[self._colNr] == game, self._colHome].values[0]
            guest           = noteTable.loc[noteTable[self._colNr] == game, self._colGuest].values[0]
            strTime         = noteTable.loc[noteTable[self._colNr] == game, self._colTime].values[0]
            # Notifications to Judges
            for i in range(0, 2):
                toaddr = mailJudge[i]
                if type(toaddr) == str:
                    if '@' in toaddr:
                        # send Email
                        fromaddr        = self.mail_ID
                        msg             = MIMEMultipart()
                        msg['From']     = fromaddr
                        msg['Subject']  = self.mailJudgeSubject
                        msg['To']       = toaddr
                        # Message body created from mail text stored in config 
                        body = self.mailJudge.format(judge[i], date, AK, home, guest, strTime)
                        msg.attach(MIMEText(body, 'plain'))
                        text = msg.as_string()
                        self.send_Mail(fromaddr, toaddr, text)
                        logging.info("E-Mail sent to " + judge[i] + ", " + str(toaddr))
                        cnt = cnt + 1
                    elif '+' in toaddr:
                        # send SMS
                        fromaddr = self.twilio_ID
                        # Message text created from text stored in config
                        text = self.textJudge.format(judge[i], date, AK, strTime)
                        self.send_SMS(fromaddr, toaddr, text)
                        logging.info("SMS sent to " + judge[i] + ", " + str(toaddr))
                        cnt = cnt + 1
                    else:
                        logging.warning("No valid phone number or email adress available at game " + str(game))
            # Notification to MV
            toaddr = mailMV
            if type(toaddr) == str:
                if '@' in toaddr:
                    # send Email
                    fromaddr        = self.mail_ID
                    msg             = MIMEMultipart()
                    msg['From']     = fromaddr
                    msg['Subject']  = self.mailMVSubject
                    msg['To']       = toaddr
                    # Message body created from mail text stored in config 
                    body = self.mailMV.format(MV, team, date, judge[0], judge[1], AK, home, guest, strTime)
                    msg.attach(MIMEText(body, 'plain'))
                    text = msg.as_string()
                    self.send_Mail(fromaddr, toaddr, text)
                    logging.info("E-Mail sent to " + MV + ", " + str(toaddr))
                    cnt = cnt + 1
                elif '+' in toaddr:
                    # send SMS
                    fromaddr = self.twilio_ID
                    # Message text created from text stored in config
                    text = self.textMV.format(MV, team, date, judge[0], judge[1], AK, strTime)
                    self.send_SMS(fromaddr, toaddr, text)
                    logging.info("SMS sent to " + MV + ", " + str(toaddr))
                    cnt = cnt + 1
                else:
                    logging.warning("No valid phone number or email address available at game " + str(game))
        return cnt

    def send_RefNotification(self, date):
        """Send referee notification to referee coordinator"""
        # collect games with missing referees
        noteTable   = self.gameTable[self.gameTable[self._colDate].str.contains(strTomorrow) & self.gameTable[self._colScore].str.contains("Heim")]
        cnt         = 0
        textGames   = ""
        for game in noteTable[self._colNr]:
            AK          = noteTable.loc[noteTable[self._colNr] == game, self._colAK].values[0]
            strTime     = noteTable.loc[noteTable[self._colNr] == game, self._colTime].values[0]
            textGames   = textGames + AK + " um " + strTime + "\n"
            cnt         = cnt + 1

        # Send notifications to all specified receivers
        for receiver in self.mailRefCoordTargets:
            toaddr = receiver["Address"]
            if type(toaddr) == str:
                if '@' in toaddr:
                    # send Email
                    fromaddr        = self.mail_ID
                    msg             = MIMEMultipart()
                    msg['From']     = fromaddr
                    msg['Subject']  = self.mailRefCoordSubject
                    msg['To']       = toaddr
                    # Message body created from mail text stored in config 
                    body = self.mailRefCoord.format(receiver["Name"], date, textGames, ', '.join(rec["Name"] for rec in self.mailRefCoordTargets))
                    msg.attach(MIMEText(body, 'plain'))
                    text = msg.as_string()
                    self.send_Mail(fromaddr, toaddr, text)
                    logging.info("E-Mail sent to " + str(receiver["Name"]) + ", " + str(toaddr))
                elif '+' in toaddr:
                    # send SMS
                    fromaddr = self.twilio_ID
                    # Message text created from text stored in config
                    text = self.mailRefCoord.format(receiver["Name"], date, textGames, ', '.join(rec["Name"] for rec in self.mailRefCoordTargets))
                    self.send_SMS(fromaddr, toaddr, text)
                    logging.info("SMS sent to " + str(receiver["Name"]) + ", " + str(toaddr))
                else:
                    logging.warning("No valid phone number or email address available for " + str(receiver["Name"]))

        return cnt


#  main program
if __name__ == '__main__':

    # Initialize logger
    logging.basicConfig(format='%(asctime)s - %(levelname)s - %(message)s', filename='helper.log',level=logging.DEBUG)
    logging.getLogger().addHandler(logging.StreamHandler())

    # Initialize class
    self = nuLigaHomeGames()

    # Download Heimspielplan from Dropbox
    self.get_fromDropbox()

    # Check nuLiga Hallenplan and update Heimspielplan
    self.get_onlineTable()
    self.get_gameTable()
    self.merge_tables()

    # Save Heimspielplan as Excel-File
    self.write_toXlsx()

    # Upload Heimspielplan to Dropbox
    self.upload_toDropbox()

    # Check if newspaper article has to be send
    gameDateSa      = self.get_today() + datetime.timedelta(days=9)
    strGameDateSa   = gameDateSa.strftime("%d.%m.%Y")
    strGameDaySa    = gameDateSa.strftime("%A")
    gameDateSo      = self.get_today() + datetime.timedelta(days=10)
    strGameDateSo   = gameDateSo.strftime("%d.%m.%Y")
    strGameDaySo    = gameDateSo.strftime("%A")

    # Send newspaper article for Saturday
    if self.gameTable[self._colDate].str.contains(strGameDateSa).any() & (strGameDaySa == "Saturday"):
        articleDate     = gameDateSa + datetime.timedelta(days=-1)
        strGameDate     = strGameDateSa
        strGameDay      = "Samstag"
        strArticleDate  = articleDate.strftime("%d.%m.%Y")
        cnt             = self.send_Article(strGameDate, strGameDay, strArticleDate)

    # Send newspaper article for Sunday
    elif self.gameTable[self._colDate].str.contains(strGameDateSo).any() & (strGameDaySo == "Sunday"):
        articleDate     = gameDateSo + datetime.timedelta(days=-2)
        strGameDate     = strGameDateSo
        strGameDay      = "Sonntag"
        strArticleDate  = articleDate.strftime("%d.%m.%Y")
        cnt             = self.send_Article(strGameDate, strGameDay, strArticleDate)

    # Check if judge notifications have to be send
    tomorrow    = self.get_today() + datetime.timedelta(days=1)
    strTomorrow = tomorrow.strftime("%d.%m.%Y")
    if self.gameTable[self._colDate].str.contains(strTomorrow).any():
        cnt = 0
        cnt = self.send_Notifications(strTomorrow)
        logging.info("Number of sent judge notifications: " + str(cnt))

    # Check if referee notifications have to be send
    if not self.gameTable[self.gameTable[self._colDate].str.contains(strTomorrow) & self.gameTable[self._colScore].str.contains("Heim")].empty:
        cnt = 0
        cnt = self.send_RefNotification(strTomorrow)
        logging.info("Number of required home referees: " + str(cnt))

    logging.info("nuLiga Helper finished")
