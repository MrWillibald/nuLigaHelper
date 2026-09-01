# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Notification dispatch: e-mail and SMS based on database content
# ---------------------------------------------------------------

import logging
import smtplib
from email.utils import formataddr
from email.message import EmailMessage

from twilio.rest import Client

from common import DEBUG_FLAG
import db


class Notifier:
    """Sends game-related notifications via e-mail or SMS."""

    def __init__(self, config: dict, session, season_year: int):
        self._season_year = season_year

        email_cfg = config["email"]
        self.smtpserver = email_cfg["smtpserver"]
        self.mail_ID = email_cfg["mail_ID"]
        self.mail_password = email_cfg["mail_password"]
        self.mail_name = email_cfg.get("mail_name", "")
        self.mailAddrNewspaper = email_cfg["mailAddrNewspaper"]
        self.mail_saleID = email_cfg.get("mail_saleID", self.mail_ID)
        self.mail_salePassword = email_cfg.get("mail_salePassword", self.mail_password)
        self.mail_error_recipient = email_cfg.get("mailAddrAdmin", self.mail_ID)

        self.__dict__.update(config["twilio"])
        self.__dict__.update(config["texts"])

        self.session = session

    # ---------------------------------------------------------------------------
    # Low-level sending
    # ---------------------------------------------------------------------------

    def send_Mail(self, msg: EmailMessage, ID: str, password: str):
        """Send e-mail via specified SMTP server."""
        if DEBUG_FLAG:
            return None
        with smtplib.SMTP_SSL(self.smtpserver) as server:
            server.login(ID, password)
            server.send_message(msg)
        return None

    def send_SMS(self, toaddr: str, text: str):
        """Send SMS via specified Twilio account."""
        if DEBUG_FLAG:
            return None
        client = Client(self.twilio_sid, self.twilio_token)
        message = client.messages.create(
            messaging_service_sid=self.twilio_service_ID, body=text, to=toaddr
        )
        return message

    def _dispatch(
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
        Send an e-mail or SMS to a single receiver depending on their contact data.
        Prefers e-mail; falls back to SMS. Returns 1 if a message was sent.
        """
        mail_id = mail_id or self.mail_ID
        mail_password = mail_password or self.mail_password
        contact_mail = receiver.get("email")
        contact_phone = receiver.get("phone")

        if isinstance(contact_mail, str) and "@" in contact_mail:
            msg = EmailMessage()
            msg["From"] = formataddr((self.mail_name, mail_id))
            msg["Subject"] = subject
            msg["To"] = formataddr((receiver["name"], contact_mail))
            msg.set_content(mail_body)
            self.send_Mail(msg, mail_id, mail_password)
            logging.info(
                f"E-Mail sent to {receiver['name']}, "
                f"{receiver.get('task', '')}, {contact_mail}"
            )
            return 1

        if isinstance(contact_phone, str) and "+" in contact_phone:
            self.send_SMS(contact_phone, sms_body)
            logging.info(
                f"SMS sent to {receiver['name']}, "
                f"{receiver.get('task', '')}, {contact_phone}"
            )
            return 1

        logging.warning(
            f"No valid phone number or email address available at game {game_nr} "
            f"for {receiver['name']}, {receiver.get('task', '')}"
        )
        return 0

    @staticmethod
    def _person_receiver(person: db.Person, task: str) -> dict:
        """Build a receiver dict from a Person instance."""
        return {"name": person.name, "email": person.email, "phone": person.phone, "task": task}

    def send_account_message(
        self,
        person: db.Person,
        subject: str,
        mail_body: str,
        sms_body: str,
    ) -> int:
        """Send an account message using the established channel preference."""
        return self._dispatch(
            self._person_receiver(person, "Anmeldung"),
            subject,
            mail_body,
            sms_body,
            game_nr=0,
        )

    def send_account_message_via(
        self, person: db.Person, channel: str, subject: str, body: str
    ) -> int:
        """Send an account message only through the explicitly selected channel."""
        receiver = self._person_receiver(person, "Anmeldung")
        if channel == "email":
            receiver["phone"] = None
        elif channel == "sms":
            receiver["email"] = None
        else:
            logging.warning("Unknown account message channel %r", channel)
            return 0
        return self._dispatch(
            receiver, subject,
            body if channel == "email" else "",
            body if channel == "sms" else "", game_nr=0,
        )

    # ---------------------------------------------------------------------------
    # Game-day notifications (judges, shop, security, cleaning + MV)
    # ---------------------------------------------------------------------------

    def notify_game_day(self, date: str) -> int:
        """Send notifications to all scheduled helpers of games on the given date."""
        cnt = 0
        games = db.get_games_on_date(self.session, date)

        for game in games:
            cnt += self._notify_game_helpers(game, date, self.mailTask, self.textTask)

            # The MV of the responsible team is reminded as long as tasks
            # of the game are still unassigned.
            mv = game.team.mv_person if game.team is not None else None
            if mv is None or not db.missing_slots(game):
                continue
            judge_names = [
                a.person.name if a is not None else ""
                for a in (
                    game.assignment_by_role(db.ROLE_TIMEKEEPER),
                    game.assignment_by_role(db.ROLE_SECRETARY),
                )
            ]
            cnt += self._dispatch(
                self._person_receiver(mv, "MV Verantwortlich"),
                subject=self.mailMVSubject,
                mail_body=self.mailMV.format(
                    mv.name, game.judge_team_name or "", date, *judge_names,
                    game.ak, game.home, game.guest, game.time,
                ),
                sms_body=self.textMV.format(
                    mv.name, game.judge_team_name or "", date, *judge_names,
                    game.ak, game.time,
                ),
                game_nr=game.game_nr,
            )

        return cnt

    # ---------------------------------------------------------------------------
    # Early service notifications (one week ahead, first game only)
    # ---------------------------------------------------------------------------

    def notify_service_early(self, date: str) -> int:
        """Send early catering preparation notifications for the first game of the day."""
        games = db.get_games_on_date(self.session, date)
        if not games:
            return 0

        game = games[0]
        sales = game.assignments_by_role(db.ROLE_SALE)[:2]
        if len(sales) < 2:
            logging.warning(f"Less than two 'Verkauf' helpers assigned for game {game.game_nr}")

        cnt = 0
        for assignment in sales:
            partners = [a for a in sales if a is not assignment]
            partner_name = partners[0].person.name if partners else ""
            receiver = self._person_receiver(assignment.person, db.ROLE_SALE)
            cnt += self._dispatch(
                receiver,
                subject=f"Vorbereitung Dienst {db.ROLE_SALE}",
                mail_body=self.mailEarlyTask.format(
                    receiver["name"], date, db.ROLE_SALE, game.ak, game.home, game.guest,
                    partner_name, game.time, partner_name,
                ),
                sms_body=self.textEarlyTask.format(
                    receiver["name"], date, db.ROLE_SALE, game.ak,
                    partner_name, game.time, partner_name,
                ),
                game_nr=game.game_nr,
                mail_id=self.mail_saleID,
                mail_password=self.mail_salePassword,
            )
        return cnt

    # ---------------------------------------------------------------------------
    # Pre-notifications (one week ahead)
    # ---------------------------------------------------------------------------

    def notify_pre(self, date: str) -> int:
        """Send pre-notifications to game judges one week ahead."""
        cnt = 0
        games = db.get_games_on_date(self.session, date)

        # Shop roles are excluded for the first game (notified via notify_service_early)
        for idx, game in enumerate(games):
            roles = list(db.GAME_DAY_ROLES)
            if idx == 0:
                roles = [r for r in roles if r != db.ROLE_SALE]
            cnt += self._notify_game_helpers(game, date, self.mailPreTask, self.textPreTask, roles)

        return cnt

    def _notify_game_helpers(
        self, game: db.Game, date: str, mail_text: str, sms_text: str,
        roles: list[str] | None = None,
    ) -> int:
        """Send task notifications to all helpers of a single game."""
        cnt = 0
        for role in roles or db.GAME_DAY_ROLES:
            assignment = game.assignment_by_role(role)
            if assignment is None:
                continue
            receiver = self._person_receiver(assignment.person, role)
            cnt += self._dispatch(
                receiver,
                subject=f"Benachrichtigung Dienst {role}",
                mail_body=mail_text.format(
                    receiver["name"], date, role, game.ak, game.home, game.guest, game.time
                ),
                sms_body=sms_text.format(receiver["name"], date, role, game.ak, game.time),
                game_nr=game.game_nr,
            )
        return cnt

    # ---------------------------------------------------------------------------
    # Date-shift notifications
    # ---------------------------------------------------------------------------

    def notify_shifts(self, shifts: list[db.ShiftEvent]) -> int:
        """Send shift notifications for all affected games."""
        cnt = 0
        for shift in shifts:
            game = self.session.get(db.Game, shift.game_id)
            if game is None:
                continue
            logging.info(
                f"Game {shift.game_nr} is shifted! "
                f"Old date: {shift.old_date} {shift.old_time} — "
                f"New date: {shift.new_date} {shift.new_time}"
            )
            for role in db.GAME_DAY_ROLES:
                assignment = game.assignment_by_role(role)
                if assignment is None:
                    continue
                receiver = self._person_receiver(assignment.person, role)
                cnt += self._dispatch(
                    receiver,
                    subject=f"Benachrichtigung Verschiebung Dienst {role}",
                    mail_body=self.mailShifted.format(
                        receiver["name"], role, game.ak, game.home, game.guest,
                        shift.old_date, shift.old_time, shift.new_date, shift.new_time,
                    ),
                    sms_body=self.textShifted.format(
                        receiver["name"], role,
                        shift.old_date, shift.old_time, shift.new_date, shift.new_time,
                    ),
                    game_nr=game.game_nr,
                )
        return cnt

    # ---------------------------------------------------------------------------
    # Missing referee notifications
    # ---------------------------------------------------------------------------

    def notify_referee_alert(self, event: db.RefereeEvent) -> int:
        """Notify referee coordinator and MV about a missing referee for one game."""
        game = self.session.get(db.Game, event.game_id)
        if game is None:
            return 0
        return self._notify_missing_referee(game, event.date, event.time)

    def notify_referees_for_date(self, date: str) -> int:
        """Check all games of a date and notify coordinator if referees are missing."""
        cnt = 0
        for game in db.get_games_on_date(self.session, date):
            if "§77" in (game.score or ""):
                cnt += self._notify_missing_referee(game, date, game.time)
        return cnt

    def _notify_missing_referee(self, game: db.Game, date: str, time: str) -> int:
        cnt = 0
        targets = list(self.mailRefCoordTargets)
        mv = game.team.mv_person if game.team is not None else None
        if mv is not None and mv.email:
            targets.append({"Name": mv.name, "Address": mv.email})

        all_names = ", ".join(t["Name"] for t in targets)

        for target in targets:
            address = target["Address"]
            receiver = {
                "name": target["Name"],
                "email": address if "@" in address else None,
                "phone": address if "+" in address else None,
            }
            text = self.mailRefCoord.format(target["Name"], game.ak, date, time, all_names)
            cnt += self._dispatch(
                receiver,
                subject=self.mailRefCoordSubject,
                mail_body=text,
                sms_body=text,
                game_nr=game.game_nr,
            )
        return cnt

    # ---------------------------------------------------------------------------
    # Admin error notification (new unknown games)
    # ---------------------------------------------------------------------------

    def notify_new_games(self, games: list[db.GameEvent]) -> int:
        """Inform the admin about scraped games that were not known before."""
        if not games:
            return 0
        logging.warning(
            "Spielnummer not contained in home schedule, please correct manually!"
        )
        msg = EmailMessage()
        msg["From"] = formataddr((self.mail_name, self.mail_ID))
        msg["Subject"] = self.mailErrorSubject
        msg["To"] = formataddr(("Admin", self.mail_error_recipient))
        msg.set_content(self.mailError)
        self.send_Mail(msg, self.mail_ID, self.mail_password)
        return 1

    # ---------------------------------------------------------------------------
    # Newspaper article
    # ---------------------------------------------------------------------------

    def send_article(self, date: str, day: str, article_date: str) -> int:
        """Send schedule article for a game day to the local newspaper."""
        cnt = 0
        tournament_mi = False
        tournament_ge = False

        schedule = ""
        for game in db.get_games_on_date(self.session, date):
            team = {"F": "Damen", "M": "Herren"}.get(game.ak, game.ak)
            time_str = (game.time or "").strip(" v").strip(" t")

            if team == "MI" and not tournament_mi:
                schedule += f"Ab {time_str} Spielfest der Minis\n"
                tournament_mi = True
                cnt += 1
            elif team == "GE" and not tournament_ge:
                schedule += f"Ab {time_str} Turnier der gemischten E-Jugend\n"
                tournament_ge = True
                cnt += 1
            elif team not in ("GE", "MI"):
                schedule += f"{time_str} {team} {game.home} - {game.guest}\n"
                cnt += 1

        logging.info(f"Send newspaper article to {self.mailAddrNewspaper}")
        msg = EmailMessage()
        msg["From"] = formataddr((self.mail_name, self.mail_ID))
        msg["To"] = self.mailAddrNewspaper
        msg["Subject"] = self.mailNewspaperSubject
        msg.set_content(self.mailNewspaper.format(article_date, day, date, schedule))
        self.send_Mail(msg, self.mail_ID, self.mail_password)

        logging.info(f"Newspaper article for {cnt} games at {date} sent")
        return cnt
