# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Notification dispatch: e-mail and SMS based on database content
# ---------------------------------------------------------------

import logging
import smtplib
from datetime import datetime
from email.utils import formataddr
from email.message import EmailMessage

from twilio.rest import Client

from common import DEBUG_FLAG
import common
import db
import messages


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

        twilio_cfg = config["twilio"]
        self.twilio_sid = twilio_cfg["twilio_sid"]
        self.twilio_token = twilio_cfg["twilio_token"]
        self.twilio_service_ID = twilio_cfg["twilio_service_ID"]
        common.warn_legacy_message_settings(config)
        self._referee_targets = common.referee_targets(config)

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
        sms_body: str | None,
        game_nr: str,
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
            logging.info("notification channel=email outcome=sent")
            return 1

        if isinstance(contact_phone, str) and "+" in contact_phone and sms_body is not None:
            self.send_SMS(contact_phone, sms_body)
            logging.info("notification channel=sms outcome=sent")
            return 1

        logging.warning("notification outcome=skipped reason=no_contact")
        return 0

    @staticmethod
    def _person_receiver(person: db.Person, task: str) -> dict:
        """Build a receiver dict from a Person instance."""
        return {"name": person.name, "email": person.email, "phone": person.phone, "task": task}

    def _dispatch_message(
        self, receiver: dict, message: messages.RenderedMessage, game_nr: str,
        mail_id: str | None = None, mail_password: str | None = None,
    ) -> int:
        return self._dispatch(
            receiver, subject=message.subject, mail_body=message.email,
            sms_body=message.sms, game_nr=game_nr,
            mail_id=mail_id, mail_password=mail_password,
        )

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
    # Game-day notifications (all occupied game duties + MV)
    # ---------------------------------------------------------------------------

    def notify_game_day(self, date: str) -> int:
        """Send notifications to all scheduled helpers of games on the given date."""
        cnt = 0
        games = db.get_games_on_date(self.session, date)

        for game in games:
            cnt += self._notify_game_helpers(game, date, "game.day_before")

            # Physical vacancies and unresolved eligibility both need follow-up.
            mv = game.team.mv_person if game.team is not None else None
            staffing = db.staffing_status(game)
            if (mv is None or mv.account_status != db.ACCOUNT_ACTIVE
                    or staffing["complete"]):
                continue
            judge_names = [
                a.person.name if a is not None else messages.render_fragment("staffing.position_unassigned")
                for a in (
                    game.assignment_by_role(db.ROLE_TIMEKEEPER),
                    game.assignment_by_role(db.ROLE_SECRETARY),
                )
            ]
            context = {
                "recipient_name": mv.name,
                "responsible_team_name": game.judge_team_name or "",
                "game_date": date,
                "timekeeper_name": judge_names[0],
                "secretary_name": judge_names[1],
                "timekeeper_task_label": db.ROLE_TIMEKEEPER,
                "secretary_task_label": db.ROLE_SECRETARY,
                "age_class": game.ak or "",
                "game_start_time": game.time or "",
                "age_eligibility_feedback": (
                    messages.render_fragment(
                        "staffing.deficiencies",
                        age_eligibility_reasons=" ".join(item["message"] for item in staffing["deficiencies"]),
                    ) if staffing["deficiencies"] else ""
                ),
            }
            if db.is_spielfest(game):
                key = "mv.spielfest.day_before"
            else:
                key = "mv.game.day_before"
                context.update(home_team_name=game.home or "", away_team_name=game.guest or "")
            cnt += self._dispatch_message(
                self._person_receiver(mv, "MV Verantwortlich"),
                messages.render(key, **context), game.game_nr,
            )

        return cnt

    def _block_time_text(self, block: db.DayBlock) -> str:
        calculated = db.calculated_block_time(block)
        if calculated is None:
            return messages.render_fragment("block.time_unset")
        text = calculated.strftime("%H:%M")
        try:
            home_date = datetime.strptime(block.date, "%d.%m.%Y").date()
        except ValueError:
            home_date = None
        if home_date is not None and calculated.date() != home_date:
            text = messages.render_fragment(
                "block.time_cross_date", block_meeting_time=text,
                block_meeting_date=calculated.strftime("%d.%m.%Y"),
            )
        return text

    def _blocks_for_date(self, date: str) -> list[db.DayBlock]:
        return db.get_day_blocks(self.session, self._season_year, date)

    @staticmethod
    def _block_reminder_task(block: db.DayBlock) -> str:
        if block.phase == db.BLOCK_CAKE_DELIVERY:
            return messages.render_fragment("task.cake", task_label=block.label)
        return block.label

    def notify_blocks_day_before(self, date: str) -> int:
        """Notify every occupied day-block slot one day ahead."""
        count = 0
        for block in self._blocks_for_date(date):
            time_text = self._block_time_text(block)
            for assignment in sorted(block.assignments, key=lambda item: item.slot):
                task = self._block_reminder_task(block)
                receiver = self._person_receiver(assignment.person, task)
                message = messages.render(
                    "block.day_before", recipient_name=receiver["name"],
                    block_date=date, task_label=task, block_meeting_time=time_text,
                )
                count += self._dispatch_message(receiver, message, f"block:{block.id}")
        return count

    # ---------------------------------------------------------------------------
    # Early preparation notifications (one week ahead)
    # ---------------------------------------------------------------------------

    def notify_service_early(self, date: str) -> int:
        """Send the special one-week reminder to preparation-block helpers."""
        blocks = [
            block for block in self._blocks_for_date(date)
            if block.phase == db.BLOCK_PREPARATION
        ]
        if not blocks:
            return 0
        block = blocks[0]
        assignments = sorted(block.assignments, key=lambda item: item.slot)
        count = 0
        time_text = self._block_time_text(block)
        for assignment in assignments:
            partner_names = ", ".join(
                other.person.name for other in assignments if other.id != assignment.id
            ) or messages.render_fragment("preparation.partner_none")
            task = block.label
            receiver = self._person_receiver(assignment.person, task)
            message = messages.render(
                "block.preparation.weekly", recipient_name=receiver["name"],
                block_date=date, task_label=task, partner_names=partner_names,
                block_meeting_time=time_text,
            )
            count += self._dispatch_message(
                receiver, message, f"block:{block.id}",
                mail_id=self.mail_saleID, mail_password=self.mail_salePassword,
            )
        return count

    def notify_cleanup_early(self, date: str) -> int:
        """Send an ordinary one-week reminder to occupied cleanup slots."""
        return self._notify_blocks_early(date, db.BLOCK_CLEANUP)

    def notify_cakes_early(self, date: str) -> int:
        """Remind each cake volunteer to deliver one cake one week ahead."""
        return self._notify_blocks_early(date, db.BLOCK_CAKE_DELIVERY)

    def _notify_blocks_early(self, date: str, phase: str) -> int:
        count = 0
        for block in self._blocks_for_date(date):
            if block.phase != phase:
                continue
            time_text = self._block_time_text(block)
            for assignment in sorted(block.assignments, key=lambda item: item.slot):
                task = self._block_reminder_task(block)
                receiver = self._person_receiver(assignment.person, task)
                message = messages.render(
                    "block.weekly", recipient_name=receiver["name"],
                    block_date=date, task_label=task, block_meeting_time=time_text,
                )
                count += self._dispatch_message(receiver, message, f"block:{block.id}")
        return count

    # ---------------------------------------------------------------------------
    # Pre-notifications (one week ahead)
    # ---------------------------------------------------------------------------

    def notify_pre(self, date: str) -> int:
        """Send pre-notifications to all assigned game helpers one week ahead."""
        cnt = 0
        games = db.get_games_on_date(self.session, date)

        for game in games:
            cnt += self._notify_game_helpers(
                game, date, "game.weekly",
                list(db.GAME_DAY_ROLES),
            )

        return cnt

    @staticmethod
    def _game_helper_assignments(game: db.Game, roles: list[str] | None = None):
        """Visit each saved position once, including retained removed duties."""
        for role in dict.fromkeys(db.GAME_DAY_ROLES if roles is None else roles):
            for assignment in game.assignments_by_role(role):
                yield role, assignment

    def _notify_game_helpers(
        self, game: db.Game, date: str, reminder_key: str,
        roles: list[str] | None = None,
    ) -> int:
        """Send task reminders using explicit day-before or weekly variants."""
        spielfest_keys = {
            "game.day_before": "spielfest.day_before",
            "game.weekly": "spielfest.weekly",
        }
        if reminder_key not in spielfest_keys:
            raise messages.MessageRenderError("Unknown game reminder key")
        cnt = 0
        for role, assignment in self._game_helper_assignments(game, roles):
            receiver = self._person_receiver(assignment.person, role)
            context = {
                "recipient_name": receiver["name"], "game_date": date,
                "task_label": role, "age_class": game.ak or "",
                "game_start_time": game.time or "",
            }
            if db.is_spielfest(game):
                key = spielfest_keys[reminder_key]
            else:
                key = reminder_key
                context.update(home_team_name=game.home or "", away_team_name=game.guest or "")
            cnt += self._dispatch_message(receiver, messages.render(key, **context), game.game_nr)
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
            for role, assignment in self._game_helper_assignments(game):
                receiver = self._person_receiver(assignment.person, role)
                context = {
                    "recipient_name": receiver["name"], "task_label": role,
                    "age_class": game.ak or "", "old_game_date": shift.old_date or "",
                    "old_game_start_time": shift.old_time or "", "new_game_date": shift.new_date or "",
                    "new_game_start_time": shift.new_time or "",
                }
                if db.is_spielfest(game):
                    key = "spielfest.shifted"
                else:
                    key = "game.shifted"
                    context.update(home_team_name=game.home or "", away_team_name=game.guest or "")
                cnt += self._dispatch_message(receiver, messages.render(key, **context), game.game_nr)
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
        targets = list(self._referee_targets)
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
            message = messages.render(
                "referee.missing", recipient_name=target["Name"], age_class=game.ak or "",
                game_date=date or "", game_start_time=time or "", notified_person_names=all_names,
            )
            cnt += self._dispatch_message(receiver, message, game.game_nr)
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
        message = messages.render("game.new")
        msg = EmailMessage()
        msg["From"] = formataddr((self.mail_name, self.mail_ID))
        msg["Subject"] = message.subject
        msg["To"] = formataddr(("Admin", self.mail_error_recipient))
        msg.set_content(message.email)
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
            team = {
                "F": messages.render_fragment("newspaper.women"),
                "M": messages.render_fragment("newspaper.men"),
            }.get(game.ak, game.ak)
            time_str = (game.time or "").strip(" v").strip(" t")

            if team == "MI" and not tournament_mi:
                schedule += messages.render_fragment("newspaper.minis_row", game_start_time=time_str)
                tournament_mi = True
                cnt += 1
            elif team == "GE" and not tournament_ge:
                schedule += messages.render_fragment("newspaper.e_youth_row", game_start_time=time_str)
                tournament_ge = True
                cnt += 1
            elif team not in ("GE", "MI"):
                schedule += messages.render_fragment(
                    "newspaper.game_row", game_start_time=time_str,
                    age_class_display_label=team or "", home_team_name=game.home or "",
                    away_team_name=game.guest or "",
                )
                cnt += 1

        logging.info("notification operation=article outcome=started")
        msg = EmailMessage()
        msg["From"] = formataddr((self.mail_name, self.mail_ID))
        msg["To"] = self.mailAddrNewspaper
        message = messages.render(
            "newspaper.article", article_date=article_date, game_weekday=day,
            game_date=date, schedule_text=schedule,
        )
        msg["Subject"] = message.subject
        msg.set_content(message.email)
        self.send_Mail(msg, self.mail_ID, self.mail_password)

        logging.info(f"Newspaper article for {cnt} games at {date} sent")
        return cnt
