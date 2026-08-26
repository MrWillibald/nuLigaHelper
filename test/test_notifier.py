# ---------------------------------------------------------------
#                          nuLigaHelper – tests
# ---------------------------------------------------------------
# Notification dispatch: channel selection, texts, receiver groups.
# All sends run in debug mode; outgoing texts are captured instead.
#
# Scenario: game 1001 on 05.09.2026 is fully assigned:
#   Zeitnehmer=Alice(mail)  Sekretär=Bob(sms)
#   Verkauf=Alice(mail)     Verkauf=Bob(sms)
#   Ordnungsdienst=Alice    Reinigung=Caro(no contact -> always skipped)
#   MV=Frank(sms), Kampfgericht-Team="BL mD"
#
# Run standalone:  python test/test_notifier.py
# Or via pytest:   pytest test/test_notifier.py
# ---------------------------------------------------------------

import helpers as h
import db
import notifier

GAME_DATE = "05.09.2026"


class TextRecorder:
    """Replaces the real mail/SMS transport and records what would be sent."""

    def __init__(self):
        self.mails = []   # (recipient_name, subject, body)
        self.smss = []    # (number, body)

    def mail(self, msg, id_, password):  # signature matches Notifier.send_Mail
        self.mails.append((msg["To"], msg["Subject"], msg.get_content()))

    def sms(self, number, body):         # signature matches Notifier.send_SMS
        self.smss.append((number, body))

    def all_text(self) -> str:
        parts = [body for _, _, body in self.mails] + [b for _, b in self.smss]
        return "\n".join(parts)


def _setup():
    """Build the scenario database and return (session, game, notifier, recorder)."""
    engine = h.make_engine()
    session = h.Session(engine)
    h.sync_sample_games(session)

    game = session.query(db.Game).filter_by(game_nr=1001).one()
    game.jteam = "BL mD"  # team providing the judges (shown in the MV text)

    people = {
        "mail": db.get_or_create_person(session, "Alice", email="alice@x.de"),
        "sms": db.get_or_create_person(session, "Bob", phone="+491700000001"),
        "none": db.get_or_create_person(session, "Caro"),
        "mv": db.get_or_create_person(session, "Frank", phone="+491700000002"),
    }
    for role, key in [
        (db.ROLE_TIMEKEEPER, "mail"), (db.ROLE_SECRETARY, "sms"),
        (db.ROLE_SALE, "mail"), (db.ROLE_SALE, "sms"),
        (db.ROLE_SECURITY, "mail"), (db.ROLE_CLEANING, "none"),
        (db.ROLE_MV, "mv"),
    ]:
        db.assign_person(session, game, people[key], role)
    session.commit()

    n = notifier.Notifier(h.load_club_config(), session, h.SEASON)
    recorder = TextRecorder()
    n.send_Mail = recorder.mail
    n.send_SMS = recorder.sms
    return session, game, n, recorder


def test_notify_game_day_prefers_mail_then_sms_and_skips_missing_contacts():
    session, game, n, rec = _setup()
    # Alice gets e-mails, Bob gets SMS, Caro (no contact data) is skipped,
    # Frank receives the MV notification by SMS -> 6 dispatches total
    assert n.notify_game_day(GAME_DATE) == 6
    assert len(rec.mails) + len(rec.smss) == 6


def test_mv_notification_contains_judge_team_and_both_judges():
    session, game, n, rec = _setup()
    n.notify_game_day(GAME_DATE)

    text = rec.all_text()
    assert "BL mD" in text, "MV message must mention the responsible team"
    assert "Alice" in text and "Bob" in text, "MV message must list both judges"


def test_service_early_notifies_both_sales():
    session, game, n, rec = _setup()
    assert n.notify_service_early(GAME_DATE) == 2


def test_pre_notifications_skip_sale_roles_of_the_first_game():
    session, game, n, rec = _setup()
    # first (only) game: Zeitnehmer(1) + Sekretär(1) + Ordnungsdienst(1)
    # + Reinigung(Caro, skipped) = 3; Verkauf is deliberately excluded
    assert n.notify_pre(GAME_DATE) == 3


def test_shift_notifications_reach_all_assigned_helpers_except_missing_contacts():
    session, game, n, rec = _setup()
    shifts = [db.ShiftEvent(game_nr=1001, old_date="04.09.2026", old_time="15:00",
                            new_date="06.09.2026", new_time="18:00")]
    # helper roles only (no MV): 5 valid contacts, Caro has none
    assert n.notify_shifts(shifts) == 5


def test_referee_alert_targets_support_mail_and_sms():
    session, game, n, rec = _setup()
    event = db.RefereeEvent(game_nr=1001, date=GAME_DATE, time="15:00")
    # config defines two targets (one phone, one e-mail); the MV only
    # has a phone number and is therefore not appended to the mail targets
    assert n.notify_referee_alert(event) == 2


if __name__ == "__main__":
    h.run_all(dict(globals()))
