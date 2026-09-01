# ---------------------------------------------------------------
#                          nuLigaHelper – tests
# ---------------------------------------------------------------
# Notification dispatch: channel selection, texts, receiver groups.
# All sends run against a TextRecorder – nothing goes out for real.
#
# Scenario: game 1001 on 05.09.2026, responsible team "BL mD".
# A person may only hold one task per game, so every slot has its own person:
#   Zeitnehmer=Alice(mail)  Sekretär=Bob(sms)
#   Verkauf=Caro(mail)      Verkauf=Dora(sms)
#   Ordnungsdienst=Ed(mail) Reinigung=Frida(sms, only when fully assigned)
#   Frank is MV of "BL mD" and receives the MV notification by SMS
#   as long as tasks of the game are still open.
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


def _setup(fully_assigned: bool = False):
    """Build the scenario database and return (session, game, notifier, recorder)."""
    engine = h.make_engine()
    session = h.Session(engine)
    h.sync_sample_games(session)

    game = session.query(db.Game).filter_by(source_key="test:1001").one()
    team = session.query(db.Team).filter_by(name="BL mD").one()  # own ak team
    game.team_id = team.id
    game.jteam = "BL mD"

    alice = db.get_or_create_person(session, "Alice", email="alice@x.de")
    bob = db.get_or_create_person(session, "Bob", phone="+491700000001")
    caro = db.get_or_create_person(session, "Caro", email="caro@x.de")
    dora = db.get_or_create_person(session, "Dora", phone="+491700000003")
    ed = db.get_or_create_person(session, "Ed", email="ed@x.de")
    frida = db.get_or_create_person(session, "Frida", phone="+491700000004")
    frank = db.get_or_create_person(session, "Frank", phone="+491700000002")
    frank.team_id = team.id
    db.set_team_mv(session, team, frank)

    # an open task means: nobody assigned at all
    cleaning = frida if fully_assigned else None
    for role, person in [
        (db.ROLE_TIMEKEEPER, alice), (db.ROLE_SECRETARY, bob),
        (db.ROLE_SALE, caro), (db.ROLE_SALE, dora),
        (db.ROLE_SECURITY, ed), (db.ROLE_CLEANING, cleaning),
    ]:
        if person is not None:
            db.assign_person(session, game, person, role)
    session.commit()

    n = notifier.Notifier(h.load_club_config(), session, h.SEASON)
    recorder = TextRecorder()
    n.send_Mail = recorder.mail
    n.send_SMS = recorder.sms
    return session, game, n, recorder


def test_notify_game_day_prefers_mail_then_sms_and_skips_missing_contacts():
    session, game, n, rec = _setup()
    # helpers: 5 assigned contacts (Caro not assigned anywhere here) plus
    # the MV reminder for the open Reinigung slot = 6
    assert n.notify_game_day(GAME_DATE) == 6


def test_automatic_account_messages_prefer_mail_with_sms_as_fallback():
    session, game, n, rec = _setup()
    both = db.Person(
        name="Both", email="both@example.test", phone="+491709999998"
    )
    phone_only = db.Person(name="Phone Only", phone="+491709999999")
    session.add_all([both, phone_only])
    session.commit()

    assert n.send_account_message(both, "Subject", "Mail body", "SMS body") == 1
    assert n.send_account_message(
        phone_only, "Subject", "Mail body", "SMS body"
    ) == 1
    assert len(rec.mails) == 1
    assert "both@example.test" in rec.mails[0][0]
    assert rec.smss == [("+491709999999", "SMS body")]


def test_mv_notification_only_while_tasks_are_open():
    session, game, n, rec = _setup(fully_assigned=True)
    assert not db.missing_slots(game), "scenario must be complete"

    cnt = n.notify_game_day(GAME_DATE)
    assert cnt == 6, "all six helper messages, but no MV reminder"
    assert all("+491700000002" != num for num, _ in rec.smss), \
        "the MV must not be contacted"


def test_mv_notification_contains_judge_team_and_both_judges():
    session, game, n, rec = _setup()
    n.notify_game_day(GAME_DATE)

    mv_smss = [body for num, body in rec.smss if num == "+491700000002"]
    assert len(mv_smss) == 1, "exactly one MV message expected"
    assert "BL mD" in mv_smss[0], "MV message must mention the responsible team"
    assert "Alice" in mv_smss[0] and "Bob" in mv_smss[0], \
        "MV message must list both judges"


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
    shifts = [db.ShiftEvent(game_id=game.id, game_nr=1001,
                            old_date="04.09.2026", old_time="15:00",
                            new_date="06.09.2026", new_time="18:00")]
    # helper roles only (no MV): 5 valid contacts, Caro has none
    assert n.notify_shifts(shifts) == 5


def test_referee_alert_targets_support_mail_and_sms():
    session, game, n, rec = _setup()
    event = db.RefereeEvent(
        game_id=game.id, game_nr=1001, date=GAME_DATE, time="15:00"
    )
    # config defines two targets (one phone, one e-mail); the MV only
    # has a phone number and is therefore not appended to the mail targets
    assert n.notify_referee_alert(event) == 2


def test_duplicate_number_events_notify_only_the_exact_game():
    engine = h.make_engine()
    session = h.Session(engine)
    rows = [
        {
            "source_key": "meeting:101", "day": "Sa", "date": GAME_DATE,
            "time": "10:00", "hall": 280340, "game_nr": 555, "ak": "BL mD",
            "home": "TuS Raubling", "guest": "Team A", "score": "",
        },
        {
            "source_key": "meeting:102", "day": "Sa", "date": GAME_DATE,
            "time": "11:00", "hall": 280340, "game_nr": 555, "ak": "BL mC",
            "home": "TuS Raubling", "guest": "Team B", "score": "",
        },
    ]
    db.sync_games(session, rows, h.SEASON)
    first = session.query(db.Game).filter_by(source_key="meeting:101").one()
    second = session.query(db.Game).filter_by(source_key="meeting:102").one()
    first_helper = db.Person(name="First Helper", email="first@x.de")
    second_helper = db.Person(name="Second Helper", email="second@x.de")
    first_team = db.get_or_create_team(session, "First Team")
    second_team = db.get_or_create_team(session, "Second Team")
    first_mv = db.Person(name="First MV", email="first-mv@x.de", team=first_team)
    second_mv = db.Person(name="Second MV", email="second-mv@x.de", team=second_team)
    session.add_all([first_helper, second_helper, first_mv, second_mv])
    session.flush()
    first.team, second.team = first_team, second_team
    db.set_team_mv(session, first_team, first_mv)
    db.set_team_mv(session, second_team, second_mv)
    db.assign_person(session, first, first_helper, db.ROLE_TIMEKEEPER)
    db.assign_person(session, second, second_helper, db.ROLE_TIMEKEEPER)
    session.commit()

    n = notifier.Notifier(h.load_club_config(), session, h.SEASON)
    recorder = TextRecorder()
    n.send_Mail, n.send_SMS = recorder.mail, recorder.sms
    shift = db.ShiftEvent(
        game_id=first.id, game_nr=555, old_date=GAME_DATE, old_time="10:00",
        new_date="06.09.2026", new_time="12:00",
    )
    assert n.notify_shifts([shift]) == 1
    assert "First Helper" in recorder.all_text()
    assert "Second Helper" not in recorder.all_text()

    recorder = TextRecorder()
    n.send_Mail, n.send_SMS = recorder.mail, recorder.sms
    alert = db.RefereeEvent(
        game_id=second.id, game_nr=555, date=GAME_DATE, time="11:00"
    )
    assert n.notify_referee_alert(alert) == 3
    recipients = " ".join(address for address, _, _ in recorder.mails)
    assert "Second MV" in recipients and "First MV" not in recipients


if __name__ == "__main__":
    h.run_all(dict(globals()))
