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
#   Ordner=Ed(mail) Kasse=Frida(sms, only when fully assigned)
#   Frank is MV of "BL mD" and receives the MV notification by SMS
#   as long as tasks of the game are still open.
#
# Run standalone:  python test/test_notifier.py
# Or via pytest:   pytest test/test_notifier.py
# ---------------------------------------------------------------

from dataclasses import replace
from datetime import date, datetime
from unittest.mock import patch
from zoneinfo import ZoneInfo

from sqlalchemy import event, select

import helpers as h
import db
import notifier
import messages
import scraper

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


def _setup(fully_assigned: bool = False, missing_roles=(), missing_preparation_slots=()):
    """Build the scenario database and return (session, game, notifier, recorder)."""
    engine = h.make_engine()
    session = h.Session(engine)
    h.sync_sample_games(session)

    game = session.query(db.Game).filter_by(game_nr="1001").one()
    team = session.query(db.Team).filter_by(name="BL mD").one()  # own ak team
    game.team_id = team.id
    game.jteam = "BL mD"

    alice = db.get_or_create_person(session, "Alice", email="alice@x.de", birth_date=h.ADULT_BIRTH_DATE)
    bob = db.get_or_create_person(session, "Bob", phone="+491700000001", birth_date=h.ADULT_BIRTH_DATE)
    caro = db.get_or_create_person(session, "Caro", email="caro@x.de", birth_date=h.ADULT_BIRTH_DATE)
    dora = db.get_or_create_person(session, "Dora", phone="+491700000003", birth_date=h.ADULT_BIRTH_DATE)
    ed = db.get_or_create_person(session, "Ed", email="ed@x.de", birth_date=h.ADULT_BIRTH_DATE)
    frida = db.get_or_create_person(session, "Frida", phone="+491700000004", birth_date=h.ADULT_BIRTH_DATE)
    frank = db.get_or_create_person(session, "Frank", phone="+491700000002", birth_date=h.ADULT_BIRTH_DATE)
    frank.teams = [team]
    db.set_team_mv(session, team, frank)

    # an open task means: nobody assigned at all
    cash_helper = frida if fully_assigned else None
    for role, person in [
        (db.ROLE_TIMEKEEPER, alice), (db.ROLE_SECRETARY, bob),
        (db.ROLE_SALE, caro), (db.ROLE_SALE, dora),
        (db.ROLE_SECURITY, ed), (db.ROLE_CASH, cash_helper),
    ]:
        if person is not None and role not in missing_roles:
            if role == db.ROLE_CASH:
                # Synthetic migrated legacy assignment, now release-only on youth.
                session.add(db.Assignment(game=game, person=person, role=role, slot=0))
            else:
                db.assign_person(session, game, person, role)
    blocks = db.get_day_blocks(session, h.SEASON, GAME_DATE)
    preparation = next(b for b in blocks if b.phase == db.BLOCK_PREPARATION)
    cleanup = next(b for b in blocks if b.phase == db.BLOCK_CLEANUP)
    if 0 not in missing_preparation_slots:
        db.claim_block_slot(session, preparation, 0, None, caro)
    if 1 not in missing_preparation_slots:
        db.claim_block_slot(session, preparation, 1, None, dora)
    db.claim_block_slot(session, cleanup, 0, None, ed)
    session.commit()

    n = notifier.Notifier(h.load_club_config(), session, h.SEASON)
    recorder = TextRecorder()
    n.send_Mail = recorder.mail
    n.send_SMS = recorder.sms
    return session, game, n, recorder


def test_notify_game_day_prefers_mail_then_sms_and_skips_missing_contacts():
    session, game, n, rec = _setup()
    # Youth's five occupied game tasks are complete without a Kasse.
    assert n.notify_game_day(GAME_DATE) == 5


def test_automatic_account_messages_prefer_mail_with_sms_as_fallback():
    session, game, n, rec = _setup()
    both = db.Person(
        birth_date=h.ADULT_BIRTH_DATE,
        name="Both", email="both@example.test", phone="+491709999998"
    )
    phone_only = db.Person(birth_date=h.ADULT_BIRTH_DATE, name="Phone Only", phone="+491709999999")
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
    session, game, n, rec = _setup(missing_roles=(db.ROLE_SECURITY,))
    session.commit()
    n.notify_game_day(GAME_DATE)

    mv_smss = [body for num, body in rec.smss if num == "+491700000002"]
    assert len(mv_smss) == 1, "exactly one MV message expected"
    assert "BL mD" in mv_smss[0], "MV message must mention the responsible team"
    assert "Alice" in mv_smss[0] and "Bob" in mv_smss[0], \
        "MV message must list both judges"


def test_service_early_notifies_preparation_block_helpers():
    session, game, n, rec = _setup()
    assert n.notify_service_early(GAME_DATE) == 2


def test_pre_notifications_include_sale_roles_of_the_first_game():
    session, game, n, rec = _setup()
    assert n.notify_pre(GAME_DATE) == 5


def test_block_reminders_cover_cleanup_weekly_and_all_occupied_slots_day_before():
    session, game, n, rec = _setup()
    assert n.notify_cleanup_early(GAME_DATE) == 1
    assert n.notify_blocks_day_before(GAME_DATE) == 3
    assert "Vorbereitung" in rec.all_text()
    assert "Aufräumen" in rec.all_text()
    assert "Vorbereitung 1" not in rec.all_text()
    assert "Aufräumen 1" not in rec.all_text()


def test_shift_notifications_reach_all_assigned_helpers_except_missing_contacts():
    session, game, n, rec = _setup()
    shifts = [db.ShiftEvent(game_id=game.id, game_nr="1001",
                            old_date="04.09.2026", old_time="15:00",
                            new_date="06.09.2026", new_time="18:00")]
    # Helper roles only (no MV): all five baseline helpers have a valid contact.
    assert n.notify_shifts(shifts) == 5


def _add_cleaning_helpers(session, game):
    helpers = [
        db.Person(
            name="Cleanup First", email="cleanup-first@example.test",
            birth_date=h.ADULT_BIRTH_DATE,
        ),
    ]
    session.add_all(helpers)
    session.flush()
    for slot, person in enumerate(helpers):
        if (db.ROLE_CLEANING, slot) in db.offered_positions(game):
            db.claim_slot(session, game, db.ROLE_CLEANING, slot, None, person)
        else:
            # Preserve fixtures for duties retained after an adult -> youth change.
            session.add(db.Assignment(game=game, person=person, role=db.ROLE_CLEANING, slot=slot))
    session.commit()
    return helpers


def test_retained_removed_youth_duties_receive_one_reminder_each_with_shared_role_names():
    session, game, sender, _ = _setup(fully_assigned=True)
    _add_cleaning_helpers(session, game)
    assert db.staffing_status(game)["complete"], "retained removed duties must not alter completeness"
    shifts = [db.ShiftEvent(
        game_id=game.id, game_nr=game.game_nr, old_date=GAME_DATE,
        old_time="15:00", new_date="06.09.2026", new_time="18:00",
    )]
    for notify in (
        lambda: sender.notify_game_day(GAME_DATE),
        lambda: sender.notify_pre(GAME_DATE),
        lambda: sender.notify_shifts(shifts),
    ):
        recorder = TextRecorder()
        sender.send_Mail, sender.send_SMS = recorder.mail, recorder.sms
        assert notify() == 7, "every saved duty, including retained Kasse/Wischer, needs one message"
        assert len(recorder.mails) == 4 and len(recorder.smss) == 3, \
            "the two occupants of repeated roles must not drop or duplicate recipients"
        recipients = [recipient for recipient, _, _ in recorder.mails]
        for name in ("Alice", "Caro", "Ed", "Cleanup First"):
            assert sum(name in recipient for recipient in recipients) == 1, \
                f"{name} must receive exactly one reminder"
        phones = [phone for phone, _ in recorder.smss]
        for phone in ("+491700000001", "+491700000003", "+491700000004"):
            assert phones.count(phone) == 1, "each SMS occupant must receive exactly one reminder"
        for recipient, subject, body in recorder.mails:
            if "Caro" in recipient:
                assert "Verkauf" in subject and "Verkauf" in body
            if "Ed" in recipient:
                assert "Ordner" in subject and "Ordner" in body
                assert "Ordnungsdienst" not in subject + body
            if "Cleanup First" in recipient:
                assert "Wischer" in subject and "Wischer" in body
        for phone, body in recorder.smss:
            if phone == "+491700000003":
                assert "Verkauf" in body
            if phone == "+491700000004":
                assert "Kasse" in body
        for text in (recorder.all_text(), " ".join(subject for _, subject, _ in recorder.mails)):
            for numbered_role in ("Verkauf 1", "Verkauf 2", "Wischer 1", "Wischer 2"):
                assert numbered_role not in text, "reminders must use the shared semantic role"


def test_ordner_reminders_and_shifts_use_current_name_for_both_channels_and_events():
    for category in ("BL M", "SPF Mini"):
        for channel in ("email", "sms"):
            session, game, sender, _ = _setup()
            game.ak = category
            person = game.assignment_by_role(db.ROLE_SECURITY).person
            person.phone = "+491700999999"
            if channel == "sms":
                person.email = None
            session.commit()
            shift = db.ShiftEvent(game_id=game.id, game_nr=game.game_nr,
                old_date=GAME_DATE, old_time="15:00", new_date="06.09.2026", new_time="18:00")
            for notify, expected in (
                (lambda: sender.notify_game_day(GAME_DATE), 6 if category == "BL M" else 5),
                (lambda: sender.notify_pre(GAME_DATE), 5),
                (lambda: sender.notify_shifts([shift]), 5),
            ):
                recorder = TextRecorder()
                sender.send_Mail, sender.send_SMS = recorder.mail, recorder.sms
                assert notify() == expected
                if channel == "email":
                    texts = [subject + body for recipient, subject, body in recorder.mails if "Ed" in recipient]
                else:
                    texts = [body for phone, body in recorder.smss if phone == person.phone]
                assert len(texts) == 1 and "Ordner" in texts[0]
                assert "Ordnungsdienst" not in texts[0]
                assert ("Spielfest" in texts[0]) == (category == "SPF Mini")
            session.close()


def test_single_wischer_and_repeated_role_input_notify_saved_person_once():
    session, game, sender, recorder = _setup()
    game.ak = "BL M"
    session.commit()
    helper = db.Person(
        name="Only Wischer", email="only-wischer@example.test",
        birth_date=h.ADULT_BIRTH_DATE,
    )
    session.add(helper)
    session.flush()
    db.claim_slot(session, game, db.ROLE_CLEANING, 0, None, helper)
    session.commit()

    assert sender._notify_game_helpers(
        game, GAME_DATE, "game.day_before",
        [db.ROLE_CLEANING, db.ROLE_CLEANING],
    ) == 1, "role requests select actual saved assignments once, including Wischer"
    assert len(recorder.mails) == 1 and "Only Wischer" in recorder.mails[0][0]
    assert "Wischer" in recorder.mails[0][2]


def test_mv_reminder_tracks_adult_added_duties_and_stops_when_all_seven_are_filled():
    session, game, sender, recorder = _setup()
    game.ak = "BL M"
    session.commit()
    assert sender.notify_game_day(GAME_DATE) == 6, \
        "adult Kasse and Wischer gaps need one MV follow-up alongside five helpers"
    assert sum(phone == "+491700000002" for phone, _ in recorder.smss) == 1

    cash = session.query(db.Person).filter_by(name="Frida").one()
    db.claim_slot(session, game, db.ROLE_CASH, 0, None, cash)
    _add_cleaning_helpers(session, game)
    assert db.staffing_status(game)["complete"], "all seven adult positions must complete staffing"
    recorder = TextRecorder()
    sender.send_Mail, sender.send_SMS = recorder.mail, recorder.sms
    assert sender.notify_game_day(GAME_DATE) == 7
    assert all(phone != "+491700000002" for phone, _ in recorder.smss), \
        "complete adult staffing must stop the MV reminder"


def test_referee_alert_targets_support_mail_and_sms():
    session, game, n, rec = _setup()
    event = db.RefereeEvent(
        game_id=game.id, game_nr="1001", date=GAME_DATE, time="15:00"
    )
    # config defines two targets (one phone, one e-mail); the MV only
    # has a phone number and is therefore not appended to the mail targets
    assert n.notify_referee_alert(event) == 2


def test_same_number_in_different_seasons_notifies_only_the_exact_game():
    engine = h.make_engine()
    session = h.Session(engine)
    rows = [
        {
            "day": "Sa", "date": GAME_DATE,
            "time": "10:00", "hall": 280340, "game_nr": "555", "ak": "BL mD",
            "home": "TuS Raubling", "guest": "Team A", "score": "",
        },
        {
            "day": "Sa", "date": "05.09.2027",
            "time": "11:00", "hall": 280340, "game_nr": "555", "ak": "BL mC",
            "home": "TuS Raubling", "guest": "Team B", "score": "",
        },
    ]
    db.sync_games(session, [rows[0]], h.SEASON)
    db.sync_games(session, [rows[1]], h.SEASON + 1)
    first = session.query(db.Game).filter_by(season_year=h.SEASON, game_nr="555").one()
    second = session.query(db.Game).filter_by(season_year=h.SEASON + 1, game_nr="555").one()
    first_helper = db.Person(birth_date=h.ADULT_BIRTH_DATE, name="First Helper", email="first@x.de")
    second_helper = db.Person(birth_date=h.ADULT_BIRTH_DATE, name="Second Helper", email="second@x.de")
    first_team = db.get_or_create_team(session, "First Team")
    second_team = db.get_or_create_team(session, "Second Team")
    first_mv = db.Person(birth_date=h.ADULT_BIRTH_DATE, name="First MV", email="first-mv@x.de", teams=[first_team])
    second_mv = db.Person(birth_date=h.ADULT_BIRTH_DATE, name="Second MV", email="second-mv@x.de", teams=[second_team])
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
        game_id=first.id, game_nr="555", old_date=GAME_DATE, old_time="10:00",
        new_date="06.09.2026", new_time="12:00",
    )
    assert n.notify_shifts([shift]) == 1
    assert "First Helper" in recorder.all_text()
    assert "Second Helper" not in recorder.all_text()

    recorder = TextRecorder()
    n.send_Mail, n.send_SMS = recorder.mail, recorder.sms
    alert = db.RefereeEvent(
        game_id=second.id, game_nr="555", date="05.09.2027", time="11:00"
    )
    assert n.notify_referee_alert(alert) == 3
    recipients = " ".join(address for address, _, _ in recorder.mails)
    assert "Second MV" in recipients and "First MV" not in recipients


def test_spielfest_notifications_describe_the_aggregate_event():
    engine = h.make_engine()
    session = h.Session(engine)
    rows = [{
        "day": "Sa", "date": GAME_DATE, "time": "10:00", "hall": 280340,
        "game_nr": "1", "ak": "SPF Mini", "home": "Team A",
        "guest": "Team B", "score": "§77",
    }]
    db.sync_games(session, scraper.collapse_spielfeste(rows), h.SEASON)
    game = session.query(db.Game).one()
    helper = db.Person(birth_date=h.ADULT_BIRTH_DATE, name="SPF Helper", email="spf@example.test")
    session.add(helper)
    db.assign_person(session, game, helper, db.ROLE_TIMEKEEPER)
    session.commit()

    n = notifier.Notifier(h.load_club_config(), session, h.SEASON)
    recorder = TextRecorder()
    n.send_Mail, n.send_SMS = recorder.mail, recorder.sms
    assert n.notify_game_day(GAME_DATE) == 1
    assert "Spielfest SPF Mini" in recorder.all_text()
    assert "gegen Team" not in recorder.all_text()
    assert n.notify_referees_for_date(GAME_DATE) == 0


def test_catalog_copy_controls_sends_and_legacy_texts_cannot_override_it():
    session, game, _, recorder = _setup()
    config = h.load_club_config()
    config["texts"] = {
        "mailTask": "PRIVATE_LEGACY_OVERRIDE",
        "textTask": "PRIVATE_LEGACY_SMS_OVERRIDE",
        "send_Mail": "attempted_attribute_override",
    }
    sender = notifier.Notifier(config, session, h.SEASON)
    sender.send_Mail, sender.send_SMS = recorder.mail, recorder.sms
    original = messages.CATALOG["game.day_before"]
    with patch.dict(messages.CATALOG, {"game.day_before": replace(
        original, email=original.email + " EMAIL_CATALOG_MARKER",
        sms=original.sms + " SMS_CATALOG_MARKER",
    )}):
        assert sender.notify_game_day(GAME_DATE) == 5
    assert all("EMAIL_CATALOG_MARKER" in body for _, _, body in recorder.mails)
    assert all("SMS_CATALOG_MARKER" in body for _, body in recorder.smss)
    assert "PRIVATE_LEGACY" not in recorder.all_text()
    assert callable(sender.send_Mail), "legacy text fields cannot overwrite notifier methods"


def test_missing_contact_skips_exactly_one_helper_and_mail_still_wins_with_both():
    session, game, sender, recorder = _setup()
    alice = session.query(db.Person).filter_by(name="Alice").one()
    alice.phone = "+491701234567"
    ed = session.query(db.Person).filter_by(name="Ed").one()
    ed.email = ed.phone = None
    session.commit()
    assert sender.notify_game_day(GAME_DATE) == 4, "missing contacts must count as skipped"
    assert sum("Alice" in recipient for recipient, _, _ in recorder.mails) == 1
    assert all(phone != alice.phone for phone, _ in recorder.smss), "ordinary reminders must prefer email"
    assert "Hallo Ed" not in recorder.all_text()


def test_referee_recorder_receives_correct_named_facts_in_both_channels():
    session, game, sender, recorder = _setup()
    sender._referee_targets = [
        {"Name": "Coordinator Mail", "Address": "referee@example.test"},
        {"Name": "Coordinator SMS", "Address": "+491700009999"},
    ]
    game.ak = "Distinct Age Class"
    event = db.RefereeEvent(game_id=game.id, game_nr=game.game_nr,
                            date="31.12.2042", time="23:17")
    assert sender.notify_referee_alert(event) == 2
    for body in [recorder.mails[0][2], recorder.smss[0][1]]:
        assert "Heimspiel der Distinct Age Class am 31.12.2042 um 23:17" in body
        assert "Coordinator Mail, Coordinator SMS" in body


def test_nullable_legacy_event_times_remain_notifiable():
    session, game, sender, recorder = _setup()
    game.time = None
    game.score = "§77"
    session.commit()
    assert sender.notify_referees_for_date(GAME_DATE) == 2, "missing saved time must not abort referee alerts"
    shifts = [db.ShiftEvent(game_id=game.id, game_nr=game.game_nr,
                            old_date=None, old_time=None,
                            new_date="06.09.2026", new_time="18:00")]
    assert sender.notify_shifts(shifts) == 5, "rescheduling from an unknown old time remains supported"


def test_preparation_uses_existing_sender_partner_names_and_meeting_time():
    session, game, sender, recorder = _setup()
    sent_credentials = []
    def record_mail(message, sender_id, password):
        sent_credentials.append((sender_id, password))
        recorder.mail(message, sender_id, password)
    sender.send_Mail = record_mail
    assert sender.notify_service_early(GAME_DATE) == 2
    assert sent_credentials == [(sender.mail_saleID, sender.mail_salePassword)]
    email_body = recorder.mails[0][2]
    assert "Dora" in email_body and "Treffpunkt: 13:30" in email_body
    assert "Caro" in recorder.smss[0][1] and "Treffpunkt: 13:30" in recorder.smss[0][1]


def test_spielfest_selects_day_before_weekly_mv_and_shift_keys_explicitly():
    session, game, sender, recorder = _setup(missing_roles=(db.ROLE_SECURITY,))
    game.ak = "SPF Mini"
    game.home = game.guest = ""
    session.commit()
    assert sender.notify_game_day(GAME_DATE) == 5, "four helpers and incomplete-staffing MV need a reminder"
    assert "morgen" in recorder.all_text() and "Spielfest SPF Mini" in recorder.all_text()
    assert "gegen" not in recorder.all_text()
    recorder = TextRecorder()
    sender.send_Mail, sender.send_SMS = recorder.mail, recorder.sms
    assert sender.notify_pre(GAME_DATE) == 4
    assert "nächste Woche" in recorder.all_text() and "morgen" not in recorder.all_text()
    shifts = [db.ShiftEvent(game_id=game.id, game_nr=game.game_nr,
                            old_date=GAME_DATE, old_time="15:00",
                            new_date="06.09.2026", new_time="18:00")]
    assert sender.notify_shifts(shifts) == 4
    assert "Spielfest SPF Mini wurde verschoben" in recorder.all_text()
    assert "gegen" not in recorder.all_text()


def test_mv_zero_and_one_timing_occupant_have_readable_labels_and_vacancy_status():
    for age_class in ("BL mD", "SPF Mini"):
        for release_roles in ((db.ROLE_TIMEKEEPER,), (db.ROLE_TIMEKEEPER, db.ROLE_SECRETARY)):
            session, game, sender, recorder = _setup(missing_roles=release_roles)
            game.ak = age_class
            if db.is_spielfest(game):
                game.home = game.guest = ""
            session.commit()
            assert sender.notify_game_day(GAME_DATE) == 6 - len(release_roles)
            mv_body = next(body for phone, body in recorder.smss if phone == "+491700000002")
            assert "Zeitnehmer: noch offen" in mv_body
            secretary_status = "noch offen" if db.ROLE_SECRETARY in release_roles else "Bob"
            assert f"Sekretär: {secretary_status}" in mv_body
            assert "BL mD" in mv_body and GAME_DATE in mv_body and "15:00" in mv_body
            assert "Bitte prüfe die offenen Dienste" in mv_body
            assert " und  die Verantwortung" not in mv_body
            if db.is_spielfest(game):
                assert "Spielfest SPF Mini" in mv_body and "gegen" not in mv_body


def test_mv_full_but_age_deficient_staffing_keeps_named_occupants_and_follow_up():
    session, game, sender, recorder = _setup()
    alice = session.query(db.Person).filter_by(name="Alice").one()
    alice.birth_date = None  # Synthetic retained legacy data; no new claim is made.
    session.commit()
    assert not db.missing_slots(game), "this scenario has no physical vacancies"
    assert sender.notify_game_day(GAME_DATE) == 6, "five helpers and age-deficiency MV follow-up expected"
    mv_body = next(body for phone, body in recorder.smss if phone == "+491700000002")
    assert "Zeitnehmer: Alice" in mv_body and "Sekretär: Bob" in mv_body
    assert "Altersanforderungen offen" in mv_body and "noch offen" not in mv_body.split("Altersanforderungen")[0]
    assert "birth_date" not in mv_body and "1990-01-01" not in mv_body


def test_block_time_fallback_and_cross_date_time_come_from_catalog_fragments():
    session, game, sender, recorder = _setup(missing_preparation_slots=(1,))
    preparation = next(block for block in sender._blocks_for_date(GAME_DATE)
                       if block.phase == db.BLOCK_PREPARATION)
    game.time = "00:30"
    session.commit()
    assert sender.notify_service_early(GAME_DATE) == 1
    assert "niemandem" in recorder.all_text() and "23:00 am 04.09.2026" in recorder.all_text()
    recorder = TextRecorder()
    sender.send_Mail, sender.send_SMS = recorder.mail, recorder.sms
    game.time = None
    session.commit()
    assert sender.notify_service_early(GAME_DATE) == 1
    assert "Treffpunkt: noch offen" in recorder.all_text(), "unknown time must stay explicit"


def test_new_game_alert_preserves_its_trigger_and_dedicated_admin_recipient():
    session, game, sender, recorder = _setup()
    assert sender.notify_new_games([]) == 0 and not recorder.mails
    assert sender.notify_new_games([db.GameEvent(game_id=game.id, game_nr=game.game_nr, ak=game.ak)]) == 1
    assert len(recorder.mails) == 1 and "admin@test.invalid" in recorder.mails[0][0]
    message = messages.render("game.new")
    assert recorder.mails[0][1] == message.subject and message.email in recorder.mails[0][2]
    assert not recorder.smss


def test_retained_newspaper_callable_uses_catalog_article_and_schedule_fragments():
    session, game, sender, recorder = _setup()
    assert sender.send_article(GAME_DATE, "Samstag", "04.09.2026") == 1
    assert len(recorder.mails) == 1 and "paper@test.invalid" in recorder.mails[0][0]
    body = recorder.mails[0][2]
    assert "am 04.09.2026" in body and "Am Samstag, den 05.09.2026" in body
    assert "15:00 BL mD TuS Raubling - SBC Traunstein" in body
    assert not recorder.smss, "newspaper path remains email-only"


def test_invalid_catalog_subject_is_refused_before_a_helper_send():
    session, game, sender, recorder = _setup()
    original = messages.CATALOG["game.day_before"]
    with patch.dict(messages.CATALOG, {"game.day_before": replace(original, subject="Invalid\nHeader")}):
        try:
            sender.notify_game_day(GAME_DATE)
        except messages.MessageRenderError:
            pass
        else:
            raise AssertionError("an invalid template produced a notification")
    assert not recorder.mails and not recorder.smss, "render failure must precede provider dispatch"


def _vacancy_setup(age_class="BL M"):
    """One empty game with a saved responsible team and recorded transports."""
    session = h.Session(h.make_engine())
    row = {**h.sample_games()[1], "ak": age_class}
    rows = scraper.collapse_spielfeste([row]) if age_class.startswith("SPF") else [row]
    db.sync_games(session, rows, h.SEASON)
    game = session.query(db.Game).one()
    # Responsibility intentionally differs from the playing team.
    team = session.query(db.Team).filter_by(is_support=True).one()
    game.team = team
    session.commit()
    sender = notifier.Notifier(h.load_club_config(), session, h.SEASON)
    recorder = TextRecorder()
    sender.send_Mail, sender.send_SMS = recorder.mail, recorder.sms
    return session, game, sender, recorder


def _vacancy_member(session, team, name="Invited", **values):
    person = db.Person(**{
        "name": name, "email": f"member-{session.query(db.Person).count()}@example.test",
        "birth_date": h.ADULT_BIRTH_DATE, "teams": [team], **values,
    })
    session.add(person)
    session.commit()
    return person


def _fill_vacancy_positions(session, game, positions):
    for role, slot in positions:
        person = _vacancy_member(session, game.team, name=f"Assigned {role} {slot}")
        db.claim_slot(session, game, role, slot, None, person)
    session.commit()


def test_team_vacancies_require_active_saved_membership_but_no_mv():
    session, game, sender, rec = _vacancy_setup()
    active = _vacancy_member(session, game.team)
    for status in (db.ACCOUNT_REGISTERED, db.ACCOUNT_VERIFIED, db.ACCOUNT_REJECTED, db.ACCOUNT_INACTIVE):
        _vacancy_member(session, game.team, name=status, account_status=status)
    playing_team = session.query(db.Team).filter_by(name=game.ak).one()
    _vacancy_member(session, playing_team, name="Playing team only", is_admin=True)
    _vacancy_member(session, None, name="No team", teams=[])
    assert game.team.mv_person is None, "recruitment must not need an MV appointment"
    assert sender.notify_team_vacancies_early(GAME_DATE) == 1
    assert len(rec.mails) == 1 and active.email in rec.mails[0][0]
    assert "Verkauf: 2 Plätze" in rec.all_text()
    assert "Kasse: 1 Platz; Wischer: 1 Platz" in rec.all_text()
    game.team = None
    session.commit()
    assert sender.notify_team_vacancies_early(GAME_DATE) == 0, "no fallback team or admin recipients"
    session.close()


def test_team_vacancies_invite_qualifying_mv_once_and_exclude_an_assigned_mv():
    session, game, sender, rec = _vacancy_setup()
    mv = _vacancy_member(session, game.team, is_admin=True)
    db.set_team_mv(session, game.team, mv)
    assert sender.notify_team_vacancies_early(GAME_DATE) == 1
    assert len(rec.mails) == 1, "MV appointment must not duplicate the member invitation"
    db.claim_slot(session, game, db.ROLE_TIMEKEEPER, 0, None, mv)
    session.commit()
    assert sender.notify_team_vacancies_early(GAME_DATE) == 0, "admin/MV has the same assignment exclusion"
    session.close()


def test_team_vacancies_use_required_counts_for_adult_youth_and_spielfest():
    for age_class in ("BL M", "BL mD", "SPF Mini vollständig"):
        session, game, sender, rec = _vacancy_setup(age_class)
        _vacancy_member(session, game.team)
        _fill_vacancy_positions(session, game, db.BASELINE_REQUIRED_POSITIONS)
        count = sender.notify_team_vacancies_early(GAME_DATE)
        if age_class == "BL M":
            assert count == 1
            assert "\n\nKasse: 1 Platz; Wischer: 1 Platz\n\n" in rec.all_text()
        else:
            assert count == 0, "youth/SPF need no cash or Wischer staffing"
        session.close()


def test_team_vacancies_exclude_all_same_game_duties_and_preserve_other_container_recipients():
    session, game, sender, rec = _vacancy_setup("BL mD")
    excluded = _vacancy_member(session, game.team, name="Same name")
    included = _vacancy_member(session, game.team, name="Same name", birth_date=None)
    replacement = _vacancy_member(session, game.team, name="Replacement searching")
    retired = _vacancy_member(session, game.team, name="Retained removed duty")
    young = _vacancy_member(session, game.team, name="Young member", birth_date=date(2020, 1, 1))
    other_game_helper = _vacancy_member(session, game.team, name="Other game")
    block_helper = _vacancy_member(session, game.team, name="Other block")
    db.claim_slot(session, game, db.ROLE_TIMEKEEPER, 0, None, excluded)
    db.claim_slot(session, game, db.ROLE_SALE, 0, None, replacement)
    token = db.request_replacement(session, game, db.ROLE_SALE, 0, replacement.id)
    # Synthetic legacy duty retained after a category change.
    session.add(db.Assignment(game=game, person=retired, role=db.ROLE_CASH, slot=0))
    other_game = db.Game(season_year=h.SEASON, game_nr="9999", date=GAME_DATE, ak="BL mD")
    session.add(other_game)
    session.commit()
    db.claim_slot(session, other_game, db.ROLE_TIMEKEEPER, 0, None, other_game_helper)
    block = next(b for b in db.get_day_blocks(session, h.SEASON, GAME_DATE) if b.phase == db.BLOCK_PREPARATION)
    db.claim_block_slot(session, block, 0, None, block_helper)
    playing_team = session.query(db.Team).filter_by(name="BL mD").one()
    included.teams.append(playing_team)
    session.commit()
    assert token and game.assignment_by_role(db.ROLE_SALE).replacement_token
    assert sender.notify_team_vacancies_early(GAME_DATE) == 4
    addresses = [address for address, _, _ in rec.mails]
    for person in (included, young, other_game_helper, block_helper):
        assert sum(person.email in address for address in addresses) == 1, "one invitation per qualifying identity"
    for person in (excluded, replacement, retired):
        assert all(person.email not in address for address in addresses), "every saved same-game duty excludes its helper"
    assert "Verkauf: 1 Platz" in rec.all_text() and "Verkauf: 2 Plätze" not in rec.all_text()
    assert "Kasse" not in rec.all_text(), "retained removed positions are not advertised"
    session.close()


def test_team_vacancies_are_season_date_scoped_and_per_game_per_invocation():
    session, game, sender, rec = _vacancy_setup()
    _vacancy_member(session, game.team)
    for season, number, game_date in (
        (h.SEASON, "1002", GAME_DATE),
        (h.SEASON - 1, game.game_nr, GAME_DATE),
        (h.SEASON, "1003", "06.09.2026"),
    ):
        session.add(db.Game(season_year=season, game_nr=number, date=game_date, ak="BL M", team=game.team))
    session.commit()
    assert sender.notify_team_vacancies_early(GAME_DATE) == 2, "only two games in this season and date qualify"
    assert sender.notify_team_vacancies_early(GAME_DATE) == 2, "reruns follow existing reminder resend behavior"
    assert len(rec.mails) == 4, "each game and invocation has its own invitation"
    session.close()


def test_team_vacancies_route_contacts_and_render_both_events_without_opponent_invention():
    for age_class in ("BL M", "SPF Mini vollständig"):
        session, game, sender, rec = _vacancy_setup(age_class)
        both = _vacancy_member(session, game.team, phone="+491700111111")
        phone = _vacancy_member(session, game.team, name="Phone", email=None, phone="+491700222222")
        _vacancy_member(session, game.team, name="Contactless", email=None)
        game.time = None
        session.commit()
        with patch.object(notifier.logging, "warning") as warning:
            assert sender.notify_team_vacancies_early(GAME_DATE) == 2, "contactless member contributes zero"
        warning.assert_called_once_with("notification outcome=skipped reason=no_contact")
        assert len(rec.mails) == len(rec.smss) == 1 and both.email in rec.mails[0][0]
        assert rec.smss[0][0] == phone.phone, "both contacts must prefer e-mail"
        for body in (rec.mails[0][2], rec.smss[0][1]):
            assert game.team.name in body and GAME_DATE in body and age_class in body
            assert "Ordner: 1 Platz" in body and "Ordnungsdienst" not in body
            if age_class.startswith("SPF"):
                assert "Spielfest SPF Mini vollständig" in body and "gegen" not in body
        session.close()


def test_team_vacancy_provider_errors_propagate_without_fallback_or_writes():
    for channel in ("email", "sms"):
        session, game, sender, rec = _vacancy_setup()
        _vacancy_member(session, game.team, email="both@example.test" if channel == "email" else None, phone="+491700111111")
        error = RuntimeError("synthetic provider failure")
        with patch.object(sender, "send_Mail" if channel == "email" else "send_SMS", side_effect=error) as send:
            try:
                sender.notify_team_vacancies_early(GAME_DATE)
            except RuntimeError as actual:
                assert actual is error, "provider failures must reach daily-job reporting"
            else:
                raise AssertionError("provider failure was silently treated as success")
        send.assert_called_once()
        assert not rec.mails and not rec.smss, "SMTP failure must not fall back to SMS"
        assert not session.new and not session.dirty and not session.deleted
        session.close()


def test_team_vacancy_dispatch_does_not_write_or_lock_even_when_cutoff_is_due():
    session, game, sender, rec = _vacancy_setup()
    _vacancy_member(session, game.team)
    assigned = _vacancy_member(session, game.team, name="Assigned")
    db.claim_slot(session, game, db.ROLE_SALE, 0, None, assigned)
    db.request_replacement(session, game, db.ROLE_SALE, 0, assigned.id)
    session.commit()
    def snapshot():
        return {table.name: session.execute(select(table)).all() for table in db.Base.metadata.sorted_tables}
    before = snapshot()
    statements = []
    def record_sql(_conn, _cursor, statement, _parameters, _context, _many):
        statements.append(statement.lstrip().split()[0].upper())
    event.listen(session.bind, "before_cursor_execute", record_sql)
    try:
        # Weekly invitations remain read-only even if explicitly invoked late.
        with patch.object(db, "duty_now", return_value=datetime(2026, 9, 4, 10, tzinfo=ZoneInfo("Europe/Berlin"))):
            assert sender.notify_team_vacancies_early(GAME_DATE) == 1
        assert set(statements) <= {"SELECT"}, "sender must never write, commit, lock or begin a serialized writer"
        assert not session.new and not session.dirty and not session.deleted
        assert snapshot() == before, "assignments, replacement tokens, cutoffs and audit history must remain identical"
    finally:
        event.remove(session.bind, "before_cursor_execute", record_sql)
        session.close()


def test_team_vacancies_ignore_deficiencies_replacement_searches_and_block_gaps():
    for age_class in ("BL M", "BL mD", "SPF Mini", "GE", "unsupported"):
        session, game, sender, rec = _vacancy_setup(age_class)
        mv = _vacancy_member(session, game.team)
        db.set_team_mv(session, game.team, mv)
        # Retained synthetic staffing includes unknown birth dates/classifications.
        for role, slot in db.required_positions(game):
            helper = _vacancy_member(session, game.team, name=f"Legacy {role} {slot}", birth_date=None)
            session.add(db.Assignment(game=game, person=helper, role=role, slot=slot))
        session.commit()
        db.request_replacement(session, game, db.ROLE_SALE, 0, game.assignment_by_role(db.ROLE_SALE).person_id)
        session.commit()
        cake = next(b for b in db.get_day_blocks(session, h.SEASON, GAME_DATE) if b.phase == db.BLOCK_CAKE_DELIVERY)
        db.configure_cake_block(session, cake, "10:00", 2, expected_cake_quantity=None, expected_delivery_time=None)
        session.commit()
        assert not db.missing_slots(game) and db.staffing_status(game)["deficiencies"]
        assert sender.notify_team_vacancies_early(GAME_DATE) == 0
        assert not rec.mails and not rec.smss, "full staffing must not recruit for age, replacement or block gaps"
        assert sender.notify_pre(GAME_DATE) == len(db.required_positions(game)), "weekly assigned-helper reminders remain available"
        assert sender.notify_game_day(GAME_DATE) == len(db.required_positions(game)) + 1, "age deficiencies still trigger day-before MV follow-up"
        session.close()


if __name__ == "__main__":
    h.run_all(dict(globals()))
