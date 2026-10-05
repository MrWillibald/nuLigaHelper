"""Cake reminders use synthetic transports, saved clocks and one-cake duties."""

from unittest.mock import patch

import helpers as h
import db
import notifier
from test_notifier import GAME_DATE, _setup


def _cake_scenario():
    session, game, sender, recorded = _setup(fully_assigned=True)
    cake = next(
        block for block in db.get_day_blocks(session, h.SEASON, GAME_DATE)
        if block.phase == db.BLOCK_CAKE_DELIVERY
    )
    db.configure_cake_block(session, cake, "10:15", 4, None, None)
    email = db.Person(name="Cake Mail", email="cake@example.test", phone="+491701111111")
    sms = db.Person(name="Cake SMS", phone="+491702222222")
    skipped = db.Person(name="Cake Without Contact")
    session.add_all([email, sms, skipped])
    session.flush()
    for slot, person in enumerate([email, sms, skipped]):
        db.claim_block_slot(session, cake, slot, None, person)
    session.commit()
    return session, cake, sender, recorded


def test_weekly_cake_reminders_identify_one_cake_saved_time_and_contact_preference():
    session, cake, sender, recorded = _cake_scenario()
    with patch("notifier.logging.warning") as warning:
        assert sender.notify_cakes_early(GAME_DATE) == 2, "contactless and empty positions send nothing"
    assert len(recorded.mails) == len(recorded.smss) == 1
    assert "cake@example.test" in recorded.mails[0][0]
    assert recorded.smss[0][0] == "+491702222222", "email remains preferred when both contacts exist"
    for body in [recorded.mails[0][2], recorded.smss[0][1]]:
        assert "Kuchenlieferung" in body and "ein Kuchen" in body
        assert GAME_DATE in body and "10:15" in body
        assert "nächste Woche" in body
    warning.assert_called_once_with("notification outcome=skipped reason=no_contact")
    session.close()


def test_day_before_cake_reminders_and_cake_gaps_never_create_mv_recipients():
    session, cake, sender, recorded = _cake_scenario()
    assert sender.notify_blocks_day_before(GAME_DATE) == 5, "three existing block helpers plus two cake contacts"
    cake_bodies = [body for _, _, body in recorded.mails if "ein Kuchen" in body]
    cake_bodies += [body for _, body in recorded.smss if "ein Kuchen" in body]
    assert len(cake_bodies) == 2
    assert all("morgen" in body and GAME_DATE in body and "10:15" in body for body in cake_bodies)
    recorded.mails.clear()
    recorded.smss.clear()
    assert sender.notify_game_day(GAME_DATE) == 6, "cake gaps do not make a fully staffed game's MV receive a reminder"
    assert all(number != "+491700000002" for number, _ in recorded.smss)
    session.close()


def test_unconfigured_and_zero_cakes_send_no_reminders():
    session, game, sender, recorded = _setup()
    cake = next(b for b in db.get_day_blocks(session, h.SEASON, GAME_DATE) if b.phase == db.BLOCK_CAKE_DELIVERY)
    assert sender.notify_cakes_early(GAME_DATE) == 0
    db.configure_cake_block(session, cake, "10:00", 0, None, None)
    session.commit()
    assert sender.notify_cakes_early(GAME_DATE) == 0
    assert not recorded.mails and not recorded.smss
    session.close()


def test_debug_cake_reminders_use_the_existing_provider_suppression():
    session, cake, sender, recorded = _cake_scenario()
    sender.send_Mail = notifier.Notifier.send_Mail.__get__(sender)
    sender.send_SMS = notifier.Notifier.send_SMS.__get__(sender)
    with patch("notifier.DEBUG_FLAG", True), patch("notifier.smtplib.SMTP_SSL") as smtp, patch("notifier.Client") as twilio:
        assert sender.notify_cakes_early(GAME_DATE) == 2
        smtp.assert_not_called()
        twilio.assert_not_called()
    session.close()


if __name__ == "__main__":
    h.run_all(dict(globals()))
