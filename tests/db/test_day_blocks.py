"""Automatic day-block lifecycle, timing, assignment and audit behavior."""

from datetime import date, datetime
from unittest.mock import patch
from zoneinfo import ZoneInfo

import pytest
from sqlalchemy.exc import IntegrityError

from tests import helpers as h
from nuligahelper import db


def _one_game(date="05.09.2026", time="00:30", number="1001"):
    return {
        "day": "Sa", "date": date, "time": time, "hall": 280340,
        "game_nr": number, "ak": "BL mD", "home": "TuS Raubling",
        "guest": "Gast", "score": "",
    }


def test_sync_creates_exact_blocks_reuses_them_and_calculates_midnight_offsets():
    engine = h.make_engine()
    with h.Session(engine) as session:
        db.sync_games(session, [_one_game()], h.SEASON)
        blocks = db.get_day_blocks(session, h.SEASON, "05.09.2026")
        assert [block.phase for block in blocks] == list(db.BLOCK_PHASES)
        original_ids = [block.id for block in blocks]
        assert db.calculated_block_time(blocks[0]) == datetime(2026, 9, 4, 23, 0)
        assert db.calculated_block_time(blocks[-1]) == datetime(2026, 9, 5, 1, 30)

        later = _one_game(time="22:30", number="1002")
        db.sync_games(session, [_one_game(), later], h.SEASON)
        blocks = db.get_day_blocks(session, h.SEASON, "05.09.2026")
        assert [block.id for block in blocks] == original_ids
        assert db.calculated_block_time(blocks[-1]) == datetime(2026, 9, 5, 23, 30)


def test_invalid_boundary_times_leave_both_blocks_available_without_invented_time():
    engine = h.make_engine()
    with h.Session(engine) as session:
        db.sync_games(session, [_one_game(time="unbekannt")], h.SEASON)
        blocks = db.get_day_blocks(session, h.SEASON, "05.09.2026")
        assert len(blocks) == 3
        assert all(db.calculated_block_time(block) is None for block in blocks)


def test_block_claim_release_enforces_container_scope_and_writes_atomic_audits():
    engine = h.make_engine()
    with h.Session(engine) as session:
        db.sync_games(session, [_one_game()], h.SEASON)
        preparation, _cake, cleanup = db.get_day_blocks(session, h.SEASON, "05.09.2026")
        assert preparation.label == "Vorbereitung"
        assert cleanup.label == "Aufräumen"
        assert db.block_slot_label(preparation, 1) == "Vorbereitung 2"
        person = db.Person(name="Helfer", email="helper@example.test")
        session.add(person)
        session.flush()

        db.claim_block_slot(session, preparation, 0, None, person)
        with pytest.raises(ValueError):
            db.claim_block_slot(session, preparation, 1, None, person)
        db.claim_block_slot(session, cleanup, 0, None, person)
        with pytest.raises(db.SlotConflictError):
            db.release_block_slot(session, preparation, 0, None)
        assert session.query(db.AssignmentAudit).count() == 2
        db.request_block_replacement(session, preparation, 0, person.id)
        session.commit()
        audits = session.query(db.AssignmentAudit).order_by(db.AssignmentAudit.id).all()
        assert [entry.action for entry in audits] == ["claim", "claim", "request_open"]
        assert all(entry.block_snapshot for entry in audits)
        assert all(entry.game_id is None for entry in audits)
        assert [entry.block_snapshot for entry in audits] == [
            "Saison 2026 | 05.09.2026 | Vorbereitung 1",
            "Saison 2026 | 05.09.2026 | Aufräumen 1",
            "Saison 2026 | 05.09.2026 | Vorbereitung 1",
        ], "Every new day-block audit records the season and original named slot."


def test_reconcile_and_sync_refuse_occupied_date_removal_without_partial_changes():
    engine = h.make_engine()
    with h.Session(engine) as session:
        db.sync_games(session, [_one_game()], h.SEASON)
        block = db.get_day_blocks(session, h.SEASON, "05.09.2026")[0]
        person = db.Person(name="Helfer", email="helper@example.test")
        session.add(person)
        session.flush()
        db.claim_block_slot(session, block, 0, None, person)
        session.commit()

        request = db.request_block_replacement(session, block, 0, person.id)
        token = request.replacement_token
        session.commit()
        before = session.query(db.AssignmentAudit).count()
        original_ids = [b.id for b in db.get_day_blocks(session, h.SEASON, block.date)]
        with pytest.raises(ValueError, match="verbindlichen"):
            db.reconcile_day_blocks(session, h.SEASON, set())
        session.rollback()
        assert [b.id for b in db.get_day_blocks(session, h.SEASON, block.date)] == original_ids
        assert block.assignment_for_slot(0).person_id == person.id
        assert block.assignment_for_slot(0).replacement_token == token
        assert session.query(db.AssignmentAudit).count() == before
        assert session.query(db.AssignmentAudit).filter_by(action="remove").count() == 0


def test_successful_sync_removes_vanished_blocks_logs_each_and_invalid_sync_does_not():
    engine = h.make_engine()
    with h.Session(engine) as session:
        db.sync_games(session, [_one_game()], h.SEASON)
        original_ids = {
            block.id for block in db.get_day_blocks(session, h.SEASON, "05.09.2026")
        }
        duplicate = [_one_game(number="same"), _one_game(number="same")]
        with pytest.raises(ValueError):
            db.sync_games(session, duplicate, h.SEASON)
        assert original_ids == {
            block.id for block in db.get_day_blocks(session, h.SEASON, "05.09.2026")
        }

        replacement = _one_game(date="06.09.2026", number="2001")
        with patch("nuligahelper.db.logging.info") as info:
            db.sync_games(session, [replacement], h.SEASON)
        assert not db.get_day_blocks(session, h.SEASON, "05.09.2026")
        removals = [
            call for call in info.call_args_list
            if call.args and call.args[0].startswith("day_block_removed")
        ]
        assert len(removals) == 3
        assert all(call.args[-1] == 0 for call in removals)


def test_retained_cash_duty_is_preserved_but_not_required_on_youth():
    engine = h.make_engine()
    with h.Session(engine) as session:
        db.sync_games(session, [_one_game()], h.SEASON)
        game = session.query(db.Game).one()
        people = []
        for index, (role, slot) in enumerate(db.required_positions(game)):
            person = db.Person(name=f"Helfer {index}-{slot}", birth_date=h.ADULT_BIRTH_DATE)
            session.add(person)
            session.flush()
            db.claim_slot(session, game, role, slot, None, person)
            people.append(person)
        assert db.missing_slots(game) == {}
        optional = db.Person(name="Optional")
        session.add(optional)
        session.flush()
        session.add(db.Assignment(game=game, role=db.ROLE_CASH, slot=0, person=optional))
        session.flush()
        assert game.assignment_by_role(db.ROLE_CASH).person is optional


def test_database_constraints_reject_duplicate_block_occupants():
    engine = h.make_engine()
    with h.Session(engine) as session:
        db.sync_games(session, [_one_game()], h.SEASON)
        block = db.get_day_blocks(session, h.SEASON, "05.09.2026")[0]
        first = db.Person(name="First")
        second = db.Person(name="Second")
        session.add_all([first, second])
        session.flush()
        session.add_all([
            db.BlockAssignment(block=block, person=first, slot=0),
            db.BlockAssignment(block=block, person=second, slot=0),
        ])
        with pytest.raises(IntegrityError):
            session.flush()


def test_admin_reopens_a_locked_block_duty_and_a_late_claim_locks_immediately():
    engine = h.make_engine()
    with h.Session(engine) as session:
        db.sync_games(session, [_one_game()], h.SEASON)
        block = db.get_day_blocks(session, h.SEASON, "05.09.2026")[0]
        admin = db.Person(name="Admin", birth_date=h.ADULT_BIRTH_DATE,
                          account_status=db.ACCOUNT_ACTIVE, is_admin=True)
        helper, late = (db.Person(name=name, birth_date=h.ADULT_BIRTH_DATE, account_status=db.ACCOUNT_ACTIVE)
                        for name in ("Block helper", "Late helper"))
        session.add_all([admin, helper, late])
        session.flush()
        db.claim_block_slot(session, block, 0, None, helper)
        session.commit()
        cutoff = db.duty_cutoff_at("05.09.2026")
        today = cutoff.date()
        with (patch("nuligahelper.db.clock.duty_now", return_value=cutoff),
              patch("nuligahelper.common.effective_today", return_value=today)):
            with pytest.raises(ValueError):
                db.release_block_slot(session, block, 0, helper.id, admin, "admin")
            session.rollback()
            removed = db.reopen_block_slot(session, block, 0, helper.id, admin, "admin")
            session.commit()
            assert removed.id == helper.id and block.assignment_for_slot(0) is None
            assert not helper.block_assignments
            audit = session.query(db.AssignmentAudit).filter_by(action="reopen").one()
            assert (audit.actor_person_id, audit.affected_person_id) == (admin.id, helper.id)
            assert audit.block_id == block.id and audit.game_id is None
            claimed = db.claim_block_slot(session, block, 0, None, late)
            session.commit()
            assert claimed.locked_at is not None, "a late claim of a reopened slot locks immediately"


def test_block_reopening_refuses_other_tiers_past_dates_and_stale_occupants():
    engine = h.make_engine()
    with h.Session(engine) as session:
        db.sync_games(session, [_one_game()], h.SEASON)
        block = db.get_day_blocks(session, h.SEASON, "05.09.2026")[0]
        admin = db.Person(name="Admin", birth_date=h.ADULT_BIRTH_DATE,
                          account_status=db.ACCOUNT_ACTIVE, is_admin=True)
        helper = db.Person(name="Block helper", birth_date=h.ADULT_BIRTH_DATE, account_status=db.ACCOUNT_ACTIVE)
        session.add_all([admin, helper])
        session.flush()
        db.claim_block_slot(session, block, 0, None, helper)
        session.commit()
        before = datetime(2026, 9, 1).date()
        with patch("nuligahelper.common.effective_today", return_value=before):
            with pytest.raises(ValueError, match="Nur Admins"):
                db.reopen_block_slot(session, block, 0, helper.id, helper, "member")
            session.rollback()
            with pytest.raises(db.SlotConflictError):
                db.reopen_block_slot(session, block, 0, admin.id, admin, "admin")
            session.rollback()
        with patch("nuligahelper.common.effective_today", return_value=datetime(2026, 9, 6).date()):
            with pytest.raises(ValueError, match="Vergangene"):
                db.reopen_block_slot(session, block, 0, helper.id, admin, "admin")
            session.rollback()
        assert block.assignment_for_slot(0).person_id == helper.id
        assert session.query(db.AssignmentAudit).filter_by(action="reopen").count() == 0


def _before(day=date(2026, 9, 1)):
    return patch("nuligahelper.common.effective_today", return_value=day)


def _active_person(session, name):
    person = db.Person(name=name, birth_date=h.ADULT_BIRTH_DATE, account_status=db.ACCOUNT_ACTIVE)
    session.add(person)
    session.flush()
    return person


def _open_blocks(session, person=None, date_text="offen"):
    """Build blocks as an older release left them for a placeholder date."""
    blocks = [db.DayBlock(season_year=h.SEASON, date=date_text, phase=phase) for phase in db.BLOCK_PHASES]
    session.add_all(blocks)
    session.flush()
    if person is not None:
        session.add(db.BlockAssignment(block=blocks[0], person=person, slot=0))
    session.commit()
    return blocks


def test_only_calendar_dates_are_block_dates():
    assert db.block_date_valid("05.09.2026")
    for value in ("", None, "offen", "31.02.2026", "2026-09-05"):
        assert not db.block_date_valid(value), f"{value!r} is an open date"


def test_open_date_games_get_no_blocks_until_a_calendar_date_is_known():
    with h.Session(h.make_engine()) as session:
        for open_date in ("offen", ""):
            db.sync_games(session, [_one_game(date=open_date)], h.SEASON)
            assert session.query(db.DayBlock).count() == 0
            assert session.query(db.Game).one().date == open_date, "the game itself stays"
        db.sync_games(session, [_one_game()], h.SEASON)
        assert len(db.get_day_blocks(session, h.SEASON, "05.09.2026")) == len(db.BLOCK_PHASES)


def test_empty_open_date_blocks_are_removed_and_occupied_ones_retained_until_emptied():
    with h.Session(h.make_engine()) as session:
        helper = _active_person(session, "Retained helper")
        admin = _active_person(session, "Admin")
        admin.is_admin = True
        preparation = _open_blocks(session, helper)[0]
        with _before():
            events = db.sync_games(session, [_one_game(date="offen"), _one_game(number="1002")], h.SEASON)
            assert events.cancelled_block_duties == [], "nothing is removed from an open date"
            assert [b.id for b in db.get_day_blocks(session, h.SEASON, "offen")] == [preparation.id], \
                "empty siblings go, the occupied block stays"
            assert preparation.assignment_for_slot(0).person_id == helper.id
            assert len(db.get_day_blocks(session, h.SEASON, "05.09.2026")) == len(db.BLOCK_PHASES), \
                "other dates are synchronized normally"
            db.sync_games(session, [_one_game(number="1002")], h.SEASON)
            assert db.get_day_blocks(session, h.SEASON, "offen"), "the block outlives its games"
            db.reopen_block_slot(session, preparation, 0, helper.id, admin, "admin")
            session.commit()
            db.sync_games(session, [_one_game(number="1002")], h.SEASON)
            assert not db.get_day_blocks(session, h.SEASON, "offen")
            audits = session.query(db.AssignmentAudit).count()
            db.sync_games(session, [_one_game(number="1002")], h.SEASON)
            assert session.query(db.AssignmentAudit).count() == audits, "a repeated sync changes nothing"


def test_retained_open_date_block_offers_no_claims_but_keeps_replacement_rules():
    with h.Session(h.make_engine()) as session:
        helper, replacement = _active_person(session, "Original"), _active_person(session, "Replacement")
        block = _open_blocks(session, helper)[0]
        before = session.query(db.AssignmentAudit).count()
        with pytest.raises(ValueError, match="kein Termin"):
            db.claim_block_slot(session, block, 1, None, replacement)
        session.rollback()
        assert block.assignment_for_slot(1) is None
        assert session.query(db.AssignmentAudit).count() == before
        token = db.request_block_replacement(session, block, 0, helper.id, actor=helper,
                                             actor_tier="member").replacement_token
        db.takeover_block_slot(session, block, 0, helper.id, token, replacement,
                               actor=replacement, actor_tier="member")
        session.commit()
        assert block.assignment_for_slot(0).person_id == replacement.id


def test_last_game_leaving_a_future_date_cancels_block_duties_with_evidence():
    for scraped, game_date in (([_one_game(date="06.09.2026")], "06.09.2026"),
                               ([_one_game(date="offen")], "offen"),
                               ([_one_game(date="06.09.2026", number="2001")], "05.09.2026")):
        with h.Session(h.make_engine()) as session, _before():
            db.sync_games(session, [_one_game()], h.SEASON)
            first, second = _active_person(session, "First"), _active_person(session, "Second")
            preparation, _cake, cleanup = db.get_day_blocks(session, h.SEASON, "05.09.2026")
            db.claim_block_slot(session, preparation, 1, None, first)
            db.claim_block_slot(session, cleanup, 0, None, first)
            db.claim_block_slot(session, cleanup, 1, None, second)
            db.request_block_replacement(session, cleanup, 1, second.id, actor=second, actor_tier="member")
            session.commit()
            events = db.sync_games(session, scraped, h.SEASON)
            assert not db.get_day_blocks(session, h.SEASON, "05.09.2026")
            assert session.query(db.BlockAssignment).count() == 0 and not first.block_assignments
            assert sorted((d.person_id, d.block_date, d.task_label) for d in events.cancelled_block_duties) == sorted([
                (first.id, "05.09.2026", db.block_slot_label(db.BLOCK_PREPARATION, 1)),
                (first.id, "05.09.2026", db.block_slot_label(db.BLOCK_CLEANUP, 0)),
                (second.id, "05.09.2026", db.block_slot_label(db.BLOCK_CLEANUP, 1)),
            ])
            removals = session.query(db.AssignmentAudit).filter_by(action="date_remove").all()
            assert len(removals) == 3 and {a.actor_tier for a in removals} == {"system"}
            assert all("05.09.2026" in a.block_snapshot for a in removals), "snapshots outlive the block"
            closures = session.query(db.AssignmentAudit).filter_by(action="request_close").all()
            assert [a.affected_person_id for a in closures] == [second.id]
            assert {game.date for game in session.query(db.Game)} >= {game_date}, \
                "the game update is saved with the cancellation; a vanished game stays stored"
            assert db.sync_games(session, scraped, h.SEASON).cancelled_block_duties == []


def test_locked_duty_of_a_cancelled_date_is_removed_and_a_returning_date_starts_empty():
    with h.Session(h.make_engine()) as session:
        with _before():
            db.sync_games(session, [_one_game()], h.SEASON)
            helper = _active_person(session, "Locked helper")
            block = db.get_day_blocks(session, h.SEASON, "05.09.2026")[0]
            db.claim_block_slot(session, block, 0, None, helper)
            session.commit()
        cutoff = db.duty_cutoff_at("05.09.2026")
        with patch("nuligahelper.db.clock.duty_now", return_value=cutoff):
            db.lock_day_before_batch(session, h.SEASON, "05.09.2026")
            session.commit()
            events = db.sync_games(session, [_one_game(date="06.09.2026")], h.SEASON)
            assert [d.person_id for d in events.cancelled_block_duties] == [helper.id]
            db.sync_games(session, [_one_game()], h.SEASON)
        returned = db.get_day_blocks(session, h.SEASON, "05.09.2026")
        assert returned and not any(b.assignments for b in returned), "removed duties are not restored"


def test_today_is_cancelled_but_past_dates_keep_their_occupied_blocks():
    for today, removed in ((date(2026, 9, 5), True), (date(2026, 9, 6), False)):
        with h.Session(h.make_engine()) as session:
            with _before():
                db.sync_games(session, [_one_game()], h.SEASON)
                helper = _active_person(session, "Past helper")
                block_id = db.get_day_blocks(session, h.SEASON, "05.09.2026")[0].id
                db.claim_block_slot(session, session.get(db.DayBlock, block_id), 0, None, helper)
                session.commit()
            clock = datetime(today.year, today.month, today.day, 12, tzinfo=ZoneInfo("Europe/Berlin"))
            with patch("nuligahelper.db.clock.duty_now", return_value=clock):
                events = db.sync_games(session, [_one_game(date="07.09.2026")], h.SEASON)
            assert bool(events.cancelled_block_duties) is removed
            assert (session.get(db.DayBlock, block_id) is None) is removed
            assert len(helper.block_assignments) == (0 if removed else 1), "past work stays counted"


def test_empty_scrape_is_refused_for_a_season_with_games_and_allowed_for_an_empty_one():
    with h.Session(h.make_engine()) as session:
        assert db.sync_games(session, [], h.SEASON) == db.SyncEvents()
        with _before():
            db.sync_games(session, [_one_game()], h.SEASON)
            helper = _active_person(session, "Scrape helper")
            block = db.get_day_blocks(session, h.SEASON, "05.09.2026")[0]
            db.claim_block_slot(session, block, 0, None, helper)
            session.commit()
            before = session.query(db.AssignmentAudit).count()
            with pytest.raises(ValueError, match="keine Spiele"):
                db.sync_games(session, [], h.SEASON)
            session.rollback()
        assert block.assignment_for_slot(0).person_id == helper.id
        assert session.query(db.AssignmentAudit).count() == before


if __name__ == "__main__":
    h.run_all(dict(globals()))
