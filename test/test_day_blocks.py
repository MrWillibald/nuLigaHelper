"""Automatic day-block lifecycle, timing, assignment and audit behavior."""

from datetime import datetime
from unittest.mock import patch

import pytest
from sqlalchemy.exc import IntegrityError

import helpers as h
import db


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
        assert db.calculated_block_time(blocks[1]) == datetime(2026, 9, 5, 1, 30)

        later = _one_game(time="22:30", number="1002")
        db.sync_games(session, [_one_game(), later], h.SEASON)
        blocks = db.get_day_blocks(session, h.SEASON, "05.09.2026")
        assert [block.id for block in blocks] == original_ids
        assert db.calculated_block_time(blocks[1]) == datetime(2026, 9, 5, 23, 30)


def test_invalid_boundary_times_leave_both_blocks_available_without_invented_time():
    engine = h.make_engine()
    with h.Session(engine) as session:
        db.sync_games(session, [_one_game(time="unbekannt")], h.SEASON)
        blocks = db.get_day_blocks(session, h.SEASON, "05.09.2026")
        assert len(blocks) == 2
        assert all(db.calculated_block_time(block) is None for block in blocks)


def test_block_claim_release_enforces_container_scope_and_writes_atomic_audits():
    engine = h.make_engine()
    with h.Session(engine) as session:
        db.sync_games(session, [_one_game()], h.SEASON)
        preparation, cleanup = db.get_day_blocks(session, h.SEASON, "05.09.2026")
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
        db.release_block_slot(session, preparation, 0, person.id)
        session.commit()
        audits = session.query(db.AssignmentAudit).order_by(db.AssignmentAudit.id).all()
        assert [entry.action for entry in audits] == ["claim", "claim", "release"]
        assert all(entry.block_snapshot for entry in audits)
        assert all(entry.game_id is None for entry in audits)


def test_reconcile_removed_date_audits_occupied_slot_and_keeps_snapshot():
    engine = h.make_engine()
    with h.Session(engine) as session:
        db.sync_games(session, [_one_game()], h.SEASON)
        block = db.get_day_blocks(session, h.SEASON, "05.09.2026")[0]
        person = db.Person(name="Helfer", email="helper@example.test")
        session.add(person)
        session.flush()
        db.claim_block_slot(session, block, 0, None, person)
        session.commit()

        db.reconcile_day_blocks(session, h.SEASON, set())
        session.commit()
        removal = session.query(db.AssignmentAudit).filter_by(action="remove").one()
        assert removal.block_id is None
        assert "05.09.2026" in removal.block_snapshot
        assert session.query(db.DayBlock).count() == 0


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
        with patch("db.logging.info") as info:
            db.sync_games(session, [replacement], h.SEASON)
        assert not db.get_day_blocks(session, h.SEASON, "05.09.2026")
        removals = [
            call for call in info.call_args_list
            if call.args and call.args[0].startswith("day_block_removed")
        ]
        assert len(removals) == 2
        assert all(call.args[-1] == 0 for call in removals)


def test_optional_support_role_is_assignable_but_not_missing():
    engine = h.make_engine()
    with h.Session(engine) as session:
        db.sync_games(session, [_one_game()], h.SEASON)
        game = session.query(db.Game).one()
        people = []
        for index, (role, count) in enumerate(db.REQUIRED_ROLE_SLOT_COUNT.items()):
            for slot in range(count):
                person = db.Person(name=f"Helfer {index}-{slot}")
                session.add(person)
                session.flush()
                db.claim_slot(session, game, role, slot, None, person)
                people.append(person)
        assert db.missing_slots(game) == {}
        optional = db.Person(name="Optional")
        session.add(optional)
        session.flush()
        db.claim_slot(session, game, db.ROLE_SUPPORT, 0, None, optional)
        assert game.assignment_by_role(db.ROLE_SUPPORT).person is optional


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


if __name__ == "__main__":
    h.run_all(dict(globals()))
