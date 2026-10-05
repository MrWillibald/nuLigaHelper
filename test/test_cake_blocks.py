"""Cake configuration, variable capacity and retained dated assignment history."""

from datetime import datetime

import pytest
from sqlalchemy.exc import IntegrityError

import helpers as h
import db


def _setup(session):
    game_data = h.sample_games()[0]
    db.sync_games(session, [game_data], h.SEASON)
    preparation, cake, cleanup = db.get_day_blocks(
        session, h.SEASON, game_data["date"]
    )
    people = [db.Person(name=f"Cake helper {index}") for index in range(5)]
    session.add_all(people)
    session.flush()
    return game_data, preparation, cake, cleanup, people


def test_new_cake_requires_setup_and_configured_zero_is_distinct():
    with h.Session(h.make_engine()) as session:
        _data, preparation, cake, cleanup, people = _setup(session)
        assert cake.phase == db.BLOCK_CAKE_DELIVERY
        assert cake.label == "Kuchenlieferung"
        assert cake.delivery_time is None and cake.cake_quantity is None
        assert not db.block_is_configured(cake)
        assert db.block_capacity(cake) == 0
        assert db.calculated_block_time(cake) is None
        assert db.block_capacity(preparation) == db.block_capacity(cleanup) == 3
        with pytest.raises(ValueError, match="eingerichtet"):
            db.claim_block_slot(session, cake, 0, None, people[0])
        db.configure_cake_block(session, cake, "10:00", 0, None, None)
        assert db.block_is_configured(cake), "Zero cakes is a saved configuration."
        assert db.block_capacity(cake) == 0
        assert db.calculated_block_time(cake) == datetime(2026, 10, 3, 10, 0)
        with pytest.raises(ValueError, match="Aufgabenplatz"):
            db.claim_block_slot(session, cake, 0, None, people[0])
        assert session.query(db.AssignmentAudit).count() == 0


def test_four_cakes_allow_four_slots_but_keep_one_task_per_block():
    with h.Session(h.make_engine()) as session:
        _data, preparation, cake, cleanup, people = _setup(session)
        db.configure_cake_block(session, cake, "10:00", 4, None, None)
        assert db.block_capacity(cake) == 4
        for slot, person in enumerate(people[:4]):
            db.claim_block_slot(session, cake, slot, None, person)
        assert db.block_slot_label(cake, 3) == "Kuchenlieferung 4"
        with pytest.raises(ValueError, match="Aufgabenplatz"):
            db.claim_block_slot(session, cake, 4, None, people[4])
        with pytest.raises(ValueError, match="Aufgabenplatz"):
            db.claim_block_slot(session, preparation, 3, None, people[4])
        db.release_block_slot(session, cake, 1, people[1].id)
        with pytest.raises(ValueError, match="bereits"):
            db.claim_block_slot(session, cake, 1, None, people[0])
        db.claim_block_slot(session, preparation, 0, None, people[0])
        db.claim_block_slot(session, cleanup, 0, None, people[0])
        game = session.query(db.Game).one()
        db.claim_slot(session, game, db.ROLE_SECURITY, 0, None, people[0])
        assert len(cake.assignments) == 3


def test_quantity_edits_preserve_numbered_occupants_and_require_explicit_release():
    with h.Session(h.make_engine()) as session:
        _data, _preparation, cake, _cleanup, people = _setup(session)
        db.configure_cake_block(session, cake, "10:00", 2, None, None)
        first = db.claim_block_slot(session, cake, 0, None, people[0])
        first_id = first.id
        db.configure_cake_block(session, cake, "11:00", 4, "10:00", 2)
        assert cake.assignment_for_slot(0).id == first_id
        assert cake.assignment_for_slot(2) is None
        db.claim_block_slot(session, cake, 3, None, people[1])
        before_audits = session.query(db.AssignmentAudit).count()
        with pytest.raises(ValueError, match="ausdrücklich freigegeben"):
            db.configure_cake_block(session, cake, "12:00", 2, "11:00", 4)
        assert (cake.delivery_time, cake.cake_quantity) == ("11:00", 4)
        assert session.query(db.AssignmentAudit).count() == before_audits
        assert cake.assignment_for_slot(3).person_id == people[1].id
        db.release_block_slot(session, cake, 3, people[1].id)
        db.configure_cake_block(session, cake, "12:00", 2, "11:00", 4)
        assert cake.assignment_for_slot(0).id == first_id
        assert cake.assignment_for_slot(1) is None
        release = session.query(db.AssignmentAudit).filter_by(action="release").one()
        assert release.role == "Kuchenlieferung" and release.slot == 3
        assert "Kuchenlieferung 4" in release.block_snapshot
        assert str(h.SEASON) in release.block_snapshot
        assert cake.date in release.block_snapshot
        assert db.block_slot_label(db.BLOCK_CAKE_DELIVERY, release.slot) == "Kuchenlieferung 4"
        with pytest.raises(ValueError):
            db.configure_cake_block(session, cake, "12:00", 0, "12:00", 2)
        db.release_block_slot(session, cake, 0, people[0].id)
        db.configure_cake_block(session, cake, "12:00", 0, "12:00", 2)
        assert db.block_capacity(cake) == 0 and db.block_is_configured(cake)


def test_invalid_settings_and_stale_edits_leave_saved_settings_and_history_unchanged():
    with h.Session(h.make_engine()) as session:
        _data, preparation, cake, _cleanup, _people = _setup(session)
        db.configure_cake_block(session, cake, "10:00", 4, None, None)
        invalid_times = (None, "", "9:00", "24:00", "10:60", "10:00 ", "٠١:٠١")
        for clock in invalid_times:
            with pytest.raises(ValueError):
                db.configure_cake_block(session, cake, clock, 4, "10:00", 4)
        for quantity in (None, -1, 4.0, 1.5, True, "4", 2**63):
            with pytest.raises(ValueError):
                db.configure_cake_block(session, cake, "10:00", quantity, "10:00", 4)
        with pytest.raises(db.CakeConfigurationConflictError) as conflict:
            db.configure_cake_block(session, cake, "11:00", 3, None, None)
        assert conflict.value.current_delivery_time == "10:00"
        assert conflict.value.current_cake_quantity == 4
        with pytest.raises(ValueError):
            db.configure_cake_block(session, preparation, "10:00", 4, None, None)
        assert (cake.delivery_time, cake.cake_quantity) == ("10:00", 4)
        assert session.query(db.AssignmentAudit).count() == 0


def test_same_date_sync_keeps_cake_configuration_and_moved_date_removes_with_audit():
    with h.Session(h.make_engine()) as session:
        data, _preparation, cake, _cleanup, people = _setup(session)
        db.configure_cake_block(session, cake, "10:00", 4, None, None)
        assignment = db.claim_block_slot(session, cake, 3, None, people[0])
        block_id, assignment_id = cake.id, assignment.id
        session.commit()
        changed = {**data, "time": "22:30"}
        extra = {**data, "game_nr": "9999", "time": "08:00"}
        db.sync_games(session, [changed, extra], h.SEASON)
        retained = session.get(db.DayBlock, block_id)
        assert (retained.delivery_time, retained.cake_quantity) == ("10:00", 4)
        assert retained.assignment_for_slot(3).id == assignment_id
        assert db.calculated_block_time(retained) == datetime(2026, 10, 3, 10, 0)
        with pytest.raises(ValueError):
            db.sync_games(session, [changed, changed], h.SEASON)
        assert session.get(db.DayBlock, block_id) is retained
        moved = {**data, "date": "04.10.2026"}
        db.sync_games(session, [moved], h.SEASON)
        session.commit()
        removal = session.query(db.AssignmentAudit).filter_by(action="remove").one()
        assert removal.block_id is None
        assert removal.block_snapshot == "Saison 2026 | 03.10.2026 | Kuchenlieferung 4"
        assert removal.actor_tier == "system"
        new_cake = next(block for block in db.get_day_blocks(
            session, h.SEASON, moved["date"]
        ) if block.phase == db.BLOCK_CAKE_DELIVERY)
        assert new_cake.id != block_id
        assert new_cake.delivery_time is None and new_cake.cake_quantity is None
        assert not new_cake.assignments


def test_inactive_people_cannot_claim_cakes_and_stale_release_preserves_occupant():
    with h.Session(h.make_engine()) as session:
        _data, _preparation, cake, _cleanup, people = _setup(session)
        db.configure_cake_block(session, cake, "10:00", 4, None, None)
        people[0].account_status = db.ACCOUNT_INACTIVE
        session.flush()
        with pytest.raises(ValueError, match="nicht eingeteilt"):
            db.claim_block_slot(session, cake, 0, None, people[0])
        db.claim_block_slot(session, cake, 0, None, people[1])
        with pytest.raises(db.SlotConflictError) as conflict:
            db.release_block_slot(session, cake, 0, people[0].id)
        assert conflict.value.current_person_id == people[1].id
        assert cake.assignment_for_slot(0).person_id == people[1].id
        assert session.query(db.AssignmentAudit).count() == 1


def test_metadata_checks_reject_negative_quantity_and_non_cake_settings():
    for phase, quantity in ((db.BLOCK_CAKE_DELIVERY, -1), (db.BLOCK_PREPARATION, 2)):
        with h.Session(h.make_engine()) as session:
            session.add(db.DayBlock(
                season_year=h.SEASON, date="01.11.2026", phase=phase,
                cake_quantity=quantity, delivery_time="10:00",
            ))
            with pytest.raises(IntegrityError):
                session.flush()
            session.rollback()


if __name__ == "__main__":
    h.run_all(dict(globals()))
