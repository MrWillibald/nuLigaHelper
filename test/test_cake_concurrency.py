"""Cake capacity edits and claims share bounded SQLite serialization."""

import threading
from unittest.mock import patch

import helpers as h
import db


def _database(quantity=4):
    engine = h.make_engine()
    path = str(engine.url.database)
    with h.Session(engine) as session:
        record = h.sample_games()[0]
        db.sync_games(session, [record], h.SEASON)
        cake = next(block for block in db.get_day_blocks(
            session, h.SEASON, record["date"]
        ) if block.phase == db.BLOCK_CAKE_DELIVERY)
        people = [db.Person(name="First cake"), db.Person(name="Second cake")]
        session.add_all(people)
        session.flush()
        if quantity is not None:
            db.configure_cake_block(session, cake, "10:00", quantity, None, None)
        session.commit()
        ids = (cake.id, people[0].id, people[1].id)
    engine.dispose()
    return path, ids


def test_claim_retries_stale_snapshot_and_refuses_removed_position():
    path, (block_id, person_id, _other_id) = _database()
    reader_engine, writer_engine = db.make_engine(path), db.make_engine(path)
    try:
        with h.Session(reader_engine) as claimant:
            claimant.connection().exec_driver_sql("BEGIN")
            cake = claimant.get(db.DayBlock, block_id)
            person = claimant.get(db.Person, person_id)
            assert db.block_capacity(cake) == 4
            with h.Session(writer_engine) as editor:
                db.configure_cake_block(
                    editor, editor.get(db.DayBlock, block_id), "11:00", 2, "10:00", 4
                )
                editor.commit()
            try:
                db.claim_block_slot(claimant, cake, 3, None, person)
            except ValueError as refusal:
                assert "Aufgabenplatz" in str(refusal)
                claimant.rollback()
            else:
                raise AssertionError("A saved reduction must invalidate a stale claim.")
        with h.Session(writer_engine) as check:
            assert check.get(db.DayBlock, block_id).cake_quantity == 2
            assert check.query(db.BlockAssignment).count() == 0
            assert check.query(db.AssignmentAudit).count() == 0
    finally:
        reader_engine.dispose()
        writer_engine.dispose()


def test_reduction_retries_stale_snapshot_and_keeps_newly_occupied_position():
    path, (block_id, person_id, _other_id) = _database()
    reader_engine, writer_engine = db.make_engine(path), db.make_engine(path)
    try:
        with h.Session(reader_engine) as editor:
            editor.connection().exec_driver_sql("BEGIN")
            cake = editor.get(db.DayBlock, block_id)
            assert not cake.assignments
            with h.Session(writer_engine) as claimant:
                db.claim_block_slot(
                    claimant, claimant.get(db.DayBlock, block_id), 3, None,
                    claimant.get(db.Person, person_id),
                )
                claimant.commit()
            try:
                db.configure_cake_block(editor, cake, "11:00", 2, "10:00", 4)
            except ValueError as refusal:
                assert "freigegeben" in str(refusal)
                editor.rollback()
            else:
                raise AssertionError("A saved high-position occupant must prevent reduction.")
        with h.Session(writer_engine) as check:
            assert (check.get(db.DayBlock, block_id).delivery_time,
                    check.get(db.DayBlock, block_id).cake_quantity) == ("10:00", 4)
            assert check.query(db.BlockAssignment).one().slot == 3
            assert check.query(db.AssignmentAudit).count() == 1
    finally:
        reader_engine.dispose()
        writer_engine.dispose()


def test_competing_configuration_edits_preserve_winner_and_return_current_settings():
    path, (block_id, _first_id, _second_id) = _database(quantity=None)
    barrier = threading.Barrier(2)
    results = {}

    def edit(key, clock, quantity):
        engine = db.make_engine(path)
        try:
            with h.Session(engine) as session:
                session.connection().exec_driver_sql("BEGIN")
                cake = session.get(db.DayBlock, block_id)
                barrier.wait()
                try:
                    db.configure_cake_block(session, cake, clock, quantity, None, None)
                    session.commit()
                    results[key] = ("winner", clock, quantity)
                except db.CakeConfigurationConflictError as exc:
                    session.rollback()
                    results[key] = ("conflict", exc.current_delivery_time, exc.current_cake_quantity)
                except Exception as exc:
                    session.rollback()
                    results[key] = ("error", str(exc))
        finally:
            engine.dispose()

    threads = [threading.Thread(target=edit, args=args) for args in (
        ("first", "10:00", 4), ("second", "11:00", 2)
    )]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(8)
        assert not thread.is_alive(), "A configuration writer exceeded its deadline."
    winners = [value[1:] for value in results.values() if value[0] == "winner"]
    conflicts = [value[1:] for value in results.values() if value[0] == "conflict"]
    assert len(winners) == 1 and conflicts == winners, results
    engine = db.make_engine(path)
    try:
        with h.Session(engine) as session:
            cake = session.get(db.DayBlock, block_id)
            assert (cake.delivery_time, cake.cake_quantity) == winners[0]
            assert session.query(db.AssignmentAudit).count() == 0
    finally:
        engine.dispose()


def test_competing_cake_claims_keep_one_assignment_and_one_audit():
    path, (block_id, first_id, second_id) = _database()
    barrier = threading.Barrier(2)
    results = {}

    def claim(person_id):
        engine = db.make_engine(path)
        try:
            with h.Session(engine) as session:
                session.connection().exec_driver_sql("BEGIN")
                cake = session.get(db.DayBlock, block_id)
                person = session.get(db.Person, person_id)
                barrier.wait()
                try:
                    db.claim_block_slot(session, cake, 3, None, person)
                    session.commit()
                    results[person_id] = ("winner", person_id)
                except db.SlotConflictError as exc:
                    session.rollback()
                    results[person_id] = ("conflict", exc.current_person_id)
                except Exception as exc:
                    session.rollback()
                    results[person_id] = ("error", str(exc))
        finally:
            engine.dispose()

    threads = [threading.Thread(target=claim, args=(person_id,)) for person_id in (
        first_id, second_id
    )]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(8)
        assert not thread.is_alive(), "A cake claim writer exceeded its deadline."
    winners = [value[1] for value in results.values() if value[0] == "winner"]
    conflicts = [value[1] for value in results.values() if value[0] == "conflict"]
    assert len(winners) == 1 and conflicts == winners, results
    engine = db.make_engine(path)
    try:
        with h.Session(engine) as session:
            assert session.query(db.BlockAssignment).one().person_id == winners[0]
            assert session.query(db.AssignmentAudit).count() == 1
    finally:
        engine.dispose()


def test_configuration_writer_timeout_keeps_settings_and_assignment_history():
    path, (block_id, _first_id, _second_id) = _database()
    locker_engine, contender_engine = db.make_engine(path), db.make_engine(path)
    try:
        with locker_engine.connect() as locker:
            locker.exec_driver_sql("BEGIN IMMEDIATE")
            with h.Session(contender_engine) as contender:
                contender.connection().exec_driver_sql("PRAGMA busy_timeout=50")
                cake = contender.get(db.DayBlock, block_id)
                with patch("db.SQLITE_TIMEOUT_SECONDS", 0.15):
                    try:
                        db.configure_cake_block(contender, cake, "11:00", 2, "10:00", 4)
                    except db.AssignmentTemporarilyUnavailableError:
                        contender.rollback()
                    else:
                        raise AssertionError("Configuration must honor the bounded writer deadline.")
            locker.rollback()
        with h.Session(contender_engine) as check:
            cake = check.get(db.DayBlock, block_id)
            assert (cake.delivery_time, cake.cake_quantity) == ("10:00", 4)
            assert check.query(db.AssignmentAudit).count() == 0
    finally:
        locker_engine.dispose()
        contender_engine.dispose()


if __name__ == "__main__":
    h.run_all(dict(globals()))
