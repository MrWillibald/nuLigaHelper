"""Synthetic domain coverage of continuous responsibility and reminder-event locking."""

from datetime import date, datetime
from unittest.mock import patch
from zoneinfo import ZoneInfo
from concurrent.futures import ThreadPoolExecutor
import threading

import helpers as h
import db
from sqlalchemy import select


def _seed():
    engine = h.make_engine()
    with h.Session(engine) as session:
        people = [db.Person(name=f"Helper {i}", birth_date=date(1990, 1, 1),
                            account_status=db.ACCOUNT_ACTIVE, is_admin=i == 0)
                  for i in range(5)]
        game = db.Game(season_year=2026, game_nr="9001", date="31.12.2099", ak="M")
        session.add_all([*people, game]); session.flush()
        blocks = [db.DayBlock(season_year=2026, date=game.date, phase=phase)
                  for phase in db.BLOCK_PHASES]
        session.add_all(blocks); session.flush()
        cake = next(block for block in blocks if block.phase == db.BLOCK_CAKE_DELIVERY)
        db.configure_cake_block(session, cake, "10:00", 3, None, None)
        session.commit()
        return engine, game.id, [p.id for p in people], [b.id for b in blocks]


def _refused(call, error=ValueError):
    try:
        call()
    except error:
        return
    raise AssertionError(f"Expected {error.__name__}; a refused mutation must retain responsibility")


def test_request_and_takeover_keep_position_identity_and_atomic_before_after_audits():
    engine, game_id, people, _ = _seed()
    with h.Session(engine) as session:
        game, first, second = session.get(db.Game, game_id), session.get(db.Person, people[1]), session.get(db.Person, people[2])
        assignment = db.claim_slot(session, game, db.ROLE_TIMEKEEPER, 0, None, first)
        original_id = assignment.id
        requested = db.request_replacement(session, game, db.ROLE_TIMEKEEPER, 0, first.id,
                                            actor=first, actor_tier="member")
        token = requested.replacement_token
        assert token and requested.person_id == first.id and requested.id == original_id
        assert not db.missing_slots(game).get(db.ROLE_TIMEKEEPER)
        db.request_replacement(session, game, db.ROLE_TIMEKEEPER, 0, first.id, actor=first, actor_tier="member")
        assert requested.replacement_token == token, "repeated opening must not replace a request generation"
        db.takeover_slot(session, game, db.ROLE_TIMEKEEPER, 0, first.id, token, second,
                         actor=second, actor_tier="member")
        session.commit()
        assert requested.person_id == second.id and requested.id == original_id
        assert requested.replacement_token is None and requested.locked_at is None
        audits = list(session.scalars(select(db.AssignmentAudit).order_by(db.AssignmentAudit.id)))
        assert [a.action for a in audits] == ["claim", "request_open", "takeover_out", "request_close", "takeover_in"]
        assert [a.affected_person_id for a in audits[-3:]] == [first.id, first.id, second.id]
        assert all("1990" not in (a.game_snapshot or "") for a in audits)
    engine.dispose()


def test_takeover_rechecks_final_sale_group_and_preserves_original_on_refusal():
    engine, game_id, people, _ = _seed()
    with h.Session(engine) as session:
        game = session.get(db.Game, game_id)
        adult, young, other = [session.get(db.Person, p) for p in people[1:4]]
        young.birth_date = other.birth_date = date(2090, 1, 1)
        first = db.claim_slot(session, game, db.ROLE_SALE, 0, None, adult)
        db.claim_slot(session, game, db.ROLE_SALE, 1, None, young)
        db.request_replacement(session, game, db.ROLE_SALE, 0, adult.id)
        token = first.replacement_token
        before = len(list(session.scalars(select(db.AssignmentAudit))))
        _refused(lambda: db.takeover_slot(session, game, db.ROLE_SALE, 0, adult.id, token, other), db.AgeEligibilityError)
        assert first.person_id == adult.id and first.replacement_token == token
        assert len(list(session.scalars(select(db.AssignmentAudit)))) == before
        _refused(lambda: db.release_slot(session, game, db.ROLE_SALE, 0, adult.id))
    engine.dispose()


def test_batch_locks_complete_cohort_and_late_empty_claims_without_delivery_dependency():
    engine, game_id, people, blocks = _seed()
    with h.Session(engine) as session:
        game = session.get(db.Game, game_id)
        person = session.get(db.Person, people[1])  # Deliberately contactless.
        assignment = db.claim_slot(session, game, db.ROLE_SECURITY, 0, None, person)
        db.request_replacement(session, game, db.ROLE_SECURITY, 0, person.id)
        block_assignments = []
        for block_id in blocks:
            block = session.get(db.DayBlock, block_id)
            block_assignments.append(db.claim_block_slot(session, block, 0, None, person))
            db.request_block_replacement(session, block, 0, person.id)
        with patch("db.duty_now", return_value=datetime(2099, 12, 30, 9, tzinfo=ZoneInfo("Europe/Berlin"))):
            marker = db.lock_day_before_batch(session, 2026, game.date)
        session.commit()  # No mail/SMS is attempted by a domain transition.
        assert assignment.locked_at == marker.locked_at and assignment.replacement_token is None
        assert all(a.locked_at == marker.locked_at and a.replacement_token is None for a in block_assignments)
        before = session.query(db.AssignmentAudit).count()
        with patch("db.duty_now", return_value=datetime(2099, 12, 30, 9, tzinfo=ZoneInfo("Europe/Berlin"))):
            db.lock_day_before_batch(session, 2026, game.date)
        session.commit()
        assert session.query(db.AssignmentAudit).count() == before, "batch retry must not duplicate lifecycle audit"
        late = db.claim_slot(session, game, db.ROLE_CLEANING, 0, None, session.get(db.Person, people[2]))
        assert late.locked_at == marker.locked_at
        for block_id in blocks:
            late_block = db.claim_block_slot(session, session.get(db.DayBlock, block_id), 1, None, session.get(db.Person, people[2]))
            assert late_block.locked_at == marker.locked_at
        session.commit()
        _refused(lambda: db.request_replacement(session, game, db.ROLE_SECURITY, 0, person.id), db.DutyLockedError)
        _refused(lambda: db.release_slot(session, game, db.ROLE_SECURITY, 0, person.id,
                                        session.get(db.Person, people[0]), "admin"), db.DutyLockedError)
        game.date = "30.12.2099"; session.commit()
        assert db.duty_is_locked(session, game, assignment), "rescheduling must retain an existing commitment lock"
    engine.dispose()


def test_block_takeovers_obey_capacity_uniqueness_and_member_authority():
    engine, _, people, blocks = _seed()
    with h.Session(engine) as session:
        block = session.get(db.DayBlock, blocks[0])
        first, second, third = [session.get(db.Person, p) for p in people[1:4]]
        assignment = db.claim_block_slot(session, block, 0, None, first)
        db.claim_block_slot(session, block, 1, None, third)
        db.request_block_replacement(session, block, 0, first.id, first, "member")
        token = assignment.replacement_token
        _refused(lambda: db.takeover_block_slot(session, block, 0, first.id, token, third))
        _refused(lambda: db.takeover_block_slot(session, block, 0, first.id, token, second, first, "member"))
        assert assignment.person_id == first.id and assignment.replacement_token == token
        db.takeover_block_slot(session, block, 0, first.id, token, second, second, "member")
        assert assignment.person_id == second.id and assignment.replacement_token is None
    engine.dispose()


def test_stale_request_generation_cannot_authorize_takeover_of_same_person():
    engine, game_id, people, _ = _seed()
    with h.Session(engine) as session:
        game = session.get(db.Game, game_id)
        first, second = [session.get(db.Person, p) for p in people[1:3]]
        assignment = db.claim_slot(session, game, db.ROLE_SECURITY, 0, None, first)
        db.request_replacement(session, game, db.ROLE_SECURITY, 0, first.id)
        stale_token = assignment.replacement_token
        assignment.replacement_token = "synthetic-new-request-generation"
        session.commit()
        _refused(lambda: db.takeover_slot(session, game, db.ROLE_SECURITY, 0, first.id, stale_token, second), db.SlotConflictError)
        assert assignment.person_id == first.id
        assert assignment.replacement_token == "synthetic-new-request-generation"
    engine.dispose()


def test_competing_takeovers_have_one_winner_and_no_partial_vacancy():
    engine, game_id, people, _ = _seed()
    with h.Session(engine) as session:
        game = session.get(db.Game, game_id)
        first = session.get(db.Person, people[1])
        db.claim_slot(session, game, db.ROLE_SECURITY, 0, None, first)
        token = db.request_replacement(session, game, db.ROLE_SECURITY, 0, first.id).replacement_token
        session.commit()
    start = threading.Barrier(2)
    def take(person_id):
        with h.Session(engine) as session:
            game, person = session.get(db.Game, game_id), session.get(db.Person, person_id)
            start.wait()
            try:
                db.takeover_slot(session, game, db.ROLE_SECURITY, 0, people[1], token, person)
                session.commit(); return "winner"
            except db.SlotConflictError:
                session.rollback(); return "conflict"
    with ThreadPoolExecutor(max_workers=2) as pool:
        outcomes = list(pool.map(take, people[2:4]))
    assert sorted(outcomes) == ["conflict", "winner"]
    with h.Session(engine) as session:
        assignment = session.scalars(select(db.Assignment).where(db.Assignment.game_id == game_id)).one()
        assert assignment.person_id in people[2:4]
        assert session.query(db.AssignmentAudit).filter_by(action="takeover_in").count() == 1
        assert session.query(db.AssignmentAudit).filter_by(action="takeover_out").count() == 1
    engine.dispose()


def test_legacy_retirement_is_prelock_only_and_management_cannot_cancel_commitments():
    engine, game_id, people, blocks = _seed()
    with h.Session(engine) as session:
        game = session.get(db.Game, game_id)
        person, admin = session.get(db.Person, people[1]), session.get(db.Person, people[0])
        old = db.claim_slot(session, game, db.ROLE_CASH, 0, None, person)
        game.ak = "mD"; session.commit()
        db.release_slot(session, game, db.ROLE_CASH, 0, person.id); session.commit()
        assert not game.assignments
        game.ak = "M"; session.commit()
        old = db.claim_slot(session, game, db.ROLE_CASH, 0, None, person)
        game.ak = "mD"; session.commit()
        with patch("db.duty_now", return_value=datetime(2099, 12, 30, 9, tzinfo=ZoneInfo("Europe/Berlin"))):
            db.lock_day_before_batch(session, 2026, game.date)
        session.commit()
        _refused(lambda: db.release_slot(session, game, db.ROLE_CASH, 0, person.id), db.DutyLockedError)
        before = session.query(db.AssignmentAudit).count()
        _refused(lambda: db.deactivate_person(session, person, admin, "admin"))
        _refused(lambda: db.delete_person(session, person, admin, "admin"))
        assert game.assignments and person.account_status == db.ACCOUNT_ACTIVE
        assert session.query(db.AssignmentAudit).count() == before
    engine.dispose()


def test_fixed_berlin_cutoff_materializes_without_a_reminder_job_and_not_before_nine():
    engine, game_id, people, _ = _seed()
    with h.Session(engine) as session:
        game, person = session.get(db.Game, game_id), session.get(db.Person, people[1])
        assignment = db.claim_slot(session, game, db.ROLE_SECURITY, 0, None, person)
        db.request_replacement(session, game, db.ROLE_SECURITY, 0, person.id)
        session.commit()
        with patch("db.duty_now", return_value=datetime(2099,12,30,8,59,tzinfo=ZoneInfo("Europe/Berlin"))):
            assert not db.duty_is_locked(session, game, assignment)
            assert db.lock_day_before_batch(session, 2026, game.date) is None
            assert session.query(db.DutyCutoff).count() == 0
        with patch("db.duty_now", return_value=datetime(2099,12,30,9,tzinfo=ZoneInfo("Europe/Berlin"))):
            assert db.duty_is_locked(session, game, assignment), "fixed cutoff cannot wait for a cron or provider"
            db.lock_due_duties(session); session.commit()
            assert assignment.locked_at is not None and assignment.replacement_token is None
            _refused(lambda: db.request_replacement(session, game, db.ROLE_SECURITY, 0, person.id), db.DutyLockedError)
            late = db.claim_slot(session, game, db.ROLE_CLEANING, 0, None, session.get(db.Person, people[2]))
            assert late.locked_at is not None
    engine.dispose()


def test_cutoff_uses_berlin_calendar_day_across_daylight_saving_changes():
    spring = db.duty_cutoff_at("30.03.2026")
    autumn = db.duty_cutoff_at("26.10.2026")
    assert (spring.day, spring.hour, spring.utcoffset().total_seconds()) == (29, 9, 7200)
    assert (autumn.day, autumn.hour, autumn.utcoffset().total_seconds()) == (25, 9, 3600)
    assert db.duty_cutoff_at("invalid") is None


def test_takeover_waiting_for_writer_rechecks_nine_oclock_before_changing_responsibility():
    from sqlalchemy import event
    engine, game_id, people, _ = _seed()
    with h.Session(engine) as session:
        game = session.get(db.Game, game_id)
        owner = session.get(db.Person, people[1])
        db.claim_slot(session, game, db.ROLE_SECURITY, 0, None, owner)
        token = db.request_replacement(session, game, db.ROLE_SECURITY, 0, owner.id).replacement_token
        session.commit()
    clock = [datetime(2099, 12, 30, 8, 59, tzinfo=ZoneInfo("Europe/Berlin"))]
    attempted = threading.Event()
    main_thread = threading.get_ident()
    def observe(connection, cursor, statement, parameters, context, many):
        if threading.get_ident() != main_thread and statement.startswith("UPDATE games"):
            attempted.set()
    event.listen(engine, "before_cursor_execute", observe)
    def waiting_takeover():
        with h.Session(engine) as session:
            try:
                db.takeover_slot(session, session.get(db.Game, game_id), db.ROLE_SECURITY, 0,
                                 people[1], token, session.get(db.Person, people[2]))
                session.commit()
                return "transferred"
            except (db.DutyLockedError, db.SlotConflictError):
                session.rollback()
                return "refused"
    try:
        with patch("db.duty_now", side_effect=lambda: clock[0]):
            with ThreadPoolExecutor(max_workers=1) as pool:
                with h.Session(engine) as writer:
                    db._lock_game(writer, game_id)
                    waiting = pool.submit(waiting_takeover)
                    assert attempted.wait(3), "takeover did not reach the contended writer"
                    clock[0] = datetime(2099, 12, 30, 9, tzinfo=ZoneInfo("Europe/Berlin"))
                    writer.commit()
                assert waiting.result(timeout=8) == "refused", "a pre-cutoff request must not bypass the cutoff after waiting"
            with h.Session(engine) as session:
                db.lock_due_duties(session); session.commit()
                assignment = session.query(db.Assignment).filter_by(game_id=game_id).one()
                assert assignment.person_id == people[1] and assignment.locked_at is not None
                assert assignment.replacement_token is None
                assert session.query(db.AssignmentAudit).filter_by(action="takeover_in").count() == 0
    finally:
        event.remove(engine, "before_cursor_execute", observe)
        engine.dispose()


def test_stale_empty_block_relationship_cannot_authorize_removal_after_a_committed_claim():
    engine, _, people, blocks = _seed()
    with h.Session(engine) as stale:
        block = stale.get(db.DayBlock, blocks[0])
        assert block.assignments == []
        with h.Session(engine) as writer:
            db.claim_block_slot(writer, writer.get(db.DayBlock, block.id), 0, None,
                                writer.get(db.Person, people[1]))
            writer.commit()
        _refused(lambda: db.reconcile_day_blocks(stale, 2026, set()))
        stale.rollback()
        assert stale.get(db.DayBlock, block.id).assignment_for_slot(0).person_id == people[1]
        assert stale.query(db.DayBlock).count() == len(blocks)
        assert stale.query(db.AssignmentAudit).filter_by(action="remove").count() == 0
    engine.dispose()


def test_deactivation_rechecks_saved_dates_instead_of_cached_past_containers():
    for is_block in (False, True):
        engine, game_id, people, blocks = _seed()
        container_type = db.DayBlock if is_block else db.Game
        container_id = blocks[0] if is_block else game_id
        with h.Session(engine) as session:
            person = session.get(db.Person, people[1])
            container = session.get(container_type, container_id)
            if is_block:
                db.claim_block_slot(session, container, 0, None, person)
            else:
                db.claim_slot(session, container, db.ROLE_SECURITY, 0, None, person)
            container.date = "31.08.2026"
            session.commit()
        with h.Session(engine) as stale:
            person = stale.get(db.Person, people[1])
            container = stale.get(container_type, container_id)
            assert container.date == "31.08.2026"
            with h.Session(engine) as writer:
                writer.get(container_type, container_id).date = "31.12.2099"
                writer.commit()
            with patch("common.effective_today", return_value=date(2026, 9, 1)):
                _refused(lambda: db.deactivate_person(stale, person))
            stale.rollback()
            assert person.account_status == db.ACCOUNT_ACTIVE
            assignments = stale.get(container_type, container_id).assignments
            assert len(assignments) == 1 and assignments[0].person_id == person.id
        engine.dispose()


if __name__ == "__main__":
    h.run_all(dict(globals()))
