"""Synthetic domain coverage of continuous responsibility and reminder-event locking."""

from datetime import date, datetime
from unittest.mock import patch
from zoneinfo import ZoneInfo
from concurrent.futures import ThreadPoolExecutor
import threading

from tests import helpers as h
from nuligahelper import db
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


def test_ordner_transitions_capture_current_label_without_rewriting_old_audits():
    engine, game_id, people, _ = _seed()
    with h.Session(engine) as session:
        game = session.get(db.Game, game_id)
        first, second = [session.get(db.Person, p) for p in people[1:3]]
        legacy = db.AssignmentAudit(
            actor_tier="system", action="claim", game_id=game.id,
            role="Ordnungsdienst", slot=0, actor_name="System",
            affected_person_name=first.name, game_snapshot="Historical game snapshot",
        )
        session.add(legacy)
        session.commit()
        assignment = db.claim_slot(session, game, "Ordnungsdienst", 0, None, first)
        assignment_id = assignment.id
        token = db.request_replacement(session, game, db.ROLE_SECURITY, 0, first.id).replacement_token
        db.cancel_replacement(session, game, db.ROLE_SECURITY, 0, first.id, token)
        token = db.request_replacement(session, game, db.ROLE_SECURITY, 0, first.id).replacement_token
        db.takeover_slot(session, game, db.ROLE_SECURITY, 0, first.id, token, second)
        db.request_replacement(session, game, db.ROLE_SECURITY, 0, second.id)
        with patch("nuligahelper.db.clock.duty_now", return_value=datetime(2099, 12, 30, 9, tzinfo=ZoneInfo("Europe/Berlin"))):
            db.lock_day_before_batch(session, 2026, game.date)
        session.commit()
        session.expire_all()
        assert assignment.id == assignment_id and assignment.role == "Ordnungsdienst"
        assert assignment.person_id == second.id and assignment.locked_at is not None
        assert assignment.replacement_token is None
        assert legacy.role == "Ordnungsdienst" and legacy.game_snapshot == "Historical game snapshot"
        entries = session.query(db.AssignmentAudit).filter(db.AssignmentAudit.id != legacy.id).all()
        assert {entry.action for entry in entries} == {
            "claim", "request_open", "request_close", "takeover_out", "takeover_in", "lock"}
        assert all(entry.role == "Ordner" for entry in entries)
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
        with patch("nuligahelper.db.clock.duty_now", return_value=datetime(2099, 12, 30, 9, tzinfo=ZoneInfo("Europe/Berlin"))):
            marker = db.lock_day_before_batch(session, 2026, game.date)
        session.commit()  # No mail/SMS is attempted by a domain transition.
        assert assignment.locked_at == marker.locked_at and assignment.replacement_token is None
        assert all(a.locked_at == marker.locked_at and a.replacement_token is None for a in block_assignments)
        before = session.query(db.AssignmentAudit).count()
        with patch("nuligahelper.db.clock.duty_now", return_value=datetime(2099, 12, 30, 9, tzinfo=ZoneInfo("Europe/Berlin"))):
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
        with patch("nuligahelper.db.clock.duty_now", return_value=datetime(2099, 12, 30, 9, tzinfo=ZoneInfo("Europe/Berlin"))):
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
        with patch("nuligahelper.db.clock.duty_now", return_value=datetime(2099,12,30,8,59,tzinfo=ZoneInfo("Europe/Berlin"))):
            assert not db.duty_is_locked(session, game, assignment)
            assert db.lock_day_before_batch(session, 2026, game.date) is None
            assert session.query(db.DutyCutoff).count() == 0
        with patch("nuligahelper.db.clock.duty_now", return_value=datetime(2099,12,30,9,tzinfo=ZoneInfo("Europe/Berlin"))):
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
        with patch("nuligahelper.db.clock.duty_now", side_effect=lambda: clock[0]):
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
            with patch("nuligahelper.common.effective_today", return_value=date(2026, 9, 1)):
                _refused(lambda: db.deactivate_person(stale, person))
            stale.rollback()
            assert person.account_status == db.ACCOUNT_ACTIVE
            assignments = stale.get(container_type, container_id).assignments
            assert len(assignments) == 1 and assignments[0].person_id == person.id
        engine.dispose()


def _reopen_setup(session, game_id, people):
    """Return the game, the admin actor and a helper holding the timekeeper duty."""
    game, admin, helper = (session.get(db.Game, game_id), session.get(db.Person, people[0]),
                           session.get(db.Person, people[1]))
    db.claim_slot(session, game, db.ROLE_TIMEKEEPER, 0, None, helper)
    session.commit()
    return game, admin, helper


def _reopen_audits(session):
    return list(session.scalars(select(db.AssignmentAudit).where(db.AssignmentAudit.action == "reopen")))


def test_admin_reopens_an_unlocked_duty_and_only_a_reopen_transition_is_recorded():
    engine, game_id, people, _ = _seed()
    with h.Session(engine) as session:
        game, admin, helper = _reopen_setup(session, game_id, people)
        removed = db.reopen_slot(session, game, db.ROLE_TIMEKEEPER, 0, helper.id, admin, "admin")
        session.commit()
        assert removed.id == helper.id, "the caller needs the removed helper for the notification"
        assert not [a for a in game.assignments if a.role == db.ROLE_TIMEKEEPER]
        assert db.missing_slots(game).get(db.ROLE_TIMEKEEPER), "a reopened slot counts as vacant"
        assert not helper.assignments, "the helper must not keep a ghost assignment"
        audit, = _reopen_audits(session)
        assert (audit.actor_person_id, audit.affected_person_id, audit.actor_tier) == (admin.id, helper.id, "admin")
        assert session.query(db.AssignmentAudit).filter_by(action="release").count() == 0, \
            "reopening is not an ordinary release"
    engine.dispose()


def test_admin_reopens_a_locked_duty_and_other_locks_of_the_date_stay():
    engine, game_id, people, _ = _seed()
    with h.Session(engine) as session:
        game, admin, helper = _reopen_setup(session, game_id, people)
        other = session.get(db.Person, people[2])
        db.claim_slot(session, game, db.ROLE_SECRETARY, 0, None, other)
        session.commit()
        with _at(2099, 12, 30):
            db.reopen_slot(session, game, db.ROLE_TIMEKEEPER, 0, helper.id, admin, "admin")
            session.commit()
            kept, = game.assignments
            assert kept.person_id == other.id and kept.locked_at is not None, "other duties keep their lock"
            late = db.claim_slot(session, game, db.ROLE_TIMEKEEPER, 0, None, session.get(db.Person, people[3]))
            session.commit()
            assert late.locked_at is not None, "a late claim of a reopened slot locks immediately"
            _refused(lambda: db.release_slot(session, game, db.ROLE_SECRETARY, 0, other.id, admin, "admin"),
                     db.DutyLockedError)
    engine.dispose()


def test_reopening_closes_an_open_request_and_a_stale_takeover_is_refused():
    engine, game_id, people, _ = _seed()
    with h.Session(engine) as session:
        game, admin, helper = _reopen_setup(session, game_id, people)
        token = db.request_replacement(session, game, db.ROLE_TIMEKEEPER, 0, helper.id,
                                       actor=helper, actor_tier="member").replacement_token
        session.commit()
        db.reopen_slot(session, game, db.ROLE_TIMEKEEPER, 0, helper.id, admin, "admin")
        session.commit()
        actions = [a.action for a in session.scalars(select(db.AssignmentAudit).order_by(db.AssignmentAudit.id))]
        assert actions == ["claim", "request_open", "request_close", "reopen"]
        incoming = session.get(db.Person, people[2])
        _refused(lambda: db.takeover_slot(session, game, db.ROLE_TIMEKEEPER, 0, helper.id, token, incoming,
                                          actor=incoming, actor_tier="member"), db.SlotConflictError)
    engine.dispose()


def test_reopening_refuses_past_duties_and_permits_today_and_open_dates():
    engine, game_id, people, _ = _seed()
    with h.Session(engine) as session:
        game, admin, helper = _reopen_setup(session, game_id, people)
        with patch("nuligahelper.common.effective_today", return_value=date(2100, 1, 1)):
            _refused(lambda: db.reopen_slot(session, game, db.ROLE_TIMEKEEPER, 0, helper.id, admin, "admin"))
        session.rollback()
        assert [a.person_id for a in game.assignments] == [helper.id], "past work stays recorded"
        assert not _reopen_audits(session)
        with patch("nuligahelper.common.effective_today", return_value=date(2099, 12, 31)):
            db.reopen_slot(session, game, db.ROLE_TIMEKEEPER, 0, helper.id, admin, "admin")
        session.commit()
        db.claim_slot(session, game, db.ROLE_SECRETARY, 0, None, helper)
        game.date = "offen"
        session.commit()
        db.reopen_slot(session, game, db.ROLE_SECRETARY, 0, helper.id, admin, "admin")
        session.commit()
        assert not game.assignments
    engine.dispose()


def test_reopening_requires_a_current_admin():
    engine, game_id, people, _ = _seed()
    with h.Session(engine) as session:
        game, admin, helper = _reopen_setup(session, game_id, people)
        attempts = [(helper, "member"), (helper, "mv"), (helper, "admin"), (None, "system"), (None, "cli")]
        for actor, tier in attempts:
            _refused(lambda: db.reopen_slot(session, game, db.ROLE_TIMEKEEPER, 0, helper.id, actor, tier))
            session.rollback()
        admin.is_admin = False
        session.commit()
        _refused(lambda: db.reopen_slot(session, game, db.ROLE_TIMEKEEPER, 0, helper.id, admin, "admin"))
        session.rollback()
        assert [a.person_id for a in game.assignments] == [helper.id]
        assert not _reopen_audits(session)
    engine.dispose()


def test_reopening_names_the_expected_occupant():
    engine, game_id, people, _ = _seed()
    with h.Session(engine) as session:
        game, admin, helper = _reopen_setup(session, game_id, people)
        _refused(lambda: db.reopen_slot(session, game, db.ROLE_TIMEKEEPER, 0, people[2], admin, "admin"),
                 db.SlotConflictError)
        _refused(lambda: db.reopen_slot(session, game, db.ROLE_SECRETARY, 0, helper.id, admin, "admin"),
                 db.SlotConflictError)
        session.rollback()
        assert [a.person_id for a in game.assignments] == [helper.id], "a stale page cannot remove a newer helper"
        assert not _reopen_audits(session)
    engine.dispose()


def test_reopening_the_sole_adult_seller_keeps_the_collective_rule_for_the_last_slot():
    engine, game_id, people, _ = _seed()
    with h.Session(engine) as session:
        game, admin, adult = (session.get(db.Game, game_id), session.get(db.Person, people[0]),
                              session.get(db.Person, people[1]))
        minor = session.get(db.Person, people[2])
        other_minor = session.get(db.Person, people[3])
        minor.birth_date = other_minor.birth_date = date(2095, 1, 1)
        db.claim_slot(session, game, db.ROLE_SALE, 0, None, adult)
        db.claim_slot(session, game, db.ROLE_SALE, 1, None, minor)
        session.commit()
        db.reopen_slot(session, game, db.ROLE_SALE, 0, adult.id, admin, "admin")
        session.commit()
        assert [a.person_id for a in game.assignments] == [minor.id]
        _refused(lambda: db.claim_slot(session, game, db.ROLE_SALE, 0, None, other_minor))
    engine.dispose()



# ---------------------------------------------------------------
# Postponed and undated games release locks their date no longer justifies
# ---------------------------------------------------------------

_BERLIN = ZoneInfo("Europe/Berlin")
_FIRST_DATE, _LATER_DATE = "05.12.2026", "20.12.2026"


def _at(year, month, day, hour=10):
    return patch("nuligahelper.db.clock.duty_now",
                 return_value=datetime(year, month, day, hour, tzinfo=_BERLIN))


def _record(game_nr, game_date, time="18:00"):
    return {"day": "Sa", "date": game_date, "time": time, "hall": 1, "game_nr": game_nr,
            "ak": "BL M", "home": "TuS Raubling", "guest": "Gast", "score": ""}


def _sync(session, *records, season=2026):
    db.sync_games(session, list(records), season)


def _audit_actions(session, action):
    return list(session.scalars(select(db.AssignmentAudit).where(
        db.AssignmentAudit.action == action).order_by(db.AssignmentAudit.id)))


def _postponement_seed():
    """Game 9101 is moved in the tests; 9102 keeps the original date and its blocks."""
    engine = h.make_engine()
    with h.Session(engine) as session:
        people = [db.Person(name=f"Helper {i}", birth_date=date(1990, 1, 1),
                            account_status=db.ACCOUNT_ACTIVE) for i in range(5)]
        session.add_all(people); session.commit()
        _sync(session, _record("9101", _FIRST_DATE), _record("9102", _FIRST_DATE, "20:00"))
        return engine, [p.id for p in people]


def _game(session, game_nr="9101"):
    return session.scalars(select(db.Game).where(db.Game.game_nr == game_nr)).one()


def _locked_game_duty(session, people, role=db.ROLE_SECURITY):
    """Claim a duty before cutoff and lock it at the original 09:00 day-before cutoff."""
    with _at(2026, 12, 1):
        assignment = db.claim_slot(session, _game(session), role, 0, None,
                                   session.get(db.Person, people[1]))
        session.commit()
    with _at(2026, 12, 4, 9):
        db.lock_due_duties(session); session.commit()
    assert assignment.locked_at is not None
    return assignment


def test_postponement_to_a_later_open_date_releases_locks_without_vacating_duties():
    engine, people = _postponement_seed()
    with h.Session(engine) as session:
        game = _game(session)
        preparation = next(block for block in db.get_day_blocks(session, 2026, _FIRST_DATE)
                           if block.phase == db.BLOCK_PREPARATION)
        with _at(2026, 12, 1):
            ordner = db.claim_slot(session, game, db.ROLE_SECURITY, 0, None, session.get(db.Person, people[1]))
            timing = db.claim_slot(session, game, db.ROLE_TIMEKEEPER, 0, None, session.get(db.Person, people[2]))
            db.request_replacement(session, game, db.ROLE_SECURITY, 0, people[1])
            block_duty = db.claim_block_slot(session, preparation, 0, None, session.get(db.Person, people[1]))
            session.commit()
        with _at(2026, 12, 4, 9):
            db.lock_due_duties(session); session.commit()
        with _at(2026, 12, 5, 8):
            late = db.claim_slot(session, game, db.ROLE_CASH, 0, None, session.get(db.Person, people[3]))
            session.commit()
        assert all(a.locked_at is not None for a in (ordner, timing, late, block_duty))
        assert ordner.replacement_token is None, "the original cutoff closes the open search"

        with _at(2026, 12, 6):
            _sync(session, _record("9101", _LATER_DATE), _record("9102", _FIRST_DATE, "20:00"))
            assert [a.locked_at for a in (ordner, timing, late)] == [None, None, None], \
                "a later, not yet due date must release every saved game-duty lock, including late claims"
            assert (ordner.person_id, timing.person_id, late.person_id) == (people[1], people[2], people[3])
            assert ordner.replacement_token is None, "a search closed by the old lock must not reopen"
            assert block_duty.locked_at is not None, "day blocks stay with their date and keep their lock"
            unlocks = _audit_actions(session, "unlock")
            assert sorted(a.affected_person_id for a in unlocks) == sorted(people[1:4])
            assert all(a.actor_tier == "system" and a.actor_person_id is None for a in unlocks)
            assert all(_LATER_DATE in a.game_snapshot for a in unlocks)

            _sync(session, _record("9101", _LATER_DATE), _record("9102", _FIRST_DATE, "20:00"))
            assert len(_audit_actions(session, "unlock")) == 3, "a repeated sync must not duplicate unlock evidence"

            request = db.request_replacement(session, _game(session), db.ROLE_SECURITY, 0, people[1])
            db.takeover_slot(session, _game(session), db.ROLE_SECURITY, 0, people[1],
                             request.replacement_token, session.get(db.Person, people[4]))
            session.commit()
            assert ordner.person_id == people[4], "unlocked duties follow ordinary replacement rules"

        locks_before = len(_audit_actions(session, "lock"))
        with _at(2026, 12, 19, 9):
            db.lock_due_duties(session); session.commit()
        assert all(a.locked_at is not None for a in (ordner, timing, late)), "the new cutoff locks again"
        assert len(_audit_actions(session, "lock")) == locks_before + 3
    engine.dispose()


def test_locks_stay_when_the_current_date_still_justifies_them():
    cases = {
        "later date whose cutoff is already due": ((2026, 12, 10), "08.12.2026", "18:00", False),
        "later date with a saved cutoff marker": ((2026, 12, 6), _LATER_DATE, "18:00", True),
        "same date with a new start time": ((2026, 12, 6), _FIRST_DATE, "16:00", False),
        "earlier date": ((2026, 12, 4), "04.12.2026", "18:00", False),
    }
    for label, (now, new_date, new_time, marker) in cases.items():
        engine, people = _postponement_seed()
        with h.Session(engine) as session:
            assignment = _locked_game_duty(session, people)
            if marker:
                session.add(db.DutyCutoff(season_year=2026, date=new_date, locked_at=datetime(2026, 12, 19, 9)))
                session.commit()
            with _at(*now):
                _sync(session, _record("9101", new_date, new_time), _record("9102", _FIRST_DATE, "20:00"))
            assert assignment.locked_at is not None, f"{label}: the saved lock must remain"
            assert not _audit_actions(session, "unlock"), f"{label}: no unlock may be recorded"
        engine.dispose()


def test_open_or_unparseable_dates_release_locks():
    for open_date in ("", "offen"):
        engine, people = _postponement_seed()
        with h.Session(engine) as session:
            assignment = _locked_game_duty(session, people)
            with _at(2026, 12, 6):
                _sync(session, _record("9101", open_date), _record("9102", _FIRST_DATE, "20:00"))
                assert assignment.locked_at is None, f"date {open_date!r} has no cutoff and cannot keep a lock"
                assert assignment.person_id == people[1]
                assert db.request_replacement(session, _game(session), db.ROLE_SECURITY, 0, people[1]).replacement_token
            assert len(_audit_actions(session, "unlock")) == 1
        engine.dispose()


def test_stored_open_date_lock_is_released_once_and_valid_dates_govern_again():
    engine, people = _postponement_seed()
    with h.Session(engine) as session:
        assignment = _locked_game_duty(session, people)
        with _at(2026, 12, 6):
            _sync(session, _record("9101", "offen"), _record("9102", _FIRST_DATE, "20:00"))
            assignment.locked_at = datetime(2026, 12, 4, 9)  # Lock saved while the date was already open.
            session.commit()
            _sync(session, _record("9101", "offen"), _record("9102", _FIRST_DATE, "20:00"))
            assert assignment.locked_at is None, "a stored open date must release a lock without a date change"
            _sync(session, _record("9101", "offen"), _record("9102", _FIRST_DATE, "20:00"))
            assert len(_audit_actions(session, "unlock")) == 2

            assignment.locked_at = datetime(2026, 12, 4, 9)
            session.commit()
            _sync(session, _record("9101", _LATER_DATE), _record("9102", _FIRST_DATE, "20:00"))
            assert assignment.locked_at is None, "open to a later not yet due date releases a remaining lock"

        with _at(2026, 12, 10):
            _sync(session, _record("9101", "offen"), _record("9102", _FIRST_DATE, "20:00"))
            _sync(session, _record("9101", "08.12.2026"), _record("9102", _FIRST_DATE, "20:00"))
            game = _game(session)
            assert db.duty_is_locked(session, game, assignment), "a valid, already due date locks again"
            db.lock_due_duties(session); session.commit()
            assert assignment.locked_at is not None
    engine.dispose()


def test_unlocked_open_date_duty_keeps_invalid_date_age_rules():
    engine, people = _postponement_seed()
    with h.Session(engine) as session:
        timing = _locked_game_duty(session, people, db.ROLE_TIMEKEEPER)
        with _at(2026, 12, 6):
            _sync(session, _record("9101", "offen"), _record("9102", _FIRST_DATE, "20:00"))
            assert timing.locked_at is None
            request = db.request_replacement(session, _game(session), db.ROLE_TIMEKEEPER, 0, people[1])
            token = request.replacement_token
            session.commit()
            _refused(lambda: db.takeover_slot(session, _game(session), db.ROLE_TIMEKEEPER, 0, people[1],
                                              token, session.get(db.Person, people[4])),
                     db.AgeEligibilityError)
            session.rollback()
            assert timing.person_id == people[1], "an invalid game date still cannot authorize a timing takeover"
    engine.dispose()


def test_first_sync_after_release_repairs_games_moved_before_it():
    engine, people = _postponement_seed()
    with h.Session(engine) as session:
        with _at(2026, 12, 1):
            _sync(session, _record("9101", _FIRST_DATE), _record("9102", _FIRST_DATE, "20:00"),
                  _record("9103", "10.12.2026"))
            _sync(session, _record("9301", "06.12.2025"), season=2025)
        stale = {}
        for game_nr, season in (("9101", 2026), ("9102", 2026), ("9103", 2026), ("9301", 2025)):
            game = session.scalars(select(db.Game).where(db.Game.game_nr == game_nr,
                                                         db.Game.season_year == season)).one()
            with _at(2026, 12, 1):
                stale[game_nr] = db.claim_slot(session, game, db.ROLE_SECURITY, 0, None,
                                               session.get(db.Person, people[1]))
        session.commit()
        # State as left by the previous release: the games were already moved, but
        # their duties still carry the original 04.12. cutoff and its lock audits.
        old_cutoff = datetime(2026, 12, 4, 9)
        for game_nr, assignment in stale.items():
            assignment.locked_at = old_cutoff
            db.audit._audit_assignment(session, assignment, "lock", None, "system")
        _game(session, "9101").date = _LATER_DATE          # still in the scrape
        _game(session, "9103").date = "21.12.2026"          # no longer in the scrape
        session.commit()
        old_audits = [(a.id, a.action, a.game_snapshot) for a in
                      session.scalars(select(db.AssignmentAudit).order_by(db.AssignmentAudit.id))]

        with _at(2026, 12, 6):
            _sync(session, _record("9101", _LATER_DATE), _record("9102", _FIRST_DATE, "20:00"))
        assert stale["9101"].locked_at is None and stale["9103"].locked_at is None, \
            "duties of games moved before this release must be repaired by the first sync"
        assert stale["9102"].locked_at == old_cutoff, "a game whose own cutoff is due keeps its lock"
        assert stale["9301"].locked_at == old_cutoff, "other seasons are not part of this sync"
        assert all(a.person_id == people[1] for a in stale.values())
        unlocks = _audit_actions(session, "unlock")
        assert sorted(a.game_id for a in unlocks) == sorted([stale["9101"].game_id, stale["9103"].game_id])
        assert [(a.id, a.action, a.game_snapshot) for a in session.scalars(
            select(db.AssignmentAudit).where(db.AssignmentAudit.id <= old_audits[-1][0])
            .order_by(db.AssignmentAudit.id))] == old_audits, "historical audits stay unchanged"
    engine.dispose()


if __name__ == "__main__":
    h.run_all(dict(globals()))
