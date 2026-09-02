"""WAL contention keeps assignment compare-and-swap and audit semantics intact."""

import os
import tempfile
import threading
import time
import uuid

import helpers as h
import db
import webapp


_TEST_DIR = tempfile.mkdtemp(prefix="nuligahelper-concurrency-")


def _new_database():
    path = os.path.join(_TEST_DIR, uuid.uuid4().hex + ".db")
    engine = db.make_engine(path)
    db.init_db(engine)
    with h.Session(engine) as setup:
        games = h.sync_sample_games(setup)
        game = setup.query(db.Game).filter_by(
            source_key=games[0]["source_key"]
        ).one()
        people = [db.Person(name=name) for name in ("First", "Second", "Third")]
        setup.add_all(people)
        setup.commit()
        ids = (game.id, *(person.id for person in people))
    engine.dispose()
    return path, ids


def _engine(path):
    return db.make_engine(path)


def _claim(path, game_id, person_id, role, slot, barrier, results, key):
    engine = _engine(path)
    try:
        with h.Session(engine) as session:
            session.connection().exec_driver_sql("BEGIN")
            game = session.get(db.Game, game_id)
            person = session.get(db.Person, person_id)
            barrier.wait()
            try:
                db.claim_slot(session, game, role, slot, None, person)
                session.commit()
                results[key] = ("winner", person_id)
            except db.SlotConflictError as exc:
                session.rollback()
                results[key] = ("conflict", exc.current_person_id)
            except ValueError as exc:
                session.rollback()
                results[key] = ("person_conflict", str(exc))
            except Exception as exc:
                session.rollback()
                results[key] = ("error", exc)
    finally:
        engine.dispose()


def test_wal_reader_can_overlap_an_independent_writer():
    path, _ids = _new_database()
    reader_engine = _engine(path)
    writer_engine = _engine(path)
    try:
        with h.Session(reader_engine) as reader:
            reader.connection().exec_driver_sql("BEGIN")
            before = reader.query(db.Person).count()
            with h.Session(writer_engine) as writer:
                writer.add(db.Person(name="Committed Writer"))
                writer.commit()
            assert reader.query(db.Person).count() == before
            reader.rollback()
            assert reader.query(db.Person).count() == before + 1
    finally:
        reader_engine.dispose()
        writer_engine.dispose()


def test_writer_lock_that_clears_within_timeout_allows_claim_and_audit():
    path, (game_id, first_id, _second_id, _third_id) = _new_database()
    locker_engine = _engine(path)
    worker_engine = _engine(path)
    started = threading.Event()
    result = {}

    with locker_engine.connect() as locker:
        locker.exec_driver_sql("BEGIN IMMEDIATE")
        locker.exec_driver_sql(
            "INSERT INTO persons (name, is_admin, account_status) "
            "VALUES ('Lock holder', 0, 'active')"
        )

        def waiting_claim():
            try:
                with h.Session(worker_engine) as session:
                    game = session.get(db.Game, game_id)
                    person = session.get(db.Person, first_id)
                    started.set()
                    db.claim_slot(
                        session, game, db.ROLE_SALE, 0, None, person
                    )
                    session.commit()
                    result["status"] = "winner"
            except Exception as exc:
                result["error"] = exc

        thread = threading.Thread(target=waiting_claim)
        thread.start()
        assert started.wait(1), "waiting writer did not start"
        time.sleep(0.25)
        locker.commit()
        thread.join(7)
        assert not thread.is_alive(), "waiting writer exceeded the busy timeout"

    worker_engine.dispose()
    locker_engine.dispose()
    assert result == {"status": "winner"}, result
    check_engine = _engine(path)
    try:
        with h.Session(check_engine) as session:
            assignment = session.query(db.Assignment).filter_by(
                game_id=game_id, role=db.ROLE_SALE, slot=0
            ).one()
            assert assignment.person_id == first_id
            assert session.query(db.AssignmentAudit).filter_by(
                game_id=game_id, action="claim"
            ).count() == 1
    finally:
        check_engine.dispose()


def test_writer_lock_beyond_timeout_is_temporarily_unavailable_without_audit():
    path, (game_id, first_id, _second_id, _third_id) = _new_database()
    locker_engine = _engine(path)
    contender_engine = _engine(path)
    started = time.monotonic()
    try:
        with locker_engine.connect() as locker:
            locker.exec_driver_sql("BEGIN IMMEDIATE")
            locker.exec_driver_sql(
                "INSERT INTO persons (name, is_admin, account_status) "
                "VALUES ('Long lock', 0, 'active')"
            )
            with h.Session(contender_engine) as contender:
                game = contender.get(db.Game, game_id)
                person = contender.get(db.Person, first_id)
                try:
                    db.claim_slot(
                        contender, game, db.ROLE_SALE, 0, None, person
                    )
                except db.AssignmentTemporarilyUnavailableError:
                    contender.rollback()
                else:
                    raise AssertionError("a five-second writer lock must time out")
            elapsed = time.monotonic() - started
            assert 4.5 <= elapsed < 7.0, elapsed
            locker.rollback()
    finally:
        contender_engine.dispose()
        locker_engine.dispose()

    check_engine = _engine(path)
    try:
        with h.Session(check_engine) as session:
            assert session.query(db.Assignment).filter_by(game_id=game_id).count() == 0
            assert session.query(db.AssignmentAudit).filter_by(game_id=game_id).count() == 0
    finally:
        check_engine.dispose()


def test_concurrent_stale_claims_commit_exactly_one_winner_and_audit():
    path, (game_id, first_id, second_id, _third_id) = _new_database()
    barrier = threading.Barrier(2)
    results = {}
    threads = [
        threading.Thread(
            target=_claim,
            args=(
                path,
                game_id,
                person_id,
                db.ROLE_SALE,
                0,
                barrier,
                results,
                str(person_id),
            ),
        )
        for person_id in (first_id, second_id)
    ]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(8)
        assert not thread.is_alive(), "concurrent claim did not finish"

    winners = [value[1] for value in results.values() if value[0] == "winner"]
    conflicts = [value[1] for value in results.values() if value[0] == "conflict"]
    assert len(winners) == 1 and conflicts == winners, results

    engine = _engine(path)
    try:
        with h.Session(engine) as session:
            assignments = session.query(db.Assignment).filter_by(
                game_id=game_id, role=db.ROLE_SALE, slot=0
            ).all()
            assert len(assignments) == 1 and assignments[0].person_id == winners[0]
            assert session.query(db.AssignmentAudit).filter_by(
                game_id=game_id, action="claim"
            ).count() == 1
    finally:
        engine.dispose()


def test_stale_release_does_not_remove_replacement_or_write_audit():
    path, (game_id, first_id, second_id, _third_id) = _new_database()
    setup_engine = _engine(path)
    with h.Session(setup_engine) as setup:
        db.claim_slot(
            setup,
            setup.get(db.Game, game_id),
            db.ROLE_SALE,
            0,
            None,
            setup.get(db.Person, first_id),
        )
        setup.commit()
    setup_engine.dispose()

    stale_engine = _engine(path)
    replacing_engine = _engine(path)
    try:
        with h.Session(stale_engine) as stale:
            stale.connection().exec_driver_sql("BEGIN")
            stale_game = stale.get(db.Game, game_id)
            assert stale_game.assignments_by_role(db.ROLE_SALE)[0].person_id == first_id

            with h.Session(replacing_engine) as replacing:
                replacing_game = replacing.get(db.Game, game_id)
                db.release_slot(
                    replacing, replacing_game, db.ROLE_SALE, 0, first_id
                )
                db.claim_slot(
                    replacing,
                    replacing_game,
                    db.ROLE_SALE,
                    0,
                    None,
                    replacing.get(db.Person, second_id),
                )
                replacing.commit()

            try:
                db.release_slot(
                    stale, stale_game, db.ROLE_SALE, 0, first_id
                )
            except db.SlotConflictError as exc:
                assert exc.current_person_id == second_id
                stale.rollback()
            else:
                raise AssertionError("a stale release must conflict")
    finally:
        stale_engine.dispose()
        replacing_engine.dispose()

    check_engine = _engine(path)
    try:
        with h.Session(check_engine) as session:
            assignment = session.query(db.Assignment).filter_by(
                game_id=game_id, role=db.ROLE_SALE, slot=0
            ).one()
            assert assignment.person_id == second_id
            audits = session.query(db.AssignmentAudit).filter_by(game_id=game_id).all()
            assert [audit.action for audit in audits] == ["claim", "release", "claim"]
    finally:
        check_engine.dispose()


def test_one_person_per_game_race_commits_only_one_assignment_and_audit():
    path, (game_id, first_id, _second_id, _third_id) = _new_database()
    barrier = threading.Barrier(2)
    results = {}
    threads = [
        threading.Thread(
            target=_claim,
            args=(path, game_id, first_id, role, 0, barrier, results, role),
        )
        for role in (db.ROLE_TIMEKEEPER, db.ROLE_SECRETARY)
    ]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(8)
        assert not thread.is_alive(), "one-person race did not finish"

    statuses = sorted(value[0] for value in results.values())
    assert statuses == ["person_conflict", "winner"], results
    engine = _engine(path)
    try:
        with h.Session(engine) as session:
            assert session.query(db.Assignment).filter_by(
                game_id=game_id, person_id=first_id
            ).count() == 1
            assert session.query(db.AssignmentAudit).filter_by(
                game_id=game_id, action="claim"
            ).count() == 1
    finally:
        engine.dispose()


def test_assignment_endpoints_keep_409_and_map_exhaustion_to_503_with_rollback():
    path, (game_id, first_id, second_id, _third_id) = _new_database()
    admin_engine = _engine(path)
    with h.Session(admin_engine) as session:
        session.get(db.Person, first_id).is_admin = True
        session.commit()
    admin_engine.dispose()

    previous_db = os.environ["NULIGAHELPER_DB"]
    os.environ["NULIGAHELPER_DB"] = path
    try:
        app = webapp.create_app()
    finally:
        os.environ["NULIGAHELPER_DB"] = previous_db
    client = app.test_client()
    token = h.sign_in(client, first_id)
    headers = h.csrf_headers(token)
    claim_body = {
        "game_id": game_id,
        "role": db.ROLE_SALE,
        "slot": 0,
        "expected_person_id": None,
        "person_id": first_id,
    }

    original_claim = db.claim_slot

    def unavailable_claim(session, *_args, **_kwargs):
        session.add(db.Person(name="Rolled Back Claim"))
        raise db.AssignmentTemporarilyUnavailableError()

    db.claim_slot = unavailable_claim
    try:
        unavailable = client.post(
            "/api/assignment/claim", json=claim_body, headers=headers
        )
    finally:
        db.claim_slot = original_claim
    assert unavailable.status_code == 503
    assert unavailable.get_json()["code"] == "temporarily_unavailable"

    engine = _engine(path)
    with h.Session(engine) as session:
        assert session.query(db.Person).filter_by(name="Rolled Back Claim").count() == 0
        original_claim(
            session,
            session.get(db.Game, game_id),
            db.ROLE_SALE,
            0,
            None,
            session.get(db.Person, second_id),
        )
        session.commit()
    engine.dispose()

    conflict = client.post("/api/assignment/claim", json=claim_body, headers=headers)
    assert conflict.status_code == 409
    assert conflict.get_json()["code"] == "conflict"
    assert conflict.get_json()["current_person_id"] == second_id

    original_release = db.release_slot

    def unavailable_release(session, *_args, **_kwargs):
        session.add(db.Person(name="Rolled Back Release"))
        raise db.AssignmentTemporarilyUnavailableError()

    db.release_slot = unavailable_release
    try:
        unavailable = client.post(
            "/api/assignment/release",
            json={
                "game_id": game_id,
                "role": db.ROLE_SALE,
                "slot": 0,
                "expected_person_id": second_id,
            },
            headers=headers,
        )
    finally:
        db.release_slot = original_release
    assert unavailable.status_code == 503
    assert unavailable.get_json()["code"] == "temporarily_unavailable"

    engine = _engine(path)
    try:
        with h.Session(engine) as session:
            assert session.query(db.Person).filter_by(
                name="Rolled Back Release"
            ).count() == 0
            assignment = session.query(db.Assignment).filter_by(
                game_id=game_id, role=db.ROLE_SALE, slot=0
            ).one()
            assert assignment.person_id == second_id
    finally:
        engine.dispose()


if __name__ == "__main__":
    h.run_all(dict(globals()))
