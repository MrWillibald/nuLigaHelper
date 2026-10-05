"""Synthetic calendar boundaries, atomic age rules and preserved staffing."""

from datetime import date, datetime
import json
import threading

import helpers as h
import age_eligibility as age
import common
import db


def _refused(call, code=None):
    try:
        call()
    except ValueError as exc:
        if code is not None:
            assert isinstance(exc, db.AgeEligibilityError), str(exc)
            assert exc.reason["code"] == code, exc.reason
        return exc
    raise AssertionError("an invalid date or age-restricted claim must be refused")


def _game(session, *, scheduled="05.10.2026", category="BL M"):
    game = db.Game(season_year=2026, game_nr="9001", date=scheduled, ak=category)
    session.add(game)
    session.flush()
    return game


def _person(session, name, born=None):
    person = db.Person(name=name, birth_date=born)
    session.add(person)
    session.flush()
    return person


def test_birth_date_validation_uses_effective_day_and_never_echoes_private_input():
    previous = (common.DEBUG_FLAG, common.CHANGE_DAY, common.DEBUG_TODAY)
    try:
        common.DEBUG_FLAG = False
        common.CHANGE_DAY = True
        common.DEBUG_TODAY = date(2026, 10, 5)
        assert db.validate_birth_date("2008-02-29") == date(2008, 2, 29)
        assert db.validate_birth_date("29.02.2008") == date(2008, 2, 29)
        assert db.validate_birth_date(date(2026, 10, 5)) == common.DEBUG_TODAY
        for value in (None, "", "2009-02-29", "2026-10-06", "20080101", datetime(2008, 1, 1)):
            exc = _refused(lambda value=value: db.validate_birth_date(value))
            if isinstance(value, str) and value:
                assert value not in str(exc), "validation must not echo a raw birth date"
        common.CHANGE_DAY = False
        common.DEBUG_FLAG = True
        common.DEBUG_TODAY = date(2000, 1, 1)
        _refused(lambda: db.validate_birth_date(date(2000, 1, 2)))
    finally:
        common.DEBUG_FLAG, common.CHANGE_DAY, common.DEBUG_TODAY = previous


def test_registration_validation_refuses_partial_creation_and_preserves_explicit_date():
    engine = h.make_engine()
    with h.Session(engine) as session:
        team = db.get_support_team(session)
        for value in (None, "2009-02-29", date.max):
            _refused(lambda: db.register_person(
                session, "Invalid", [team], "invalid@example.test", birth_date=value
            ))
            assert session.query(db.Person).count() == 0
        pending = db.register_person(session, "Pending", [team],
                                     "pending@example.test", birth_date="2008-02-29")
        db.verify_person(session, pending)
        db.approve_person(session, pending)
        session.commit()
        assert pending.birth_date == date(2008, 2, 29)
        assert db.has_team(pending, team)


def test_explicit_category_tokens_handle_leagues_and_reject_unsupported_classes():
    for label in ("M", "BL M", "BOL-F", "(LL) F", "ÜBOL M"):
        assert db.classify_game_category(label) == age.ADULT, label
    for token in ("mA", "mB", "mC", "mD", "mE", "wA", "wB", "wC", "wD", "wE"):
        for label in (token, "BL " + token, "ÜBOL-" + token):
            assert db.classify_game_category(label) == age.YOUTH, label
    for label in ("SPF", "SPF Mini", "spf E-Jugend"):
        assert db.classify_game_category(label) == age.YOUTH, label
    for label in (None, "", "GE", "BL GE", "mF", "MA", "M1", "SPFoo", "M F mA", "unknown"):
        assert db.classify_game_category(label) == age.UNKNOWN, label


def test_calendar_years_advance_inclusively_and_leap_birthdays_advance_on_march_first():
    born = date(2008, 10, 5)
    assert age.completed_age(born, date(2026, 10, 4)) == 17
    assert age.completed_age(born, date(2026, 10, 5)) == 18
    leap = date(2008, 2, 29)
    assert age.completed_age(leap, date(2026, 2, 28)) == 17
    assert age.completed_age(leap, date(2026, 3, 1)) == 18
    assert age.completed_age(leap, date(2024, 2, 29)) == 16


def test_timing_minimums_are_evaluated_on_game_date_without_admin_or_system_override():
    engine = h.make_engine()
    with h.Session(engine) as session:
        game = _game(session)
        person = _person(session, "Boundary", date(2008, 10, 5))
        younger = _person(session, "Birthday tomorrow", date(2008, 10, 6))
        unknown = _person(session, "Legacy unknown")
        for role, minimum in ((db.ROLE_TIMEKEEPER, 18), (db.ROLE_SECRETARY, 16)):
            game.ak = "BL M"
            person.birth_date = date(2026 - minimum, 10, 5)
            younger.birth_date = date(2026 - minimum, 10, 6)
            assert db.claim_eligibility(game, role, person) is None
            assert db.claim_eligibility(game, role, younger)["minimum_age"] == minimum
            assert db.claim_eligibility(game, role, unknown)["code"] == "unknown_birth_date"
            game.ak = "BL mA"
            person.birth_date = date(2012, 10, 5)
            younger.birth_date = date(2012, 10, 6)
            assert db.claim_eligibility(game, role, person) is None
            assert db.claim_eligibility(game, role, younger)["code"] == "underage"
        game.ak = "BL M"
        younger.birth_date = date(2010, 1, 1)
        session.commit()
        for tier in ("member", "mv", "admin", "system"):
            _refused(lambda: db.claim_slot(session, game, db.ROLE_TIMEKEEPER,
                                          0, None, younger, actor_tier=tier), "underage")
        assert session.query(db.Assignment).count() == 0
        assert session.query(db.AssignmentAudit).count() == 0
        game.date = "invalid"
        assert db.claim_eligibility(game, db.ROLE_TIMEKEEPER, person)["code"] == "unresolved_game_date"
        game.date = "05.10.2026"
        game.ak = "GE"
        assert db.claim_eligibility(game, db.ROLE_TIMEKEEPER, person)["code"] == "unresolved_category"
        assert db.claim_eligibility(game, db.ROLE_SECURITY, unknown) is None
        assert db.claim_eligibility(game, db.ROLE_CASH, unknown)["code"] == "duty_not_offered"


def test_future_game_allows_upcoming_birthday_and_past_admin_correction_uses_past_date():
    engine = h.make_engine()
    with h.Session(engine) as session:
        game = _game(session, scheduled="05.10.2027")
        person = _person(session, "Turns adult next year", date(2009, 10, 5))
        db.claim_slot(session, game, db.ROLE_TIMEKEEPER, 0, None, person)
        session.commit()
        db.release_slot(session, game, db.ROLE_TIMEKEEPER, 0, person.id)
        game.date = "04.10.2027"
        session.commit()
        _refused(lambda: db.assign_person(session, game, person, db.ROLE_TIMEKEEPER,
                                          actor_tier="admin"), "underage")
        game.date = "04.10.2025"
        session.commit()
        _refused(lambda: db.assign_person(session, game, person, db.ROLE_SECRETARY,
                                          actor_tier="admin"), "underage")


def test_sale_group_permits_young_or_unknown_first_seller_and_requires_an_assigned_adult():
    engine = h.make_engine()
    with h.Session(engine) as session:
        game = _game(session, category="GE")
        young = _person(session, "Young", date(2013, 1, 1))
        unknown = _person(session, "Unknown")
        adult = _person(session, "Adult", date(2008, 10, 5))
        session.commit()
        db.claim_slot(session, game, db.ROLE_SALE, 0, None, unknown)
        session.commit()
        assert db.staffing_status(game)["deficiencies"][0]["code"] == "missing_adult_seller"
        _refused(lambda: db.claim_slot(session, game, db.ROLE_SALE, 1, None, young),
                 "missing_adult_seller")
        assert len(game.assignments_by_role(db.ROLE_SALE)) == 1
        assert session.query(db.AssignmentAudit).count() == 1
        db.claim_slot(session, game, db.ROLE_SALE, 1, None, adult)
        session.commit()
        assert db.staffing_status(game)["deficiencies"] == []
        db.release_slot(session, game, db.ROLE_SALE, 1, adult.id)
        session.commit()
        assert db.claim_eligibility(game, db.ROLE_SALE, young, 1)["code"] == "missing_adult_seller"
        assert db.claim_eligibility(game, db.ROLE_SALE, adult, 1) is None
        # A replacement evaluates the sibling, never the current slot occupant.
        assert db.claim_eligibility(game, db.ROLE_SALE, young, 0) is None


def test_sale_claim_with_unknown_game_date_is_refused_even_when_group_stays_incomplete():
    engine = h.make_engine()
    with h.Session(engine) as session:
        game = _game(session, scheduled="invalid")
        person = _person(session, "Legacy seller")
        session.commit()
        _refused(lambda: db.claim_slot(session, game, db.ROLE_SALE, 0, None, person),
                 "unresolved_game_date")
        assert session.query(db.Assignment).count() == 0
        assert session.query(db.AssignmentAudit).count() == 0


def test_refused_role_replacement_rolls_back_previous_releases_claims_and_audits():
    engine = h.make_engine()
    with h.Session(engine) as session:
        game = _game(session)
        adult = _person(session, "Original adult", date(1980, 1, 1))
        original_young = _person(session, "Original younger", date(2010, 1, 1))
        replacements = [_person(session, name, date(2012, 1, 1))
                        for name in ("First younger", "Second younger")]
        session.commit()
        db.set_role_assignments(session, game, db.ROLE_SALE, [adult.id, original_young.id])
        ids_before = [(assignment.id, assignment.person_id) for assignment in game.assignments]
        audits_before = [entry.id for entry in session.query(db.AssignmentAudit)]
        _refused(lambda: db.set_role_assignments(session, game, db.ROLE_SALE,
                                                [person.id for person in replacements]),
                 "missing_adult_seller")
        session.commit()
        assert [(assignment.id, assignment.person_id) for assignment in game.assignments] == ids_before
        assert [entry.id for entry in session.query(db.AssignmentAudit)] == audits_before


def test_staffing_rechecks_changed_dates_and_category_without_altering_rows_or_audits():
    engine = h.make_engine()
    with h.Session(engine) as session:
        game = _game(session, category="BL mA")
        people = [_person(session, str(index), date(2008, 10, 5)) for index in range(5)]
        roles = [db.ROLE_TIMEKEEPER, db.ROLE_SECRETARY, db.ROLE_SALE, db.ROLE_SALE, db.ROLE_SECURITY]
        for person, role in zip(people, roles):
            db.assign_person(session, game, person, role)
        session.commit()
        audit_before = [entry.id for entry in session.query(db.AssignmentAudit)]
        assignment_before = [(entry.id, entry.person_id) for entry in game.assignments]
        assert db.staffing_status(game)["complete"]
        people[0].birth_date = date(2012, 10, 5)
        game.ak = "BL M"
        game.date = "04.10.2026"
        people[3].birth_date = None
        session.commit()
        status = db.staffing_status(game)
        assert status["vacancies"] == {db.ROLE_CASH: 1, db.ROLE_CLEANING: 2}, \
            "changing youth to adult must expose only the three newly required duties"
        assert {reason["code"] for reason in status["deficiencies"]} == {"underage", "missing_adult_seller"}
        assert not status["complete"]
        assert [(entry.id, entry.person_id) for entry in game.assignments] == assignment_before
        assert [entry.id for entry in session.query(db.AssignmentAudit)] == audit_before
        payload = json.dumps(status)
        assert "2008-10-05" not in payload and "2012-10-05" not in payload
        assert "birth_date" not in payload and '"age"' not in payload
        db.release_slot(session, game, db.ROLE_TIMEKEEPER, 0, people[0].id)
        session.commit()
        assert db.staffing_status(game)["vacancies"] == {
            db.ROLE_TIMEKEEPER: 1, db.ROLE_CASH: 1, db.ROLE_CLEANING: 2}


def test_loaded_person_game_and_seller_dates_are_refreshed_before_claim():
    engine = h.make_engine()
    with h.Session(engine) as setup:
        game = _game(setup)
        adult = _person(setup, "Initially adult", date(1980, 1, 1))
        young = _person(setup, "Younger", date(2012, 1, 1))
        db.claim_slot(setup, game, db.ROLE_SALE, 0, None, adult)
        setup.commit()
        game_id, adult_id, young_id = game.id, adult.id, young.id
    with h.Session(engine) as stale:
        game = stale.get(db.Game, game_id)
        young = stale.get(db.Person, young_id)
        assert game.assignments[0].person.birth_date == date(1980, 1, 1)
        with h.Session(engine) as correction:
            correction.get(db.Person, adult_id).birth_date = date(2012, 1, 1)
            correction.commit()
        _refused(lambda: db.claim_slot(stale, game, db.ROLE_SALE, 1, None, young),
                 "missing_adult_seller")
        stale.rollback()
        assert len(game.assignments) == 1
    with h.Session(engine) as stale:
        game = stale.get(db.Game, game_id)
        person = stale.get(db.Person, adult_id)
        with h.Session(engine) as correction:
            correction.get(db.Person, adult_id).birth_date = date(1980, 1, 1)
            correction.get(db.Game, game_id).ak = "GE"
            correction.commit()
        _refused(lambda: db.claim_slot(stale, game, db.ROLE_SECRETARY, 0, None,
                                       stale.get(db.Person, young_id)), "unresolved_category")


def test_concurrent_young_sale_claims_to_different_slots_revalidate_after_serialization():
    engine = h.make_engine()
    with h.Session(engine) as setup:
        game = _game(setup)
        people = [_person(setup, name, date(2012, 1, 1)) for name in ("First", "Second")]
        setup.commit()
        game_id, ids = game.id, [person.id for person in people]
    barrier = threading.Barrier(2)
    results = []

    def claim(slot, person_id):
        with h.Session(engine) as session:
            session.connection().exec_driver_sql("BEGIN")
            game = session.get(db.Game, game_id)
            person = session.get(db.Person, person_id)
            assert game.assignments == []
            barrier.wait(timeout=3)
            try:
                db.claim_slot(session, game, db.ROLE_SALE, slot, None, person)
                session.commit()
                results.append("winner")
            except db.AgeEligibilityError as exc:
                session.rollback()
                results.append(exc.reason["code"])
            except BaseException as exc:
                session.rollback()
                results.append(repr(exc))

    threads = [threading.Thread(target=claim, args=(slot, person_id))
               for slot, person_id in enumerate(ids)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join(8)
        assert not thread.is_alive(), "serialized claims must finish within contention deadline"
    assert sorted(results) == ["missing_adult_seller", "winner"], results
    with h.Session(engine) as session:
        assert session.query(db.Assignment).count() == 1
        assert session.query(db.AssignmentAudit).count() == 1


if __name__ == "__main__":
    h.run_all(dict(globals()))
