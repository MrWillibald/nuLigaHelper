"""Synthetic calendar boundaries, atomic age rules and preserved staffing."""

from datetime import date, datetime
import json
import os
import threading
from unittest.mock import patch

import helpers as h
from nuligahelper.db import age_eligibility as age
from nuligahelper import common
from nuligahelper import db


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
    with patch.dict(os.environ, {"NULIGAHELPER_TODAY": "2026-10-05"}):
        assert db.validate_birth_date("2008-02-29") == date(2008, 2, 29)
        assert db.validate_birth_date("29.02.2008") == date(2008, 2, 29)
        assert db.validate_birth_date(date(2026, 10, 5)) == common.effective_today()
        for value in (None, "", "2009-02-29", "2026-10-06", "20080101", datetime(2008, 1, 1)):
            exc = _refused(lambda value=value: db.validate_birth_date(value))
            if isinstance(value, str) and value:
                assert value not in str(exc), "validation must not echo a raw birth date"
    with patch.dict(os.environ, {"NULIGAHELPER_TODAY": "2000-01-01"}):
        _refused(lambda: db.validate_birth_date(date(2000, 1, 2)))


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
        for role, minimum in ((db.ROLE_TIMEKEEPER, 16), (db.ROLE_SECRETARY, 16)):
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
        younger.birth_date = date(2011, 1, 1)
        session.commit()
        admin = _person(session, "Actual admin", date(1980, 1, 1))
        admin.is_admin = True
        session.commit()
        for tier in ("member", "mv", "admin", "system"):
            acting = None if tier == "system" else admin if tier == "admin" else younger
            with patch("nuligahelper.common.effective_today", return_value=date(2026,9,1)):
                _refused(lambda: db.claim_slot(session, game, db.ROLE_TIMEKEEPER,
                                              0, None, younger, actor=acting, actor_tier=tier), "underage")
        assert session.query(db.Assignment).count() == 0
        assert session.query(db.AssignmentAudit).count() == 0
        game.date = "invalid"
        assert db.claim_eligibility(game, db.ROLE_TIMEKEEPER, person)["code"] == "unresolved_game_date"
        game.date = "05.10.2026"
        game.ak = "GE"
        assert db.claim_eligibility(game, db.ROLE_TIMEKEEPER, person)["code"] == "unresolved_category"
        assert db.claim_eligibility(game, db.ROLE_SECURITY, unknown)["code"] == "unknown_birth_date"
        assert db.claim_eligibility(game, db.ROLE_CASH, unknown)["code"] == "duty_not_offered"


def test_future_game_allows_upcoming_birthday_and_past_admin_correction_uses_past_date():
    engine = h.make_engine()
    with h.Session(engine) as session:
        game = _game(session, scheduled="05.10.2027")
        person = _person(session, "Reaches timing minimum next year", date(2011, 10, 5))
        db.claim_slot(session, game, db.ROLE_TIMEKEEPER, 0, None, person)
        session.commit()
        db.request_replacement(session, game, db.ROLE_TIMEKEEPER, 0, person.id)
        candidate = _person(session, "Replacement with same birthday", person.birth_date)
        admin = _person(session, "Real administrator", date(1980,1,1))
        admin.is_admin = True
        token = game.assignment_by_role(db.ROLE_TIMEKEEPER).replacement_token
        game.date = "04.10.2027"
        session.commit()
        _refused(lambda: db.takeover_slot(session, game, db.ROLE_TIMEKEEPER, 0, person.id, token, candidate,
                                          actor=admin, actor_tier="admin"), "underage")
        game.date = "04.10.2025"
        session.commit()
        _refused(lambda: db.assign_person(session, game, candidate, db.ROLE_SECRETARY,
                                          actor=admin, actor_tier="admin"), "underage")


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
        db.request_replacement(session, game, db.ROLE_SALE, 1, adult.id)
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
        request = db.request_replacement(session, game, db.ROLE_SALE, 0, adult.id)
        session.commit()
        ids_before = [(assignment.id, assignment.person_id) for assignment in game.assignments]
        audits_before = [entry.id for entry in session.query(db.AssignmentAudit)]
        _refused(lambda: db.takeover_slot(session, game, db.ROLE_SALE, 0, adult.id,
                                          request.replacement_token, replacements[0]),
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
        assert status["vacancies"] == {db.ROLE_CASH: 1, db.ROLE_CLEANING: 1}, \
            "changing youth to adult must expose only the two newly required duties"
        assert {reason["code"] for reason in status["deficiencies"]} == {"underage", "missing_adult_seller"}
        assert not status["complete"]
        assert [(entry.id, entry.person_id) for entry in game.assignments] == assignment_before
        assert [entry.id for entry in session.query(db.AssignmentAudit)] == audit_before
        payload = json.dumps(status)
        assert "2008-10-05" not in payload and "2012-10-05" not in payload
        assert "birth_date" not in payload and '"age"' not in payload
        db.request_replacement(session, game, db.ROLE_TIMEKEEPER, 0, people[0].id)
        session.commit()
        assert db.staffing_status(game)["vacancies"] == {
            db.ROLE_CASH: 1, db.ROLE_CLEANING: 1}
        assert game.assignment_by_role(db.ROLE_TIMEKEEPER).person_id == people[0].id
        assert game.assignment_by_role(db.ROLE_TIMEKEEPER).replacement_token


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


def _concurrent_young_sale_claims(category):
    engine = h.make_engine()
    with h.Session(engine) as setup:
        game = _game(setup, category=category)
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


def test_concurrent_young_sale_claims_to_different_slots_revalidate_after_serialization():
    for category in ("BL M", "GE", "", "unsupported"):
        _concurrent_young_sale_claims(category)


def test_individual_policy_boundaries_unknown_dates_and_categories():
    assert db.ROLE_SECURITY == "Ordnungsdienst", "the persisted role key must stay compatible"
    assert db.position_label(db.ROLE_SECURITY, 0) == "Ordner"
    assert db.position_label(db.ROLE_SALE, 1) == "Verkauf 2"
    # Keep the full policy matrix explicit: unknown categories only block timing.
    from types import SimpleNamespace
    for category in ("BL M", "BL F", "BL mA", "BK wE", "SPF Mini", "GE", None, "unsupported"):
        game = SimpleNamespace(ak=category, date="05.10.2026")
        youth = age.classify_game_category(category) == age.YOUTH
        for role in (db.ROLE_TIMEKEEPER, db.ROLE_SECRETARY, db.ROLE_SECURITY):
            minimum = 18 if role == db.ROLE_SECURITY else 14 if youth else 16
            person = SimpleNamespace(birth_date=date(2026 - minimum, 10, 5))
            if role != db.ROLE_SECURITY and age.classify_game_category(category) == age.UNKNOWN:
                assert age.individual_eligibility(game, role, person)["code"] == "unresolved_category"
                continue
            assert age.individual_eligibility(game, role, person) is None, (category, role)
            game.date = "04.10.2026"
            reason = age.individual_eligibility(game, role, person, 0)
            assert reason["code"] == "underage" and reason["minimum_age"] == minimum
            assert reason["role"] == role, "reason identifiers must retain their transport key"
            assert reason["message"].startswith(db.role_label(role) + ":")
            game.date = "05.10.2026"
            person.birth_date = None
            reason = age.individual_eligibility(game, role, person)
            assert reason["code"] == "unknown_birth_date"
            assert reason["message"].startswith(db.role_label(role) + ":")
            for invalid in ("invalid", "", None):
                game.date = invalid
                reason = age.individual_eligibility(game, role, person)
                assert reason["code"] == "unresolved_game_date"
                assert reason["message"].startswith(db.role_label(role) + ":")
            game.date = "05.10.2026"
    for role, born, year in ((db.ROLE_TIMEKEEPER, date(2008, 2, 29), 2025),
                             (db.ROLE_SECURITY, date(2008, 2, 29), 2026)):
        game = SimpleNamespace(ak="BL M", date=f"28.02.{year}")
        person = SimpleNamespace(birth_date=born)
        # Zeitnehmer is already 16 in 2025, Ordnungsdienst is still 17.
        reason = age.individual_eligibility(game, role, person)
        assert (reason is None) == (role == db.ROLE_TIMEKEEPER)
        game.date = f"01.03.{year}"
        assert age.individual_eligibility(game, role, person) is None
    # The revised adult timing threshold is met on the leap birthday itself.
    game = SimpleNamespace(ak="BL M", date="28.02.2024")
    person = SimpleNamespace(birth_date=date(2008, 2, 29))
    assert age.individual_eligibility(game, db.ROLE_TIMEKEEPER, person)["code"] == "underage"
    game.date = "29.02.2024"
    assert age.individual_eligibility(game, db.ROLE_TIMEKEEPER, person) is None


def test_sale_exemption_is_only_for_recognized_youth_including_invalid_dates():
    for category in ("BL mA", "BK wE", "SPF Mini", "BL M", "GE", "", "unsupported"):
        engine = h.make_engine()
        with h.Session(engine) as session:
            game = _game(session, category=category)
            young = _person(session, "Younger", date(2013, 1, 1))
            unknown = _person(session, "Unknown")
            session.commit()
            youth = age.classify_game_category(category) == age.YOUTH
            db.claim_slot(session, game, db.ROLE_SALE, 0, None, young)
            session.commit()
            if youth:
                db.claim_slot(session, game, db.ROLE_SALE, 1, None, unknown)
                session.commit()
                assert not db.staffing_status(game)["deficiencies"]
            else:
                _refused(lambda: db.claim_slot(session, game, db.ROLE_SALE, 1, None, unknown),
                         "missing_adult_seller")
                session.rollback()
            for invalid in (None, "", "invalid"):
                game.date = invalid
                reason = db.claim_eligibility(game, db.ROLE_SALE, unknown, 1)
                assert (reason is None) == youth, (category, invalid, reason)
                coverage = age.sale_coverage(game, [young, unknown])
                assert (coverage is None) == youth
                if not youth:
                    assert reason["code"] == coverage["code"] == "unresolved_game_date"
                # Even first-slot age checking must honor the exemption.
                assert (db.claim_eligibility(game, db.ROLE_SALE, unknown, 0) is None) == youth


def test_youth_takeover_can_remove_only_adult_but_non_youth_preserves_request():
    for category in ("BL mD", "SPF Mini", "BL M", "GE", "", "unsupported"):
        engine = h.make_engine()
        with h.Session(engine) as session:
            game = _game(session, category=category)
            adult = _person(session, "Outgoing adult", date(1980, 1, 1))
            young = _person(session, "Younger sibling", date(2012, 1, 1))
            unknown = _person(session, "Unknown incoming")
            session.commit()
            db.claim_slot(session, game, db.ROLE_SALE, 0, None, adult)
            db.claim_slot(session, game, db.ROLE_SALE, 1, None, young)
            request = db.request_replacement(session, game, db.ROLE_SALE, 0, adult.id)
            session.commit()
            token, assignment_id = request.replacement_token, request.id
            audits = session.query(db.AssignmentAudit).count()
            if age.classify_game_category(category) == age.YOUTH:
                saved = db.takeover_slot(session, game, db.ROLE_SALE, 0, adult.id, token, unknown)
                session.commit()
                assert saved.id == assignment_id and saved.person_id == unknown.id
                assert saved.replacement_token is None
                assert not db.staffing_status(game)["deficiencies"]
            else:
                _refused(lambda: db.takeover_slot(session, game, db.ROLE_SALE, 0, adult.id, token, unknown),
                         "missing_adult_seller")
                session.rollback()
                assert request.person_id == adult.id and request.replacement_token == token
                assert session.query(db.AssignmentAudit).count() == audits


def test_stale_security_claim_and_takeover_refresh_person_and_game_after_retry():
    from sqlalchemy.exc import OperationalError
    for takeover in (False, True):
        for correction in ("birth", "date"):
            engine = h.make_engine()
            with h.Session(engine) as setup:
                game = _game(setup, category="GE")
                outgoing = _person(setup, "Outgoing", date(1980, 1, 1))
                incoming = _person(setup, "Incoming", date(2008, 10, 5))
                setup.commit()
                token = None
                if takeover:
                    db.assign_person(setup, game, outgoing, db.ROLE_SECURITY)
                    request = db.request_replacement(setup, game, db.ROLE_SECURITY, 0, outgoing.id)
                    token = request.replacement_token
                    setup.commit()
                game_id, incoming_id, outgoing_id = game.id, incoming.id, outgoing.id
                audit_count = setup.query(db.AssignmentAudit).count()
            with h.Session(engine) as stale:
                game = stale.get(db.Game, game_id)
                incoming = stale.get(db.Person, incoming_id)
                assert db.claim_eligibility(game, db.ROLE_SECURITY, incoming) is None
                original = stale.execute
                attempted = False

                def execute(statement, *args, **kwargs):
                    nonlocal attempted
                    if not attempted and getattr(statement, "is_update", False):
                        attempted = True
                        with h.Session(engine) as concurrent:
                            if correction == "birth":
                                concurrent.get(db.Person, incoming_id).birth_date = date(2008, 10, 6)
                            else:
                                concurrent.get(db.Game, game_id).date = "04.10.2026"
                            concurrent.commit()
                        raise OperationalError("synthetic contention", {}, Exception("database is locked"))
                    return original(statement, *args, **kwargs)

                with patch.object(stale, "execute", side_effect=execute):
                    if takeover:
                        _refused(lambda: db.takeover_slot(stale, game, db.ROLE_SECURITY, 0,
                                                          outgoing_id, token, incoming), "underage")
                    else:
                        _refused(lambda: db.claim_slot(stale, game, db.ROLE_SECURITY, 0, None, incoming), "underage")
                stale.rollback()
                assert attempted, "test must exercise the serialized retry"
                assert stale.query(db.AssignmentAudit).count() == audit_count
                if takeover:
                    saved = game.assignment_by_role(db.ROLE_SECURITY)
                    assert saved.person_id == outgoing_id and saved.replacement_token == token
                else:
                    assert game.assignments == []


def test_stale_youth_sale_category_and_sibling_state_are_rechecked_after_retry():
    from sqlalchemy.exc import OperationalError
    for takeover in (False, True):
        for category in ("BL M", "GE", ""):
            engine = h.make_engine()
            with h.Session(engine) as setup:
                game = _game(setup, category="BL mD")
                sibling = _person(setup, "Sibling adult", date(1980, 1, 1))
                outgoing = _person(setup, "Outgoing adult", date(1981, 1, 1))
                incoming = _person(setup, "Unknown incoming")
                setup.commit()
                db.claim_slot(setup, game, db.ROLE_SALE, 0, None, sibling)
                token = None
                if takeover:
                    db.claim_slot(setup, game, db.ROLE_SALE, 1, None, outgoing)
                    token = db.request_replacement(setup, game, db.ROLE_SALE, 1, outgoing.id).replacement_token
                setup.commit()
                game_id, sibling_id, outgoing_id, incoming_id = game.id, sibling.id, outgoing.id, incoming.id
                audit_count = setup.query(db.AssignmentAudit).count()
            with h.Session(engine) as stale:
                game = stale.get(db.Game, game_id)
                incoming = stale.get(db.Person, incoming_id)
                assert db.claim_eligibility(game, db.ROLE_SALE, incoming, 1) is None
                original = stale.execute
                attempted = False

                def execute(statement, *args, **kwargs):
                    nonlocal attempted
                    if not attempted and getattr(statement, "is_update", False):
                        attempted = True
                        with h.Session(engine) as concurrent:
                            concurrent.get(db.Game, game_id).ak = category
                            concurrent.get(db.Person, sibling_id).birth_date = date(2012, 1, 1)
                            concurrent.commit()
                        raise OperationalError("synthetic contention", {}, Exception("database is locked"))
                    return original(statement, *args, **kwargs)

                with patch.object(stale, "execute", side_effect=execute):
                    if takeover:
                        _refused(lambda: db.takeover_slot(stale, game, db.ROLE_SALE, 1,
                                                          outgoing_id, token, incoming), "missing_adult_seller")
                    else:
                        _refused(lambda: db.claim_slot(stale, game, db.ROLE_SALE, 1, None, incoming), "missing_adult_seller")
                stale.rollback()
                assert attempted
                assert stale.query(db.AssignmentAudit).count() == audit_count
                assert game.assignment_by_role(db.ROLE_SALE, 0).person_id == sibling_id
                saved = game.assignment_by_role(db.ROLE_SALE, 1)
                if takeover:
                    assert saved.person_id == outgoing_id and saved.replacement_token == token
                else:
                    assert saved is None


if __name__ == "__main__":
    h.run_all(dict(globals()))
