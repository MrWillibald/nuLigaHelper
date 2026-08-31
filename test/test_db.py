# ---------------------------------------------------------------
#                          nuLigaHelper – tests
# ---------------------------------------------------------------
# Database layer: bootstrap, game sync events, ordering, helpers.
#
# Run standalone:  python test/test_db.py
# Or via pytest:   pytest test/test_db.py
# ---------------------------------------------------------------

import helpers as h
import db
from datetime import datetime


def _events_for(session, scraped):
    return db.sync_games(session, scraped, h.SEASON)


def _assert_empty(events):
    assert not (events.shifts or events.referee_alerts
                or events.new_games or events.removed_games), \
        f"expected no events, got {events}"


def _game_by_nr(session, game_nr):
    return session.query(db.Game).filter_by(
        season_year=h.SEASON, game_nr=game_nr).one()


def test_bootstrap_creates_only_support_team():
    engine = h.make_engine()
    with h.Session(engine) as session:
        teams = session.query(db.Team).all()
        assert len(teams) == 1, "a fresh database must only contain the support team"
        assert teams[0].name == "Supporter" and teams[0].is_support


def test_init_db_is_idempotent():
    engine = h.make_engine()
    db.init_db(engine)
    db.init_db(engine)
    with h.Session(engine) as session:
        assert session.query(db.Team).count() == 1


def test_sync_derives_ak_teams_but_never_links_games():
    engine = h.make_engine()
    with h.Session(engine) as session:
        games = h.sync_sample_games(session)

        ak_names = {g["ak"] for g in games}
        team_names = {t.name for t in session.query(db.Team)}
        assert ak_names <= team_names, "every age class must have a team"
        assert "Supporter" in team_names

        unlinked = session.query(db.Game).filter(db.Game.team_id.is_(None)).count()
        assert unlinked == len(games), "games must not be pre-assigned to their own team"


def test_sync_reports_shift_referee_and_new_game_events():
    engine = h.make_engine()
    with h.Session(engine) as session:
        games = h.sync_sample_games(session)

        # identical data again -> no events at all
        import copy
        _assert_empty(_events_for(session, copy.deepcopy(games)))

        # shift one game, remove the referee of another, add an unknown one
        modified = copy.deepcopy(games)
        modified[0]["date"], modified[0]["time"] = "06.09.2026", "18:15"
        modified[1]["score"] = "27:25 §77"
        modified.append({**modified[0], "game_nr": 9999, "ak": "miA"})
        events = _events_for(session, modified)

        assert [(s.game_nr, s.new_date) for s in events.shifts] == [(1003, "06.09.2026")]
        assert [r.game_nr for r in events.referee_alerts] == [1001]
        assert events.new_games == [9999]

        # third run: nothing changed -> no repeated notifications
        _assert_empty(_events_for(session, modified))


def test_removed_games_are_reported():
    engine = h.make_engine()
    with h.Session(engine) as session:
        games = h.sync_sample_games(session)
        events = _events_for(session, games[:2])
        assert set(events.removed_games) == {g["game_nr"] for g in games[2:]}


def test_chronological_ordering_includes_month_and_year():
    engine = h.make_engine()
    with h.Session(engine) as session:
        h.sync_sample_games(session)

        rows = session.query(db.Game).filter_by(season_year=h.SEASON).all()
        rows.sort(key=db.game_sort_key)
        nrs = [g.game_nr for g in rows]
        assert nrs == [1001, 1002, 1004, 1003, 1005, 2001], nrs
        same_day = [g.time for g in rows if g.date == "03.10.2026"]
        assert same_day == ["10:00", "17:30"], "times must be ordered within a day"


def test_delete_person_removes_their_assignments():
    engine = h.make_engine()
    with h.Session(engine) as session:
        games = h.sync_sample_games(session)
        person = db.get_or_create_person(session, "Alice", email="alice@x.de")
        game = _game_by_nr(session, games[0]["game_nr"])
        db.assign_person(session, game, person, db.ROLE_TIMEKEEPER)
        session.commit()

        db.delete_person(session, person)
        assert session.get(db.Person, person.id) is None
        assert game.assignment_by_role(db.ROLE_TIMEKEEPER) is None


def test_set_role_assignments_replaces_all_slots_of_a_role():
    engine = h.make_engine()
    with h.Session(engine) as session:
        games = h.sync_sample_games(session)
        game = _game_by_nr(session, games[0]["game_nr"])
        alice = db.get_or_create_person(session, "Alice")
        bob = db.get_or_create_person(session, "Bob")

        db.set_role_assignments(session, game, db.ROLE_SALE, [alice.id])
        db.set_role_assignments(session, game, db.ROLE_SALE, [alice.id, bob.id])
        names = [a.person.name for a in game.assignments_by_role(db.ROLE_SALE)]
        assert names == ["Alice", "Bob"], names

        db.set_role_assignments(session, game, db.ROLE_SALE, [bob.id])
        names = [a.person.name for a in game.assignments_by_role(db.ROLE_SALE)]
        assert names == ["Bob"], "removed slots must disappear completely"


def test_assign_person_blocks_a_second_task_for_the_same_game():
    engine = h.make_engine()
    with h.Session(engine) as session:
        games = h.sync_sample_games(session)
        game = _game_by_nr(session, games[0]["game_nr"])
        alice = db.get_or_create_person(session, "Alice")

        db.assign_person(session, game, alice, db.ROLE_TIMEKEEPER)
        session.commit()

        try:
            db.assign_person(session, game, alice, db.ROLE_SECRETARY)
        except ValueError as exc:
            assert db.ROLE_TIMEKEEPER in str(exc), \
                "the error must name the role the person already holds"
        else:
            raise AssertionError("a second task for the same game must be rejected")

        again = db.assign_person(session, game, alice, db.ROLE_TIMEKEEPER)
        assert again.role == db.ROLE_TIMEKEEPER, "assigning the same role stays idempotent"


def test_duplicate_names_are_valid_distinct_identities():
    engine = h.make_engine()
    with h.Session(engine) as session:
        first = db.Person(name="Alex")
        second = db.Person(name="Alex")
        session.add_all([first, second])
        session.commit()
        assert first.id != second.id


def test_registration_state_reaches_active_roster_membership():
    engine = h.make_engine()
    with h.Session(engine) as session:
        team = db.get_or_create_team(session, "BL mD")
        person = db.register_person(session, "Alex", team, email="alex@example.test")
        assert person.account_status == db.ACCOUNT_REGISTERED
        db.verify_person(session, person, datetime(2026, 8, 1, 12, 0))
        assert person.account_status == db.ACCOUNT_VERIFIED and person.verified_at
        db.approve_person(session, person, datetime(2026, 8, 2, 12, 0))
        assert person.account_status == db.ACCOUNT_ACTIVE
        assert person.team_id == team.id and person.approved_at


def test_slot_claim_release_conflicts_and_audit_survives_person_deletion():
    engine = h.make_engine()
    with h.Session(engine) as session:
        games = h.sync_sample_games(session)
        game = _game_by_nr(session, games[0]["game_nr"])
        person = db.get_or_create_person(session, "Alex")
        db.claim_slot(session, game, db.ROLE_SALE, 0, None, person)
        session.commit()
        try:
            db.claim_slot(session, game, db.ROLE_SALE, 0, None, person)
        except db.SlotConflictError as exc:
            assert exc.current_person_id == person.id
        else:
            raise AssertionError("claiming an occupied slot must conflict")

        db.release_slot(session, game, db.ROLE_SALE, 0, person.id)
        session.commit()
        assert game.assignments_by_role(db.ROLE_SALE) == []
        db.claim_slot(session, game, db.ROLE_SALE, 1, None, person)
        session.commit()
        db.delete_person(session, person)
        entries = session.query(db.AssignmentAudit).all()
        assert len(entries) == 4, "each successful mutation must have one audit row"
        assert all(e.affected_person_name == "Alex" for e in entries)
        assert all(e.affected_person_id is None for e in entries)


def test_roster_queries_exclude_unapproved_and_inactive_people():
    engine = h.make_engine()
    with h.Session(engine) as session:
        team = db.get_or_create_team(session, "BL mD")
        active = db.Person(name="Active", team=team)
        pending = db.Person(
            name="Pending", desired_team=team, account_status=db.ACCOUNT_VERIFIED
        )
        inactive = db.Person(
            name="Inactive", team=team, account_status=db.ACCOUNT_INACTIVE
        )
        session.add_all([active, pending, inactive])
        session.commit()
        assert db.get_all_persons(session) == [active]
        assert db.get_team_members(session, team) == [active]
        try:
            db.set_team_mv(session, team, inactive)
        except ValueError as exc:
            assert "nicht aktiv" in str(exc)
        else:
            raise AssertionError("an inactive person must not become MV")


def test_deactivation_keeps_past_and_audits_future_release():
    engine = h.make_engine()
    with h.Session(engine) as session:
        team = db.get_or_create_team(session, "BL mD")
        person = db.Person(name="Alex", team=team)
        actor = db.Person(name="Admin", is_admin=True)
        past = db.Game(
            season_year=h.SEASON, game_nr=9001, date="01.01.2020"
        )
        future = db.Game(
            season_year=h.SEASON, game_nr=9002, date="31.12.2099"
        )
        session.add_all([person, actor, past, future])
        session.flush()
        team.mv_person_id = person.id
        db.assign_person(session, past, person, db.ROLE_TIMEKEEPER)
        db.assign_person(session, future, person, db.ROLE_TIMEKEEPER)
        session.commit()
        before = session.query(db.AssignmentAudit).count()

        db.deactivate_person(session, person, actor)
        assert person.account_status == db.ACCOUNT_INACTIVE
        assert team.mv_person_id is None
        assert past.assignment_by_role(db.ROLE_TIMEKEEPER) is not None
        assert future.assignment_by_role(db.ROLE_TIMEKEEPER) is None
        assert session.query(db.AssignmentAudit).count() == before + 1
        assert person not in db.get_all_persons(session)

        db.reactivate_person(session, person)
        assert person in db.get_all_persons(session)
        assert future.assignment_by_role(db.ROLE_TIMEKEEPER) is None


if __name__ == "__main__":
    h.run_all(dict(globals()))
