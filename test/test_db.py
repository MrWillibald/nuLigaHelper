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


if __name__ == "__main__":
    h.run_all(dict(globals()))
