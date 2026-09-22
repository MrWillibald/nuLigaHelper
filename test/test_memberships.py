"""Focused domain checks for set-based team membership."""

import helpers as h
import db


def test_person_membership_supports_zero_one_several_and_deduplicates_input():
    engine = h.make_engine()
    with db.Session(engine) as session:
        alpha = db.Team(name="Alpha")
        beta = db.Team(name="beta")
        person = db.Person(name="Member")
        session.add_all([alpha, beta, person])
        session.flush()

        assert db.membership_team_ids(person) == ()
        assert db.membership_label(person) == "ohne Team"

        db.replace_person_teams(session, person, [beta.id])
        assert db.membership_team_ids(person) == (beta.id,)

        db.replace_person_teams(session, person, [beta.id, alpha.id, beta.id])
        assert db.membership_team_ids(person) == (alpha.id, beta.id)
        assert db.membership_team_names(person) == ("Alpha", "beta")
        assert db.person_label(person) == "Member · Alpha, beta"
        session.commit()

        rows = session.execute(db.person_teams.select()).all()
        assert sorted(rows) == sorted([(person.id, alpha.id), (person.id, beta.id)])


def test_membership_replace_rejects_unknown_team_atomically_and_clears_mv():
    engine = h.make_engine()
    with db.Session(engine) as session:
        first = db.Team(name="First")
        second = db.Team(name="Second")
        person = db.Person(name="MV")
        session.add_all([first, second, person])
        session.flush()
        db.replace_person_teams(session, person, [first.id, second.id])
        db.set_team_mv(session, first, person)

        before = db.membership_team_ids(person)
        try:
            db.replace_person_teams(session, person, [second.id, 999999])
        except ValueError:
            pass
        else:
            raise AssertionError("unknown team must reject the full membership set")
        assert db.membership_team_ids(person) == before
        assert first.mv_person_id == person.id

        db.replace_person_teams(session, person, [second.id])
        session.commit()
        assert db.membership_team_ids(person) == (second.id,)
        assert first.mv_person_id is None


def test_lifecycle_retains_memberships_clears_mvs_and_deletes_only_target_rows():
    engine = h.make_engine()
    with db.Session(engine) as session:
        first = db.Team(name="First")
        second = db.Team(name="Second")
        target = db.Person(name="Target")
        other = db.Person(name="Other")
        session.add_all([first, second, target, other])
        session.flush()
        db.replace_person_teams(session, target, [first.id, second.id])
        db.replace_person_teams(session, other, [second.id])
        db.set_team_mv(session, first, target)
        db.set_team_mv(session, second, target)

        db.deactivate_person(session, target)
        assert db.membership_team_ids(target) == (first.id, second.id)
        assert first.mv_person_id is None and second.mv_person_id is None
        db.reactivate_person(session, target)
        assert db.membership_team_ids(target) == (first.id, second.id)
        assert first.mv_person_id is None and second.mv_person_id is None

        target_id = target.id
        other_id = other.id
        second_id = second.id
        db.delete_person(session, target)
        assert session.get(db.Person, target_id) is None
        assert session.execute(
            db.person_teams.select().where(db.person_teams.c.person_id == target_id)
        ).all() == []
        assert session.execute(
            db.person_teams.select().where(db.person_teams.c.person_id == other_id)
        ).all() == [(other_id, second_id)]


if __name__ == "__main__":
    h.run_all(dict(globals()))
