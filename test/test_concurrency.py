"""Concurrent slot claims retain one winner and reject the stale writer."""

import helpers as h
import db


def test_two_stale_claims_keep_the_first_winner():
    engine = h.make_engine()
    with h.Session(engine) as setup:
        games = h.sync_sample_games(setup)
        game = setup.query(db.Game).filter_by(game_nr=games[0]["game_nr"]).one()
        first = db.Person(name="First")
        second = db.Person(name="Second")
        setup.add_all([first, second])
        setup.commit()
        game_id, first_id, second_id = game.id, first.id, second.id

    first_session = h.Session(engine)
    second_session = h.Session(engine)
    try:
        first_game = first_session.get(db.Game, game_id)
        second_game = second_session.get(db.Game, game_id)
        list(first_game.assignments)
        list(second_game.assignments)
        db.claim_slot(
            first_session, first_game, db.ROLE_SALE, 0, None,
            first_session.get(db.Person, first_id),
        )
        first_session.commit()
        try:
            db.claim_slot(
                second_session, second_game, db.ROLE_SALE, 0, None,
                second_session.get(db.Person, second_id),
            )
        except db.SlotConflictError as exc:
            assert exc.current_person_id == first_id
        else:
            raise AssertionError("the stale second claim must conflict")
    finally:
        first_session.close()
        second_session.close()

    with h.Session(engine) as session:
        assignments = session.query(db.Assignment).filter_by(
            game_id=game_id, role=db.ROLE_SALE, slot=0
        ).all()
        assert len(assignments) == 1 and assignments[0].person_id == first_id


def test_stale_release_does_not_remove_replacement_or_write_audit():
    engine = h.make_engine()
    with h.Session(engine) as setup:
        games = h.sync_sample_games(setup)
        game = setup.query(db.Game).filter_by(game_nr=games[0]["game_nr"]).one()
        first = db.Person(name="First")
        second = db.Person(name="Second")
        setup.add_all([first, second])
        setup.flush()
        db.claim_slot(setup, game, db.ROLE_SALE, 0, None, first)
        setup.commit()
        game_id, first_id, second_id = game.id, first.id, second.id

    stale = h.Session(engine)
    replacing = h.Session(engine)
    try:
        stale_game = stale.get(db.Game, game_id)
        list(stale_game.assignments)
        current_game = replacing.get(db.Game, game_id)
        db.release_slot(
            replacing, current_game, db.ROLE_SALE, 0, first_id
        )
        db.claim_slot(
            replacing, current_game, db.ROLE_SALE, 0, None,
            replacing.get(db.Person, second_id),
        )
        replacing.commit()
        audit_count = stale.query(db.AssignmentAudit).count()
        try:
            db.release_slot(stale, stale_game, db.ROLE_SALE, 0, first_id)
        except db.SlotConflictError as exc:
            assert exc.current_person_id == second_id
            stale.rollback()
        else:
            raise AssertionError("a stale release must conflict")
        assert stale.query(db.AssignmentAudit).count() == audit_count
    finally:
        stale.close()
        replacing.close()
    with h.Session(engine) as session:
        assignment = session.query(db.Assignment).filter_by(
            game_id=game_id, role=db.ROLE_SALE, slot=0
        ).one()
        assert assignment.person_id == second_id


if __name__ == "__main__":
    h.run_all(dict(globals()))
