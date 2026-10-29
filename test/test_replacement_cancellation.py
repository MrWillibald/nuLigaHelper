"""Synthetic self-service search closure, generation races and fixed-cutoff checks."""
from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager
from datetime import date, datetime
import threading
from unittest.mock import patch
from zoneinfo import ZoneInfo

from lxml import html
from sqlalchemy import event

import helpers as h
from nuligahelper import db
from test_sticky_web import fixture, client_for, write, audit_count


@contextmanager
def scenario():
    directory, engine, app, ids = fixture()
    try:
        with patch("nuligahelper.common.effective_today", return_value=date(2026, 10, 30)), patch(
            "nuligahelper.db.clock.duty_now", return_value=datetime(2026, 10, 30, 8, tzinfo=ZoneInfo("Europe/Berlin"))
        ):
            yield engine, app, ids
    finally:
        engine.dispose()
        directory.cleanup()


def targets(ids):
    yield "/api/assignment/", dict(game_id=ids["game"], role=db.ROLE_SALE, slot=0)
    for phase in db.BLOCK_PHASES:
        yield "/api/block-assignment/", dict(block_id=ids[phase], slot=0)


def saved(session, target):
    if "game_id" in target:
        return session.query(db.Assignment).filter_by(**target).one()
    return session.query(db.BlockAssignment).filter_by(**target).one()


def open_search(owner, manager, prefix, target, person_id):
    assert write(owner, prefix + "claim", **target, expected_person_id=None).json["ok"]
    result = write(manager, prefix + "request-replacement", **target, expected_person_id=person_id)
    assert result.json["ok"]
    return result.json["assignment"]["replacement_token"]


def test_current_helper_can_end_a_manager_opened_search_in_every_container():
    with scenario() as (engine, app, ids):
        owner, admin = client_for(app, ids, "Original"), client_for(app, ids, "Admin")
        for prefix, target in targets(ids):
            token = open_search(owner, admin, prefix, target, ids["Original"])
            with h.Session(engine) as session:
                assignment_id = saved(session, target).id
            before = audit_count(engine)
            closed = write(owner, prefix + "cancel-replacement", **target,
                           expected_person_id=ids["Original"], expected_request_token=token)
            assert closed.status_code == 200 and closed.json["ok"]
            state = closed.json["assignment"]
            assert state["person_id"] == ids["Original"] and not state["seeking_replacement"]
            assert state["replacement_token"] is None and state["can_request_replacement"]
            if "block" in closed.json:
                assert closed.json["block"]["progress"]["filled"] == 1
            else:
                assert closed.json["staffing"]["required_filled"] == 1
            assert audit_count(engine) == before + 1
            with h.Session(engine) as session:
                assignment = saved(session, target)
                assert assignment.id == assignment_id and assignment.person_id == ids["Original"]
                audit = session.query(db.AssignmentAudit).order_by(db.AssignmentAudit.id.desc()).first()
                assert (audit.action, audit.actor_person_id, audit.affected_person_id) == (
                    "request_close", ids["Original"], ids["Original"])


def test_cancellation_requires_current_ownership_active_session_and_csrf():
    with scenario() as (engine, app, ids):
        owner, other = client_for(app, ids, "Original"), client_for(app, ids, "Replacement")
        for prefix, target in targets(ids):
            token = open_search(owner, owner, prefix, target, ids["Original"])
            payload = dict(**target, expected_person_id=ids["Original"], expected_request_token=token)
            before = audit_count(engine)
            assert write(other, prefix + "cancel-replacement", **payload).status_code == 403
            assert write(app.test_client(), prefix + "cancel-replacement", **payload).status_code == 401
            assert owner.get(prefix + "cancel-replacement").status_code == 405
            assert owner.post(prefix + "cancel-replacement", json=payload).status_code == 403
            assert owner.post("/dienst", data={**payload, "action": "cancel-replacement"}).status_code == 403
            for status, expected_status in ((db.ACCOUNT_VERIFIED, 403), (db.ACCOUNT_INACTIVE, 401)):
                with h.Session(engine) as session:
                    session.get(db.Person, ids["Original"]).account_status = status
                    session.commit()
                assert write(client_for(app, ids, "Original"), prefix + "cancel-replacement", **payload).status_code == expected_status
            with h.Session(engine) as session:
                session.get(db.Person, ids["Original"]).account_status = db.ACCOUNT_ACTIVE
                session.commit()
                assert saved(session, target).replacement_token == token
            assert audit_count(engine) == before, "refused cancellation cannot record a successful closure"
            owner = client_for(app, ids, "Original")


def test_stale_closed_reopened_and_former_occupant_requests_preserve_current_generation():
    with scenario() as (engine, app, ids):
        owner, other = client_for(app, ids, "Original"), client_for(app, ids, "Replacement")
        for prefix, target in targets(ids):
            old = open_search(owner, owner, prefix, target, ids["Original"])
            payload = dict(**target, expected_person_id=ids["Original"], expected_request_token=old)
            assert write(owner, prefix + "cancel-replacement", **payload).json["ok"]
            before = audit_count(engine)
            assert write(owner, prefix + "cancel-replacement", **payload).status_code == 409
            assert audit_count(engine) == before
            reopened = write(owner, prefix + "request-replacement", **target, expected_person_id=ids["Original"]).json["assignment"]["replacement_token"]
            assert reopened and reopened != old
            before = audit_count(engine)
            assert write(owner, prefix + "cancel-replacement", **payload).status_code == 409
            assert write(other, prefix + "takeover", **payload).status_code == 409
            assert audit_count(engine) == before
            assert write(other, prefix + "takeover", **target, expected_person_id=ids["Original"], expected_request_token=reopened).json["ok"]
            current = write(other, prefix + "request-replacement", **target, expected_person_id=ids["Replacement"]).json["assignment"]["replacement_token"]
            before = audit_count(engine)
            assert write(owner, prefix + "cancel-replacement", **payload).status_code == 409
            assert write(owner, prefix + "cancel-replacement", **target, expected_person_id=ids["Replacement"], expected_request_token=current).status_code == 403
            assert audit_count(engine) == before
            assert write(other, prefix + "cancel-replacement", **target, expected_person_id=ids["Replacement"], expected_request_token=current).json["ok"]


def test_non_javascript_action_uses_the_same_generation_and_preserves_cancel_noop():
    with scenario() as (engine, app, ids):
        owner, other = client_for(app, ids, "Original"), client_for(app, ids, "Replacement")
        for prefix, target in targets(ids):
            token = open_search(owner, owner, prefix, target, ids["Original"])
            before = audit_count(engine)
            page = html.fromstring(owner.get("/dienst", query_string=target).data)
            form = page.xpath('//form[button[@value="cancel-replacement"]]')[0]
            assert "Ersatzsuche beenden" in form.text_content()
            payload = {element.get("name"): element.get("value") for element in form.xpath('.//input[@name]')}
            assert payload["expected_request_token"] == token
            payload["action"] = "cancel-replacement"
            assert audit_count(engine) == before, "opening the form must not mutate a search"
            assert owner.post("/dienst", data={**payload, "decision": "cancel"}).status_code == 302
            assert audit_count(engine) == before
            assert other.post("/dienst", data=payload).status_code == 403
            assert owner.post("/dienst", data=payload).status_code == 302
            with h.Session(engine) as session:
                assert saved(session, target).replacement_token is None
                assert saved(session, target).person_id == ids["Original"]
            assert owner.post("/dienst", data=payload).status_code == 409
            assert audit_count(engine) == before + 1


def test_admin_and_responsible_team_mv_scope_remain_without_foreign_block_rights():
    with scenario() as (engine, app, ids):
        owner, admin, mv = [client_for(app, ids, name) for name in ("Original", "Admin", "MV")]
        with h.Session(engine) as session:
            game = session.get(db.Game, ids["game"])
            game.team_id = session.query(db.Team).filter_by(name="BL M").one().id
            session.commit()
        for prefix, target in targets(ids):
            token = open_search(owner, admin, prefix, target, ids["Original"])
            payload = dict(**target, expected_person_id=ids["Original"], expected_request_token=token)
            assert write(admin, prefix + "cancel-replacement", **payload).json["ok"]
            token = write(owner, prefix + "request-replacement", **target, expected_person_id=ids["Original"]).json["assignment"]["replacement_token"]
            response = write(mv, prefix + "cancel-replacement", **target, expected_person_id=ids["Original"], expected_request_token=token)
            assert response.status_code == (200 if "game_id" in target else 403)


def domain_change(session, target, ids, token, cancel):
    is_game = "game_id" in target
    container = session.get(db.Game if is_game else db.DayBlock, target.get("game_id", target.get("block_id")))
    actor = session.get(db.Person, ids["Original" if cancel else "Replacement"])
    args = [session, container] + ([target["role"]] if is_game else []) + [0, ids["Original"], token]
    if cancel:
        func = db.cancel_replacement if is_game else db.cancel_block_replacement
    else:
        args.append(actor)
        func = db.takeover_slot if is_game else db.takeover_block_slot
    return func(*args, actor=actor, actor_tier="member")


def test_racing_cancellation_and_takeover_have_one_winner_without_a_vacancy():
    with scenario() as (engine, app, ids):
        owner = client_for(app, ids, "Original")
        for prefix, target in targets(ids):
            token = open_search(owner, owner, prefix, target, ids["Original"])
            start = threading.Barrier(2)
            def run(cancel):
                with h.Session(engine) as session:
                    start.wait()
                    try:
                        domain_change(session, target, ids, token, cancel)
                        session.commit()
                        return "closed" if cancel else "transferred"
                    except db.SlotConflictError:
                        session.rollback()
                        return "conflict"
            with ThreadPoolExecutor(max_workers=2) as pool:
                outcomes = list(pool.map(run, (True, False)))
            assert outcomes.count("conflict") == 1
            with h.Session(engine) as session:
                assignment = saved(session, target)
                assert assignment.person_id == ids["Original" if "closed" in outcomes else "Replacement"]
                assert assignment.replacement_token is None
                closures = session.query(db.AssignmentAudit).filter_by(action="request_close")
                closures = closures.filter_by(game_id=target["game_id"]) if "game_id" in target else closures.filter_by(block_id=target["block_id"])
                assert closures.count() == 1


def test_cancellation_waiting_for_writer_rechecks_the_fixed_cutoff():
    for index in range(2):
        with scenario() as (engine, app, ids):
            owner = client_for(app, ids, "Original")
            prefix, target = list(targets(ids))[index]
            token = open_search(owner, owner, prefix, target, ids["Original"])
            clock = [datetime(2026, 10, 31, 8, 59, tzinfo=ZoneInfo("Europe/Berlin"))]
            attempted = threading.Event()
            main_thread = threading.get_ident()
            table = "games" if "game_id" in target else "day_blocks"
            def observe(connection, cursor, statement, parameters, context, many):
                if threading.get_ident() != main_thread and statement.startswith("UPDATE " + table):
                    attempted.set()
            def wait_to_cancel():
                with h.Session(engine) as session:
                    try:
                        domain_change(session, target, ids, token, True)
                        session.commit()
                        return "closed"
                    except (db.DutyLockedError, db.SlotConflictError):
                        session.rollback()
                        return "refused"
            event.listen(engine, "before_cursor_execute", observe)
            try:
                with patch("nuligahelper.db.clock.duty_now", side_effect=lambda: clock[0]):
                    with ThreadPoolExecutor(max_workers=1) as pool:
                        with h.Session(engine) as holder:
                            if "game_id" in target:
                                db._lock_game(holder, target["game_id"])
                            else:
                                db._lock_day_block(holder, target["block_id"])
                            pending = pool.submit(wait_to_cancel)
                            assert attempted.wait(3), "cancellation did not reach writer contention"
                            clock[0] = datetime(2026, 10, 31, 9, tzinfo=ZoneInfo("Europe/Berlin"))
                            holder.commit()
                        assert pending.result(timeout=8) == "refused"
                    with h.Session(engine) as session:
                        db.lock_due_duties(session)
                        session.commit()
                        assert saved(session, target).person_id == ids["Original"]
                        assert saved(session, target).locked_at is not None
                        assert session.query(db.AssignmentAudit).filter_by(action="request_close", actor_person_id=ids["Original"]).count() == 0
            finally:
                event.remove(engine, "before_cursor_execute", observe)


def test_fixed_cutoff_refuses_owner_and_admin_cancellation_for_every_container():
    with scenario() as (engine, app, ids):
        owner, admin = client_for(app, ids, "Original"), client_for(app, ids, "Admin")
        prepared = []
        for prefix, target in targets(ids):
            token = open_search(owner, owner, prefix, target, ids["Original"])
            prepared.append((prefix, target, token))
        with patch("nuligahelper.db.clock.duty_now", return_value=datetime(2026, 10, 31, 9, tzinfo=ZoneInfo("Europe/Berlin"))):
            for prefix, target, token in prepared:
                payload = dict(**target, expected_person_id=ids["Original"], expected_request_token=token)
                assert write(owner, prefix + "cancel-replacement", **payload).status_code in (400, 409)
                assert write(admin, prefix + "cancel-replacement", **payload).status_code in (400, 409)
                page = html.fromstring(owner.get("/dienst", query_string=target).data)
                assert not page.xpath('//button[@value="cancel-replacement"]')
                with h.Session(engine) as session:
                    assert saved(session, target).person_id == ids["Original"]
                    assert saved(session, target).locked_at is not None
                    assert session.query(db.AssignmentAudit).filter_by(action="request_close", actor_person_id=ids["Original"]).count() == 0


if __name__ == "__main__":
    h.run_all(dict(globals()))
