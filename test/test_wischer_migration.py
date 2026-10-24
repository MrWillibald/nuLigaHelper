"""Synthetic migration, stale-client and recovery coverage for the one-time removal."""

import os
from pathlib import Path
import shutil
import sqlite3
import tempfile
from datetime import date, datetime
from concurrent.futures import ThreadPoolExecutor
from threading import Barrier
from unittest.mock import patch
from zoneinfo import ZoneInfo

from alembic import command
from alembic.script import ScriptDirectory
from lxml import html

import helpers as h
import db
import notifier
import schema_migrations as sm
from schema_fixtures import create_versioned_database

PREVIOUS = "0007_binding_duty_assignments"


def _source(path, revision=PREVIOUS):
    engine = create_versioned_database(path, revision)
    with sqlite3.connect(path) as con:
        con.executemany(
            "INSERT INTO persons (id,name,is_admin,account_status,birth_date) VALUES (?,?,?,'active',?)",
            [(i, f"Fixture {i}", int(i == 1), "1990-01-01" if i < 4 else None) for i in range(1, 12)],
        )
        con.executemany("INSERT INTO person_teams VALUES (?,1)", [(i,) for i in range(1, 12)])
        con.execute("UPDATE teams SET mv_person_id=2 WHERE id=1")
        con.execute("UPDATE persons SET email='fixture-' || id || '@example.test' WHERE id>=3")
        con.executemany(
            "INSERT INTO games (id,season_year,game_nr,date,time,ak,home,guest,team_id) "
            "VALUES (?,2026,?,?,'12:00',?,'Fixture Home','Fixture Away',1)",
            [(10, "9910", "01.11.2026", "M"), (11, "9911", "01.08.2026", "F"),
             (12, "9912", "02.11.2026", "mD"), (13, "9913", "03.11.2026", "GE")],
        )
        con.execute("UPDATE games SET season_year=2025,date='01.08.2025' WHERE id=11")
        con.executemany(
            "INSERT INTO assignments (id,game_id,person_id,role,slot) VALUES (?,?,?,?,?)",
            [(100, 10, 4, "Reinigung", 0), (101, 10, 5, "Reinigung", 1),
             (102, 11, 6, "Reinigung", 1), (103, 12, 7, "Reinigung", 0),
             (104, 12, 8, "Reinigung", 1), (105, 13, 9, "Reinigung", 1),
             (106, 10, 3, "Kasse", 0)],
        )
        if revision == PREVIOUS:
            con.execute("UPDATE assignments SET replacement_token='saved-first-token' WHERE id IN (100,103)")
            con.execute("UPDATE assignments SET replacement_token='deleted-token' WHERE id IN (101,104)")
            con.execute("UPDATE assignments SET locked_at='2026-07-31 09:00:00' WHERE id=102")
            con.execute("UPDATE assignments SET locked_at='2026-10-30 09:00:00' WHERE id=105")
            con.execute("UPDATE assignments SET locked_at='2026-11-01 09:00:00',replacement_token=NULL WHERE id=103")
        con.execute(
            "INSERT INTO assignment_audit (id,changed_at,actor_tier,action,affected_person_id,"
            "game_id,role,slot,actor_name,affected_person_name,game_snapshot) "
            "VALUES (200,'2026-01-01','system','claim',5,10,'Reinigung',1,'System','Old name','verbatim snapshot')"
        )
    return engine


def _rows(path, table):
    with sqlite3.connect(path) as con:
        return con.execute(f'SELECT * FROM "{table}" ORDER BY id').fetchall()


def _app(path):
    import webapp
    with patch.dict(os.environ, {"NULIGAHELPER_DB": str(path)}):
        return webapp.create_app()


def test_successor_preserves_first_state_all_other_data_and_exact_old_plus_new_audits():
    assert sm.head_revisions() == (sm.HEAD_REVISION,)
    assert ScriptDirectory.from_config(sm.alembic_config()).get_revision(sm.WISCHER_REVISION).down_revision == PREVIOUS
    for source in ("0006_game_duty_staffing", PREVIOUS):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "source.db"
            engine = _source(path, source)
            original = _rows(path, "assignments")
            old_audits = _rows(path, "assignment_audit")
            with patch.object(notifier.Notifier, "send_Mail", side_effect=AssertionError("migration sent mail")), \
                 patch.object(notifier.Notifier, "send_SMS", side_effect=AssertionError("migration sent SMS")):
                result = sm.migrate_to_head(path, engine)
            assert result.backup_path and _rows(result.backup_path, "assignments") == original
            after = _rows(path, "assignments")
            for before, row in zip([a for a in original if a[0] in (100, 103, 106)], after):
                expected = list(before)
                if before[3] == "Reinigung":
                    expected[3] = "Wischer"
                if source != PREVIOUS:
                    expected += [None, None]
                assert tuple(expected) == row, "first identity, lock, request and unrelated duties must survive"
            assert {row[0] for row in after} == {100, 103, 106}, "do not promote second-only occupants"
            audits = _rows(path, "assignment_audit")
            assert audits[:len(old_audits)] == old_audits
            with h.Session(engine) as session:
                removals = session.query(db.AssignmentAudit).filter_by(action="migration_remove").all()
                assert {(a.game_id, a.affected_person_id, a.role, a.slot) for a in removals} == {
                    (10, 5, "Reinigung", 1), (11, 6, "Reinigung", 1),
                    (12, 8, "Reinigung", 1), (13, 9, "Reinigung", 1),
                }
                assert all(a.actor_tier == "system" and a.actor_person_id is None
                           and a.actor_name == "System (Migration 0008)" for a in removals)
                assert all("1990-01-01" not in a.game_snapshot for a in removals)
                assert session.get(db.Game, 11).assignments == [], "past personal counts lose the deleted duty too"
                assert session.query(db.Assignment).filter(db.Assignment.person_id.in_([5,6,8,9])).count() == 0
            # The guarded command also checked all original table fingerprints.
            assert sm.migrate_to_head(path, engine).backup_path is None
            assert _rows(path, "assignment_audit") == audits, "rerun must append nothing"
            db.verify_db(engine)
            engine.dispose()


def test_unexpected_wischer_and_invalid_cleaning_slots_refuse_before_domain_writes():
    for revision, role, slot in [(PREVIOUS, "Wischer", 0), (PREVIOUS, "Wischer", 1),
                                 (PREVIOUS, "Reinigung", 2), (PREVIOUS, "Reinigung", -1),
                                 ("0005_cake_delivery_blocks", "Wischer", 0)]:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "invalid.db"
            engine = create_versioned_database(path, revision)
            with sqlite3.connect(path) as con:
                con.execute("INSERT INTO persons (id,name,is_admin,account_status) VALUES (1,'Fixture',0,'active')")
                con.execute("INSERT INTO games (id,season_year,game_nr) VALUES (1,2026,'9900')")
                con.execute("INSERT INTO assignments (id,game_id,person_id,role,slot) VALUES (1,1,1,?,?)", (role,slot))
            before = path.read_bytes()
            try:
                sm.migrate_to_head(path, engine)
            except sm.SchemaMigrationError as exc:
                assert "Unexpected Wischer target" in str(exc) and "Retained backup:" in str(exc)
            else:
                raise AssertionError("unexpected target data must refuse, never merge")
            assert path.read_bytes() == before
            assert len(list(Path(directory).glob("*.pre-schema-*.db"))) == 1
            engine.dispose()


def test_prior_revision_refuses_runtime_empty_upgrade_and_fresh_init_have_no_removal_evidence():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "empty.db"
        engine = create_versioned_database(path, PREVIOUS)
        before = path.read_bytes()
        for start in (lambda: db.verify_db(engine), lambda: _app(path)):
            try:
                start()
            except db.SQLiteInitializationError as exc:
                assert "migrate-schema" in str(exc)
            else:
                raise AssertionError("startup must refuse the old schema")
        assert path.read_bytes() == before
        sm.migrate_to_head(path, engine)
        assert _rows(path, "assignments") == _rows(path, "assignment_audit") == []
        engine.dispose()
        fresh = h.make_engine()
        with h.Session(fresh) as session:
            assert session.query(db.AssignmentAudit).count() == 0
        fresh.dispose()


def test_postflight_checks_existing_and_appended_evidence_and_restorable_snapshot():
    for sql in (
        "UPDATE assignment_audit SET game_snapshot='corrupt' WHERE id=200",
        "DELETE FROM assignment_audit WHERE action='migration_remove' AND affected_person_id=5",
        "UPDATE assignment_audit SET actor_tier='member' WHERE action='migration_remove'",
        "UPDATE assignments SET person_id=11 WHERE id=100",
    ):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "recover.db"
            engine = _source(path)
            before = _rows(path, "assignments")
            original = sm._run_alembic

            def corrupt(connection, operation, revision):
                original(connection, operation, revision)
                if operation is command.upgrade:
                    connection.exec_driver_sql(sql)

            with patch.object(sm, "_run_alembic", side_effect=corrupt):
                try:
                    sm.migrate_to_head(path, engine)
                except sm.SchemaMigrationError as exc:
                    assert "Retained-data postflight failed" in str(exc) and "Retained backup:" in str(exc)
                else:
                    raise AssertionError("old and appended evidence must both be checked")
            engine.dispose()
            snapshot, = Path(directory).glob("*.pre-schema-*.db")
            assert _rows(snapshot, "assignments") == before
            shutil.copyfile(snapshot, path)
            engine = db.make_engine(str(path))
            sm.migrate_to_head(path, engine)
            assert len(_rows(path, "assignment_audit")) == 5
            engine.dispose()


def test_evidence_and_deletion_rollback_together_when_deletion_fails():
    for failure_point in ("deletion", "revision"):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "atomic.db"
            engine = _source(path)
            before = _rows(path, "assignments"), _rows(path, "assignment_audit")
            with sqlite3.connect(path) as con:
                target = "DELETE ON assignments" if failure_point == "deletion" else "UPDATE ON alembic_version"
                con.execute(f"CREATE TRIGGER fail_removal BEFORE {target} "
                            "BEGIN SELECT RAISE(ABORT, 'synthetic removal failure'); END")
            try:
                sm.migrate_to_head(path, engine)
            except sm.SchemaMigrationError as exc:
                assert "synthetic removal failure" in str(exc)
            else:
                raise AssertionError("injected deletion failure must abort")
            assert (_rows(path, "assignments"), _rows(path, "assignment_audit")) == before
            assert sm.inspect_schema(path).revision == PREVIOUS
            engine.dispose()

def test_remaining_wischer_claims_serialize_and_deleted_helpers_receive_no_reminders():
    from test_notifier import TextRecorder

    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "reminders.db"
        engine = _source(path)
        sm.migrate_to_head(path, engine)
        with h.Session(engine) as session:
            sender = notifier.Notifier(h.load_club_config(), session, h.SEASON)
            for notify in (
                lambda: sender.notify_pre("01.11.2026"),
                lambda: sender.notify_game_day("01.11.2026"),
                lambda: sender.notify_shifts([db.ShiftEvent(
                    game_id=10, game_nr="9910", old_date="01.11.2026", old_time="12:00",
                    new_date="08.11.2026", new_time="13:00",
                )]),
            ):
                recorder = TextRecorder()
                sender.send_Mail, sender.send_SMS = recorder.mail, recorder.sms
                assert notify() == 2, "only saved Kasse and Wischer have contacts/duties"
                assert len(recorder.mails) == 2 and not recorder.smss
                recipients = " ".join(recipient for recipient, _, _ in recorder.mails)
                assert "fixture-3@example.test" in recipients and "fixture-4@example.test" in recipients
                assert "fixture-5@example.test" not in recipients
                assert "Wischer" in recorder.all_text() and "Reinigung" not in recorder.all_text()
            # The only past assignment was explicitly deleted, so neither a
            # reminder nor a personal duty remains for that helper.
            assert sender.notify_pre("01.08.2025") == 0
        # Empty slot zero remains empty after deleting a lone second occupant.
        # Two sessions may compete for it; only one can create the new duty.
        barrier = Barrier(2)

        def claim(person_id):
            with h.Session(engine) as session:
                game = session.get(db.Game, 11)
                person = session.get(db.Person, person_id)
                barrier.wait(timeout=5)
                try:
                    assignment = db.claim_slot(session, game, "Wischer", 0, None, person)
                    session.commit()
                    assert assignment.locked_at is not None, "late claims immediately lock"
                    return "saved"
                except db.SlotConflictError:
                    session.rollback()
                    return "conflict"

        with ThreadPoolExecutor(max_workers=2) as pool:
            outcomes = list(pool.map(claim, (5, 6)))
        assert sorted(outcomes) == ["conflict", "saved"]
        with h.Session(engine) as session:
            assert session.query(db.Assignment).filter_by(game_id=11).count() == 1
            assert session.query(db.AssignmentAudit).filter_by(game_id=11, action="claim").count() == 1
        engine.dispose()


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_migrated_stale_clients_cannot_recreate_duties_and_admin_history_is_readable(_today):
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "clients.db"
        engine = _source(path)
        sm.migrate_to_head(path, engine)
        app = _app(path)
        before = _rows(path, "assignments"), _rows(path, "assignment_audit")
        for actor in (1, 2, 5):  # admin, MV, former helper/member
            client = app.test_client()
            h.sign_in(client, actor)
            for role, slot in (("Reinigung", 0), ("Reinigung", 1), ("Wischer", 1)):
                for endpoint in ("claim", "request-replacement", "takeover", "cancel-replacement"):
                    response = client.post(f"/api/assignment/{endpoint}", json={
                        "game_id": 10, "role": role, "slot": slot,
                        "person_id": actor, "expected_person_id": None if endpoint == "claim" else 5,
                        "expected_request_token": "deleted-token",
                    }, headers=h.csrf_headers())
                    assert response.status_code == 400, (role, slot, endpoint, response.get_json())
            candidates = client.get("/api/games/10/candidates").get_json()
            assert "Wischer:0" in candidates["slots"] and len(candidates["slots"]) == 7
            assert all(not key.startswith("Reinigung") and key != "Wischer:1" for key in candidates["slots"])
        with h.Session(engine) as session:
            for role, slot in (("Reinigung", 0), ("Reinigung", 1), ("Wischer", 1)):
                for action in ("claim", "request", "takeover"):
                    game = session.get(db.Game, 10)
                    incoming = session.get(db.Person, 5)
                    try:
                        if action == "claim":
                            db.claim_slot(session, game, role, slot, None, incoming)
                        elif action == "request":
                            db.request_replacement(session, game, role, slot, 5)
                        else:
                            db.takeover_slot(session, game, role, slot, 5, "deleted-token", incoming)
                    except ValueError:
                        session.rollback()
                    else:
                        raise AssertionError("serialized helpers must reject deleted or renamed duties")
        assert (_rows(path, "assignments"), _rows(path, "assignment_audit")) == before
        guest = app.test_client().get("/").get_data(as_text=True)
        assert "Wischer" in guest and "Reinigung" not in guest and "Fixture 5" not in guest
        assert "person_id" not in guest and "birth_date" not in guest and "1990-01-01" not in guest and "@example.test" not in guest
        admin = app.test_client()
        h.sign_in(admin, 1)
        history = html.fromstring(admin.get("/audit").get_data(as_text=True))
        assert len(history.xpath('//*[contains(@class,"audit-action-migration_remove")]')) == 4
        assert "Durch Migration entfernt" in history.text_content() and "verbatim snapshot" in history.text_content()
        assert "Reinigung" in history.text_content(), "historical role text must remain readable"
        # Remaining unknown-date Wischer is eligible, transfers atomically before
        # cutoff, then remains protected at and after the shared fixed deadline.
        with h.Session(engine) as session:
            game = session.get(db.Game, 10)
            db.takeover_slot(session, game, "Wischer", 0, 4, "saved-first-token", session.get(db.Person, 5))
            session.commit()
            assert game.assignment_by_role("Wischer").person_id == 5
            for hour in (9, 10):
                with patch("db.duty_now", return_value=datetime(2026,10,31,hour,tzinfo=ZoneInfo("Europe/Berlin"))):
                    try:
                        db.request_replacement(session, game, "Wischer", 0, 5)
                    except db.DutyLockedError:
                        session.rollback()
                    else:
                        raise AssertionError("migration must not add a runtime lock bypass")
        engine.dispose()


if __name__ == "__main__":
    h.run_all(dict(globals()))
