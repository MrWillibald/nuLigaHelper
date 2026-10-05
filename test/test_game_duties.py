"""Synthetic category-dependent duties, retained assignments and position CAS."""

import contextlib
import io
import os
import tempfile
from datetime import date
from unittest.mock import patch

from lxml import html

import helpers as h
import db
import manage_db
import webapp


def _site():
    path = os.path.join(h._TEST_DIR, f"duties-{next(tempfile._get_candidate_names())}.db")
    engine = db.make_engine(path)
    db.initialize_db(engine)
    with h.Session(engine) as session:
        team = db.Team(name="Duty Team")
        admin = db.Person(name="Duty Admin", birth_date=h.ADULT_BIRTH_DATE, is_admin=True)
        mv = db.Person(name="Duty MV", birth_date=h.ADULT_BIRTH_DATE, teams=[team])
        member = db.Person(name="Duty Member", birth_date=h.ADULT_BIRTH_DATE, teams=[team])
        outsider = db.Person(name="Duty Outsider", birth_date=h.ADULT_BIRTH_DATE)
        helpers = [db.Person(name=f"Staff {n}", birth_date=h.ADULT_BIRTH_DATE, teams=[team]) for n in range(8)]
        games = [db.Game(season_year=h.SEASON, game_nr=str(9500+n), date="20.10.2026",
                         time="12:00", ak=ak, home="Duty Home", guest=f"Opponent {n}", team=team)
                 for n, ak in enumerate(("BL M", "BL wD", "GE", "SPF Mini"))]
        session.add_all([admin, mv, member, outsider, *helpers, *games])
        session.flush()
        team.mv_person_id = mv.id
        ids = {"admin": admin.id, "mv": mv.id, "member": member.id,
               "outsider": outsider.id, "helpers": [p.id for p in helpers],
               "games": [g.id for g in games]}
        session.commit()
    previous = os.environ["NULIGAHELPER_DB"]
    os.environ["NULIGAHELPER_DB"] = path
    try:
        app = webapp.create_app()
    finally:
        os.environ["NULIGAHELPER_DB"] = previous
    return app, engine, path, ids


def _client(app, actor):
    client = app.test_client()
    h.sign_in(client, actor)
    return client


def _claim(client, game, person, role, slot=0):
    return client.post("/api/assignment/claim", json={
        "game_id": game, "person_id": person, "role": role, "slot": slot,
        "expected_person_id": None,
    }, headers=h.csrf_headers())


def _release(client, game, person, role, slot=0):
    return client.post("/api/assignment/release", json={
        "game_id": game, "role": role, "slot": slot, "expected_person_id": person,
    }, headers=h.csrf_headers())


def _cli(path, *args):
    parsed = manage_db.build_parser().parse_args(["--db", path, *map(str, args)])
    output = io.StringIO()
    with contextlib.redirect_stdout(output):
        parsed.func(parsed)
    return output.getvalue()


def test_explicit_classifier_controls_offered_and_required_duties_for_all_reviewed_classes():
    assert db.ROLE_SLOT_COUNT[db.ROLE_SECURITY] == 1
    assert db.ROLE_SLOT_COUNT[db.ROLE_CLEANING] == 2
    assert db.ROLE_CLEANING != db.ROLE_CASH and db.ROLE_SUPPORT == db.ROLE_CASH
    groups = {
        "adult": ("M", "F", "BL M", "BOL F"),
        "youth": tuple(f"BL {sex}{age}" for sex in "mw" for age in "ABCDE") + ("SPF", "SPF Mini", "spf mini"),
        "unknown": (None, "", "GE", "BL GE", "M1", "mF", "U18", "SPFoo", "mA F", "M SPF", "mD GE"),
    }
    for category, labels in groups.items():
        for label in labels:
            game = db.Game(ak=label, date="20.10.2026")
            assert db.classify_game_category(label) == category, f"misclassified {label!r}"
            offered = db.offered_positions(game)
            assert offered == db.required_positions(game)
            assert len(offered) == (8 if category == "adult" else 5)
            assert sum(role == db.ROLE_SECURITY for role, _ in offered) == 1
            if category != "adult":
                assert not any(role in {db.ROLE_CASH, db.ROLE_CLEANING} for role, _ in offered)
            assert db.staffing_status(game)["classification_unresolved"] == (category == "unknown")


def test_adult_repeated_cleaning_positions_are_independent_audited_and_stale_safe():
    app, engine, path, ids = _site()
    with h.Session(engine) as session:
        game = session.get(db.Game, ids["games"][0])
        first, second = [session.get(db.Person, n) for n in ids["helpers"][:2]]
        a = db.claim_slot(session, game, db.ROLE_CLEANING, 1, None, second)
        assert a.slot == 1 and game.assignment_by_role(db.ROLE_CLEANING, 1).person_id == second.id
        assert game.assignment_by_role(db.ROLE_CLEANING) is None, "slot zero must not borrow slot one"
        db.claim_slot(session, game, db.ROLE_CLEANING, 0, None, first)
        session.commit()
        assert [p.id for p in game.receivers_for_roles([db.ROLE_CLEANING, db.ROLE_CLEANING])] == [first.id, second.id]
        try:
            db.claim_slot(session, game, db.ROLE_CLEANING, 1, None, first)
        except db.SlotConflictError as exc:
            assert exc.current_person_id == second.id
            session.rollback()
        else:
            raise AssertionError("stale claims must preserve the second occupant")
        try:
            db.claim_slot(session, game, db.ROLE_CASH, 0, None, first)
        except ValueError as exc:
            assert "bereits" in str(exc)
            session.rollback()
        else:
            raise AssertionError("cleaning and cash must not bypass one-task-per-game")
        try:
            db.claim_slot(session, game, db.ROLE_SECURITY, 1, None, first)
        except ValueError as exc:
            assert "Ungültiger" in str(exc)
        else:
            raise AssertionError("a second Ordnungsdienst must not exist")
        db.release_slot(session, game, db.ROLE_CLEANING, 1, second.id)
        session.commit()
        assert game.assignment_by_role(db.ROLE_CLEANING, 0).person_id == first.id
        assert game.assignment_by_role(db.ROLE_CLEANING, 1) is None
        assert [(a.action, a.slot) for a in session.query(db.AssignmentAudit).order_by(db.AssignmentAudit.id)] == [
            ("claim", 1), ("claim", 0), ("release", 1)]


def test_saved_category_refuses_stale_new_duties_but_keeps_retained_assignments_releasable():
    app, engine, path, ids = _site()
    with h.Session(engine) as setup:
        game = setup.get(db.Game, ids["games"][0])
        db.claim_slot(setup, game, db.ROLE_CASH, 0, None, setup.get(db.Person, ids["member"]))
        setup.commit()
    with h.Session(engine) as stale:
        game = stale.get(db.Game, ids["games"][0])
        assert game.ak == "BL M"
        assert len(game.assignments) == 1
        with h.Session(engine) as writer:
            writer.get(db.Game, game.id).ak = "BL wD"
            writer.commit()
        try:
            db.claim_slot(stale, game, db.ROLE_CLEANING, 0, None, stale.get(db.Person, ids["outsider"]))
        except ValueError as exc:
            assert "nicht angeboten" in str(exc)
            stale.rollback()
        else:
            raise AssertionError("saved category must supersede a stale adult view")
        assert stale.query(db.Assignment).count() == 1
        assert stale.query(db.AssignmentAudit).count() == 1
        assert db.staffing_status(game)["required_total"] == 5
        db.release_slot(stale, game, db.ROLE_CASH, 0, ids["member"])
        stale.commit()
        assert stale.query(db.Assignment).count() == 0
        assert stale.query(db.AssignmentAudit).count() == 2


def test_category_markup_candidates_claim_refusals_and_retained_release_match_existing_tiers():
    app, engine, path, ids = _site()
    adult, youth, unknown, spf = ids["games"]
    admin = _client(app, ids["admin"])
    page = html.fromstring(admin.get("/").get_data(as_text=True))
    for game_id, count in ((adult, 8), (youth, 5), (unknown, 5), (spf, 5)):
        card = page.get_element_by_id(f"game-{game_id}")
        selects = card.xpath('.//select[@data-role]')
        assert len(selects) == count, "category markup must omit removed fields"
        assert len(card.xpath('.//select[@data-role="Ordnungsdienst"]')) == 1
        assert "optional" not in card.text_content()
        assert not card.get("open"), "duties must remain initially collapsed"
        payload = admin.get(f"/api/games/{game_id}/candidates").get_json()
        assert len(payload["slots"]) == count
        assert payload["staffing"]["required_total"] == count
    for tier in ("admin", "mv", "member"):
        client = _client(app, ids[tier])
        for game_id in (youth, unknown, spf):
            for role, slot in ((db.ROLE_CASH, 0), (db.ROLE_CLEANING, 0), (db.ROLE_CLEANING, 1)):
                response = _claim(client, game_id, ids[tier], role, slot)
                assert response.status_code == 400, f"{tier} bypassed {role} category availability"
                assert "nicht angeboten" in response.get_json()["error"]
    assert _claim(app.test_client(), adult, ids["member"], db.ROLE_CASH).status_code == 401
    assert _claim(_client(app, ids["member"]), adult, ids["outsider"], db.ROLE_CLEANING, 1).status_code == 403

    with h.Session(engine) as session:
        game = session.get(db.Game, youth)
        session.add(db.Assignment(game=game, role=db.ROLE_CASH, slot=0,
                                  person=session.get(db.Person, ids["member"])))
        session.commit()
    member = _client(app, ids["member"])
    card = html.fromstring(member.get("/").get_data(as_text=True)).get_element_by_id(f"game-{youth}")
    retained = card.xpath('.//select[@data-release-only="true"]')
    assert len(retained) == 1 and "Bestehende Einteilung" in card.text_content()
    assert [o.get("value") for o in retained[0].xpath("option")] == ["", str(ids["member"])]
    payload = member.get(f"/api/games/{youth}/candidates").get_json()
    assert payload["slots"]["Kasse:0"] == {"candidate_ids": [ids["member"]],
        "occupant_id": ids["member"], "release_only": True}
    for tier in ("admin", "mv"):
        payload = _client(app, ids[tier]).get(f"/api/games/{youth}/candidates").get_json()
        assert payload["slots"]["Kasse:0"]["candidate_ids"] == [ids["member"]], "no replacement candidate"
    guest = app.test_client().get("/").get_data(as_text=True)
    assert "Duty Member" in guest and "Bestehende Einteilung" in guest
    assert "data-occupant-id" not in guest and "data-release-only" not in guest
    assert "birth_date" not in guest and "1990-01-01" not in guest
    assert _release(_client(app, ids["outsider"]), youth, ids["member"], db.ROLE_CASH).status_code == 403
    assert _release(member, youth, ids["member"], db.ROLE_CASH).status_code == 200
    assert _release(member, youth, ids["member"], db.ROLE_CASH).status_code == 409
    card = html.fromstring(admin.get("/").get_data(as_text=True)).get_element_by_id(f"game-{youth}")
    assert not card.xpath('.//select[@data-role="Kasse"]'), "released removed fields must disappear"
    with h.Session(engine) as session:
        assert session.query(db.AssignmentAudit).one().action == "release"


def test_reports_use_offered_positions_and_count_retained_duties_without_vacancies():
    app, engine, path, ids = _site()
    adult, youth = ids["games"][:2]
    with h.Session(engine) as session:
        for game_id in (adult, youth):
            game = session.get(db.Game, game_id)
            for (role, slot), person_id in zip(db.BASELINE_REQUIRED_POSITIONS, ids["helpers"]):
                db.claim_slot(session, game, role, slot, None, session.get(db.Person, person_id))
        game = session.get(db.Game, youth)
        session.add(db.Assignment(game=game, person=session.get(db.Person, ids["member"]), role=db.ROLE_CASH, slot=0))
        session.commit()
        assert db.missing_slots(session.get(db.Game, adult)) == {"Kasse": 1, "Reinigung": 2}
        assert db.missing_slots(game) == {}
        assert db.staffing_status(game)["complete"]
        assert db.staffing_status(game)["required_filled"] == 5
    text = _client(app, ids["admin"]).get("/statistik").get_data(as_text=True)
    tree = html.fromstring(text)
    assert "Duty Member" in tree.text_content() and "Kasse" in tree.text_content()
    assert "Opponent 0" in tree.text_content() and "Opponent 1" not in tree.text_content(), "youth removed duties must not cause report gaps"
    page = _client(app, ids["admin"]).get("/").get_data(as_text=True)
    assert "5 von 8 Pflichtdiensten besetzt" in page and "5 von 5 Pflichtdiensten besetzt" in page


def test_cli_supports_both_cleaning_positions_and_refuses_removed_category_duties():
    app, engine, path, ids = _site()
    adult, youth = ids["games"][:2]
    for role in db.ROLE_SLOT_COUNT:
        assert manage_db.build_parser().parse_args(["assign", str(adult), role, str(ids["member"])]).role == role
    assert "Reinigung 1" in _cli(path, "assign", adult, db.ROLE_CLEANING, ids["helpers"][0])
    assert "Reinigung 2" in _cli(path, "assign", adult, db.ROLE_CLEANING, ids["helpers"][1])
    listing = _cli(path, "--season", h.SEASON, "list-games")
    assert "Reinigung 1: Staff 0" in listing and "Reinigung 2: Staff 1" in listing
    for role in (db.ROLE_CASH, db.ROLE_CLEANING):
        try:
            _cli(path, "assign", youth, role, ids["member"])
        except SystemExit as exc:
            assert "nicht angeboten" in str(exc)
        else:
            raise AssertionError("CLI cannot override offered game duties")
    _cli(path, "unassign", adult, db.ROLE_CLEANING, ids["helpers"][1])
    with h.Session(engine) as session:
        game = session.get(db.Game, adult)
        assert game.assignment_by_role(db.ROLE_CLEANING, 0) is not None
        assert game.assignment_by_role(db.ROLE_CLEANING, 1) is None


for _name, _test in list(globals().items()):
    if _name.startswith("test_") and callable(_test):
        globals()[_name] = patch("common.effective_today", new=lambda: date(2026, 9, 1))(_test)

if __name__ == "__main__":
    h.run_all(dict(globals()))
