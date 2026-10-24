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
    assert db.ROLE_SLOT_COUNT[db.ROLE_CLEANING] == 1
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
            assert len(offered) == (7 if category == "adult" else 5)
            assert sum(role == db.ROLE_SECURITY for role, _ in offered) == 1
            if category != "adult":
                assert not any(role in {db.ROLE_CASH, db.ROLE_CLEANING} for role, _ in offered)
            assert db.staffing_status(game)["classification_unresolved"] == (category == "unknown")


def test_single_wischer_has_no_age_minimum_and_rejects_stale_positions():
    app, engine, path, ids = _site()
    with h.Session(engine) as session:
        game = session.get(db.Game, ids["games"][0])
        first, second = [session.get(db.Person, n) for n in ids["helpers"][:2]]
        first.birth_date = None
        a = db.claim_slot(session, game, db.ROLE_CLEANING, 0, None, first)
        session.commit()
        assert a.slot == 0 and a.person_id == first.id
        for role, slot in (("Reinigung", 0), ("Reinigung", 1), (db.ROLE_CLEANING, 1)):
            try:
                db.claim_slot(session, game, role, slot, None, second)
            except ValueError:
                session.rollback()
            else:
                raise AssertionError("old-name and second-position claims must fail")
        try:
            db.claim_slot(session, game, db.ROLE_CASH, 0, None, first)
        except ValueError as exc:
            assert "bereits" in str(exc)
            session.rollback()
        else:
            raise AssertionError("Wischer cannot bypass one duty per game")
        request = db.request_replacement(session, game, db.ROLE_CLEANING, 0, first.id)
        token = request.replacement_token
        session.commit()
        db.takeover_slot(session, game, db.ROLE_CLEANING, 0, first.id, token, second)
        session.commit()
        assert game.assignment_by_role(db.ROLE_CLEANING).person_id == second.id
        assert [(a.action, a.slot) for a in session.query(db.AssignmentAudit).order_by(db.AssignmentAudit.id)] == [
            ("claim", 0), ("request_open", 0), ("takeover_out", 0),
            ("request_close", 0), ("takeover_in", 0)]


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
    for game_id, count in ((adult, 7), (youth, 5), (unknown, 5), (spf, 5)):
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
            for role, slot in ((db.ROLE_CASH, 0), (db.ROLE_CLEANING, 0)):
                response = _claim(client, game_id, ids[tier], role, slot)
                assert response.status_code == 400, f"{tier} bypassed {role} category availability"
                assert "nicht angeboten" in response.get_json()["error"]
    assert _claim(app.test_client(), adult, ids["member"], db.ROLE_CASH).status_code == 401
    assert _claim(_client(app, ids["member"]), adult, ids["outsider"], db.ROLE_CLEANING, 0).status_code == 403

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
    retained_slot = payload["slots"]["Kasse:0"]
    assert retained_slot["candidate_ids"] == [ids["member"]]
    assert retained_slot["occupant_id"] == ids["member"] and retained_slot["release_only"] is True
    assert retained_slot["locked"] is False and not retained_slot["seeking_replacement"], "obsolete duties must not gain replacement flow"
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
        assert db.missing_slots(session.get(db.Game, adult)) == {"Kasse": 1, "Wischer": 1}
        assert db.missing_slots(game) == {}
        assert db.staffing_status(game)["complete"]
        assert db.staffing_status(game)["required_filled"] == 5
    text = _client(app, ids["admin"]).get("/statistik").get_data(as_text=True)
    tree = html.fromstring(text)
    assert "Duty Member" in tree.text_content() and "Kasse" in tree.text_content()
    assert "Opponent 0" in tree.text_content() and "Opponent 1" not in tree.text_content(), "youth removed duties must not cause report gaps"
    page = _client(app, ids["admin"]).get("/").get_data(as_text=True)
    assert "5 von 7 Pflichtdiensten besetzt" in page and "5 von 5 Pflichtdiensten besetzt" in page


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_initial_bars_and_statistics_for_adult_youth_spf_and_retained_duties(_today):
    app, engine, _path, ids = _site()
    with h.Session(engine) as session:
        youth = session.get(db.Game, ids["games"][1])
        session.add(db.Assignment(game=youth, person=session.get(db.Person, ids["member"]),
                                  role=db.ROLE_CASH, slot=0))
        session.commit()
    guest = html.fromstring(app.test_client().get("/").get_data(as_text=True))
    admin = _client(app, ids["admin"])
    signed_in = html.fromstring(admin.get("/").get_data(as_text=True))
    stats = html.fromstring(admin.get("/statistik").get_data(as_text=True))
    with h.Session(engine) as session:
        for game_id in ids["games"]:
            game = session.get(db.Game, game_id)
            card = guest.xpath(f'//div[contains(@class,"game-meta") and contains(.,"Nr. {game.game_nr}")]/ancestor::details[1]')[0]
            vacancies = db.staffing_status(game)["vacancies"]
            expected = [f"{role} ({count}×)" if count > 1 else role
                        for role, count in vacancies.items()]
            staffing = db.staffing_status(game)
            filled, total = staffing["required_filled"], staffing["required_total"]
            percent = (filled * 100 + total // 2) // total
            for view in (card, signed_in.get_element_by_id(f"game-{game_id}")):
                assert view.xpath('.//*[@data-progress-count]/text()') == [f"{filled} von {total} Pflichtdiensten besetzt"]
                assert view.xpath('.//*[@data-progress-percent]/text()') == [f"{percent} %"]
                assert view.xpath('.//*[@class="coverage-fill"]/@style') == [f"width:{percent}%"]
                assert not view.xpath('.//*[@data-vacancies or @class="coverage-gauge"] | .//svg')
            # The statistics row uses exactly the same public missing-duty labels.
            identity = f"Spielfest {game.ak}" if db.is_spielfest(game) else game.guest
            row = stats.xpath(f'//tr[contains(.,"{identity}")]')[0]
            assert row.xpath('.//span[contains(@class,"chip-gap")]/text()') == (
                expected + [item["message"] for item in db.staffing_status(game)["deficiencies"]]
            )
            assert card.xpath('.//*[@role="progressbar"]/@aria-valuemax') == [str(len(db.required_positions(game)))]
            assert card.get("open") is None and "Details" in card.xpath('./summary')[0].text_content()
    text = html.tostring(guest, encoding="unicode")
    assert "person_id" not in text and "data-occupant-id" not in text and "birth_date" not in text


def test_cli_supports_one_wischer_and_refuses_removed_category_duties():
    app, engine, path, ids = _site()
    adult, youth = ids["games"][:2]
    for role in db.ROLE_SLOT_COUNT:
        assert manage_db.build_parser().parse_args(["assign", str(adult), role, str(ids["member"])]).role == role
    with contextlib.redirect_stderr(io.StringIO()):
        try:
            manage_db.build_parser().parse_args(["assign", str(adult), "Reinigung", str(ids["member"])])
        except SystemExit as exc:
            assert exc.code == 2
        else:
            raise AssertionError("old role text must be rejected by the CLI parser")
    assert "Wischer" in _cli(path, "assign", adult, db.ROLE_CLEANING, ids["helpers"][0])
    try:
        _cli(path, "assign", adult, db.ROLE_CLEANING, ids["helpers"][1])
    except SystemExit:
        pass
    else:
        raise AssertionError("CLI must not create a second Wischer")
    listing = _cli(path, "--season", h.SEASON, "list-games")
    assert "Wischer: Staff 0" in listing and "Wischer 2" not in listing
    for role in (db.ROLE_CASH, db.ROLE_CLEANING):
        try:
            _cli(path, "assign", youth, role, ids["member"])
        except SystemExit as exc:
            assert "nicht angeboten" in str(exc)
        else:
            raise AssertionError("CLI cannot override offered game duties")
    try:
        _cli(path, "unassign", adult, db.ROLE_CLEANING, ids["helpers"][0])
    except SystemExit as exc:
        assert "verbindlich" in str(exc)
    else:
        raise AssertionError("CLI must not vacate an offered binding duty")
    with h.Session(engine) as session:
        game = session.get(db.Game, adult)
        assert game.assignment_by_role(db.ROLE_CLEANING, 0) is not None
        assert game.assignment_by_role(db.ROLE_CLEANING, 1) is None


for _name, _test in list(globals().items()):
    if _name.startswith("test_") and callable(_test):
        globals()[_name] = patch("common.effective_today", new=lambda: date(2026, 9, 1))(_test)

if __name__ == "__main__":
    h.run_all(dict(globals()))
