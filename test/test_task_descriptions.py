"""Public task guidance and accessible schedule controls use one role catalog."""

import os
import tempfile
from datetime import date
from unittest.mock import patch

from lxml import html

import helpers as h
import db
import webapp


def _site():
    path = os.path.join(h._TEST_DIR, f"task-help-{next(tempfile._get_candidate_names())}.db")
    engine = db.make_engine(path)
    db.initialize_db(engine)
    previous = os.environ["NULIGAHELPER_DB"]
    os.environ["NULIGAHELPER_DB"] = path
    try:
        app = webapp.create_app()
    finally:
        os.environ["NULIGAHELPER_DB"] = previous
    with h.Session(engine) as session:
        db.sync_games(session, [
            {"day": "Sa", "date": "05.09.2026", "time": "12:00", "hall": 1,
             "game_nr": "7101", "ak": "BL mD", "home": "TuS", "guest": "Youth", "score": ""},
            {"day": "Sa", "date": "05.09.2026", "time": "18:00", "hall": 1,
             "game_nr": "7102", "ak": "BL F", "home": "TuS", "guest": "Adults", "score": ""},
        ], h.SEASON)
        admin = db.Person(id=41001, name="Task Help Admin", is_admin=True,
                          birth_date=h.ADULT_BIRTH_DATE, email="help-admin@fixture.test")
        assigned = db.Person(id=41002, name="Public Assigned Helper",
                             birth_date=h.ADULT_BIRTH_DATE, email="private-help@fixture.test",
                             phone="+4915112345678")
        unassigned = db.Person(id=41003, name="Private Unassigned Roster",
                               birth_date=h.ADULT_BIRTH_DATE, email="roster-help@fixture.test")
        session.add_all([admin, assigned, unassigned])
        session.flush()
        adult = session.query(db.Game).filter_by(game_nr="7102").one()
        db.claim_slot(session, adult, db.ROLE_SALE, 0, None, assigned)
        blocks = {block.phase: block for block in db.get_day_blocks(session, h.SEASON, "05.09.2026")}
        cake = blocks[db.BLOCK_CAKE_DELIVERY]
        db.configure_cake_block(session, cake, "10:00", 4, None, None)
        db.claim_block_slot(session, blocks[db.BLOCK_PREPARATION], 0, None, assigned)
        db.claim_block_slot(session, cake, 0, None, assigned)
        session.commit()
        ids = {"admin": admin.id, "assigned": assigned.id, "cake": cake.id}
    return app, ids


def _client(app, person_id=None):
    client = app.test_client()
    if person_id is not None:
        h.sign_in(client, person_id)
    return client


def _task_controls(tree):
    buttons = tree.xpath('//button[@data-task-help]')
    seen_ids = set()
    descriptions = {}
    for button in buttons:
        heading = button.getparent()
        task_label = button.getprevious()
        assert task_label is not None and task_label.tag in {"label", "span"}, (
            "each task label must be immediately followed by its separate information button"
        )
        assert button.tag == "button" and button.get("type") == "button", (
            "opening task help must not submit an assignment form"
        )
        assert not button.xpath("ancestor::label"), "help must not activate the assignment select's label"
        label = task_label.text_content().strip()
        assert button.get("aria-label") == "Informationen zu " + label, (
            "the information control needs a task-specific accessible name"
        )
        assert button.get("aria-expanded") == "false"
        description_id = button.get("aria-controls")
        assert description_id and description_id == button.get("aria-describedby")
        assert description_id not in seen_ids, "each rendered control needs a unique description association"
        seen_ids.add(description_id)
        description = tree.get_element_by_id(description_id)
        assert description.get("data-task-description") is not None
        assert description.get("popover") == "manual", (
            "task guidance should use an independently controlled overlay in supported browsers"
        )
        assert description.get("tabindex") == "0", "long floating guidance must allow keyboard scrolling"
        assert description.get("hidden") is not None
        assert description.getparent() is heading.getparent(), (
            "task guidance should stay with its own assignment field"
        )
        text = description.text_content()
        assert text.strip(), "every displayed task needs readable guidance"
        descriptions.setdefault(label, set()).add(text)
        selects = heading.getparent().xpath('./select')
        if selects:
            assert len(selects) == 1 and task_label.tag == "label"
            assert task_label.get("for") == selects[0].get("id"), (
                "the original task label must still activate its assignment control"
            )
        else:
            assert task_label.tag == "span", "read-only task text should not label an unrelated control"
    return buttons, descriptions


def test_catalog_covers_actual_roles_and_phases_without_obsolete_proposal_roles():
    expected = set(db.ROLE_SLOT_COUNT) | set(db.BLOCK_PHASES)
    assert set(db.TASK_DESCRIPTIONS) == expected, (
        "task additions and renames must update the description catalog alongside supported roles"
    )
    assert all(isinstance(text, str) and text.strip() for text in db.TASK_DESCRIPTIONS.values())
    assert "Unterstützung" not in db.TASK_DESCRIPTIONS and db.ROLE_MV not in db.TASK_DESCRIPTIONS
    assert db.ROLE_CASH in expected and db.ROLE_CLEANING in expected and db.BLOCK_CAKE_DELIVERY in expected


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_guest_member_and_admin_render_all_task_help_with_matching_shared_descriptions(_today):
    app, ids = _site()
    guest_descriptions = None
    for person_id in (None, ids["assigned"], ids["admin"]):
        tree = html.fromstring(_client(app, person_id).get("/").get_data(as_text=True))
        buttons, descriptions = _task_controls(tree)
        assert len(buttons) == 5 + 8 + 3 + 4 + 3, (
            "youth/adult duties and every preparation, cake and cleanup position need information controls"
        )
        for role, count in db.ROLE_SLOT_COUNT.items():
            for slot in range(count):
                assert descriptions[db.position_label(role, slot)] == {db.TASK_DESCRIPTIONS[role]}
        for phase in db.BLOCK_PHASES:
            capacity = 4 if phase == db.BLOCK_CAKE_DELIVERY else db.BLOCK_SLOT_COUNT
            for slot in range(capacity):
                label = f"{db.BLOCK_PHASE_LABELS[phase]} {slot + 1}"
                assert descriptions[label] == {db.TASK_DESCRIPTIONS[phase]}
        if person_id is None:
            guest_descriptions = descriptions
        else:
            assert descriptions == guest_descriptions, "public task guidance must not depend on access tier"
            assert len(tree.xpath('//select[@data-role or @data-block-assignment]')) == len(buttons)


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_guest_help_adds_generic_text_without_contact_roster_or_internal_identifier_leaks(_today):
    app, _ids = _site()
    page = _client(app).get("/").get_data(as_text=True)
    tree = html.fromstring(page)
    buttons, _descriptions = _task_controls(tree)
    assert "Public Assigned Helper" in page, "existing public assignment names must remain visible"
    for private in ("private-help@fixture.test", "+4915112345678", "help-admin@fixture.test",
                    "roster-help@fixture.test", "Private Unassigned Roster", "1990-01-01"):
        assert private not in page, "generic task help must preserve schedule privacy: " + private
    assert not tree.xpath('//select[@data-role or @data-block-assignment]')
    for internal in ("data-game=", "data-block=", "data-block-card=", "data-game-card=",
                     "data-occupant-id=", "data-block-slot=", "data-candidate-url=", "person_id",
                     "41001", "41002", "41003"):
        assert internal not in page, "guest help must not expose assignment identifiers: " + internal
    for button in buttons:
        assert set(button.attrib) <= {"type", "class", "data-task-help", "aria-label", "aria-controls",
                                     "aria-describedby", "aria-expanded"}, (
            "generic information controls must not carry assignment or person metadata"
        )


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_catalog_text_is_escaped_and_not_interpreted_as_markup(_today):
    app, ids = _site()
    canary = '<script data-help-injection="yes">unsafe()</script> & "quoted"'
    with patch.dict(db.TASK_DESCRIPTIONS, {db.ROLE_SALE: canary}):
        for person_id in (None, ids["admin"]):
            page = _client(app, person_id).get("/").get_data(as_text=True)
            tree = html.fromstring(page)
            assert not tree.xpath('//script[@data-help-injection]'), "description text must never become active HTML"
            _buttons, descriptions = _task_controls(tree)
            assert descriptions["Verkauf 1"] == descriptions["Verkauf 2"] == {canary}
            assert "&lt;script" in page and "&amp;" in page, "plain-text descriptions must be HTML escaped"


@patch("common.effective_today", return_value=date(2026, 9, 1))
def test_saved_cake_responses_supply_description_for_every_dynamic_position(_today):
    app, ids = _site()
    client = _client(app, ids["admin"])
    response = client.post(f'/api/blocks/{ids["cake"]}/cake-settings', json={
        "cake_quantity": 5, "delivery_time": "10:30",
        "expected_cake_quantity": 4, "expected_delivery_time": "10:00",
    }, headers=h.csrf_headers())
    assert response.status_code == 200
    saved = response.get_json()["block"]
    candidates = client.get(f'/api/blocks/{ids["cake"]}/candidates').get_json()["block"]
    for block in (saved, candidates):
        assert block["description"] == db.TASK_DESCRIPTIONS[db.BLOCK_CAKE_DELIVERY]
        assert len(block["slots"]) == 5
        assert {slot["description"] for slot in block["slots"]} == {block["description"]}, (
            "new or refreshed cake controls must receive the same semantic guidance from saved state"
        )


if __name__ == "__main__":
    h.run_all(dict(globals()))
