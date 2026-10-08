"""Responsive schedule/statistics presentation preserves effective data and privacy."""

import os
import tempfile
from datetime import date
from unittest.mock import patch
from urllib.parse import urlencode

from lxml import html

import helpers as h
import db
import webapp


TODAY = date(2026, 10, 8)


def _site(populate=True):
    path = os.path.join(h._TEST_DIR, f"mobile-statistics-{next(tempfile._get_candidate_names())}.db")
    engine = db.make_engine(path)
    db.initialize_db(engine)
    previous = os.environ["NULIGAHELPER_DB"]
    os.environ["NULIGAHELPER_DB"] = path
    try:
        app = webapp.create_app()
    finally:
        os.environ["NULIGAHELPER_DB"] = previous
    with h.Session(engine) as session:
        admin = db.Person(name="Fixture Administrator", birth_date=h.ADULT_BIRTH_DATE,
                          email="mobile-admin@fixture.invalid", is_admin=True)
        session.add(admin)
        session.flush()
        ids = {"admin": admin.id}
        if populate:
            games = [game for game in h.sample_games() if game["game_nr"] in ("1001", "1005")]
            db.sync_games(session, games, h.SEASON)
            support = db.get_support_team(session)
            playing = session.query(db.Team).filter_by(name="BL mC").one()
            future = session.query(db.Game).filter_by(game_nr="1005").one()
            past = session.query(db.Game).filter_by(game_nr="1001").one()
            future.team = support
            seller = db.Person(name="Seller <Fixture>", birth_date=h.ADULT_BIRTH_DATE,
                               email="mobile-seller@fixture.invalid", phone="+491701234567",
                               teams=[support, playing])
            legacy = db.Person(name="Legacy Fixture", birth_date=None, teams=[playing])
            private = db.Person(name="Unassigned Private Fixture", birth_date=h.ADULT_BIRTH_DATE,
                                email="mobile-private@fixture.invalid")
            session.add_all([seller, legacy, private])
            session.flush()
            db.claim_slot(session, future, db.ROLE_SALE, 0, None, seller)
            # Legacy assignments survive eligibility reevaluation and remain counted.
            session.add(db.Assignment(game=future, person=legacy,
                                      role=db.ROLE_TIMEKEEPER, slot=0))
            session.add(db.Assignment(game=past, person=seller,
                                      role=db.ROLE_CASH, slot=0))
            cake = next(block for block in db.get_day_blocks(session, h.SEASON, future.date)
                        if block.phase == db.BLOCK_CAKE_DELIVERY)
            db.configure_cake_block(session, cake, "16:00", 2, None, None)
            db.claim_block_slot(session, cake, 0, None, seller)
            ids.update(playing=playing.id, support=support.id, seller=seller.id)
        session.commit()
    return app, engine, ids


def _page(app, path="/", person_id=None):
    client = app.test_client()
    if person_id is not None:
        h.sign_in(client, person_id)
    with patch.object(webapp.common, "effective_today", return_value=TODAY):
        response = client.get(path)
    assert response.status_code == 200, "synthetic presentation page should remain reachable"
    return html.fromstring(response.get_data(as_text=True)), response.get_data(as_text=True)


def _stat_sections(tree):
    return tree.xpath('//details[contains(concat(" ", normalize-space(@class), " "), " statistics-disclosure ")]')


def test_schedule_intro_and_filters_use_native_expanded_fallback_without_candidate_hooks():
    app, _engine, ids = _site()
    for viewer in (None, ids["seller"]):
        tree, text = _page(app, person_id=viewer)
        disclosures = tree.xpath('//details[@data-responsive-disclosure]')
        assert len(disclosures) == 2, "only general hints and filters receive responsive defaults"
        assert all(node.get("open") is not None and node.get("data-desktop-open") == "true"
                   for node in disclosures), "desktop/no-JavaScript content must start available"
        assert all(not node.xpath('.//*[@data-candidate-url or @data-block-card]')
                   for node in disclosures), "general disclosures must not load assignment candidates"
        assert disclosures[0].xpath('./summary')[0].text_content() == "Hinweise"
        assert disclosures[1].xpath('./summary')[0].text_content() == "Filter"
        form = disclosures[1].xpath('.//form')[0]
        assert form.get("method") == "get" and form.get("action") == "/"
        assert not tree.xpath('//p[@aria-label="Aktive Spielplanfilter"]')
        if viewer:
            assert "Mindestalter am Spieltag" in disclosures[0].text_content()
        assert tree.xpath('//div[@data-eligibility-status and not(@hidden)]')
        assert not disclosures[0].xpath('.//*[@data-eligibility-status or @data-cake-status]'), (
            "operational warnings must keep their game/day-block context"
        )
        assert "mobile-seller@fixture.invalid" not in text and "+491701234567" not in text
        assert "1990-01-01" not in text and "birth_date" not in text
        assert "Unassigned Private Fixture" not in text, "disclosure must not add a roster payload"


def test_combined_schedule_filters_keep_safe_outside_summary_and_selected_values():
    app, _engine, ids = _site()
    selected = {"date": "01.11.2026", "playing_team": f'team-{ids["playing"]}',
                "responsible_team": f'team-{ids["support"]}', "person": "Seller <Fixture>"}
    tree, text = _page(app, "/?" + urlencode(selected))
    summary = tree.xpath('//p[@aria-label="Aktive Spielplanfilter"]')[0]
    assert not summary.xpath('ancestor::details'), "applied filters must remain apparent while collapsed"
    content = summary.text_content()
    assert "01.11.2026" in content and "BL mC" in content and "Supporter" in content
    assert "Seller <Fixture>" in content and not summary.xpath('.//Fixture')
    assert selected["playing_team"] not in content and selected["responsible_team"] not in content
    assert summary.xpath('./a[@href="/"]/text()') == ["Filter löschen"]
    for name in ("date", "playing_team", "responsible_team"):
        assert tree.xpath(f'//select[@name="{name}"]/option[@selected]/@value') == [selected[name]]
    assert tree.xpath('//input[@name="person"]/@value') == [selected["person"]]
    assert "Nr. 1005" in text and "Nr. 1001" not in text
    assert not tree.xpath('//details[contains(@class,"task-block-card")]'), (
        "responsible-team filters must keep excluding team-independent day blocks"
    )


def test_unknown_schedule_filters_keep_selected_values_without_echoing_identifiers_in_summary():
    app, _engine, _ids = _site()
    selected = {"date": "unknown <date>", "playing_team": "team-999999",
                "responsible_team": "unknown <team>", "person": "<script>synthetic</script>"}
    tree, text = _page(app, "/?" + urlencode(selected))
    summary = tree.xpath('//p[@aria-label="Aktive Spielplanfilter"]')[0]
    assert "Unbekannter Spieltag" in summary.text_content()
    assert summary.text_content().count("Unbekannte Mannschaft") == 2
    assert "team-999999" not in summary.text_content() and "unknown <team>" not in summary.text_content()
    assert not summary.xpath('.//script'), "entered names must be escaped in the outside summary"
    for name in ("date", "playing_team", "responsible_team"):
        option = tree.xpath(f'//select[@name="{name}"]/option[@selected]')[0]
        assert option.get("value") == selected[name], "unknown filters must survive reopening and submission"
        assert option.text_content().startswith("Unbekannt")
    assert "Keine Spiele oder Tagesdienste für diese Filter gefunden." in text
    assert "Nr. 1005" not in text
    restored, _text = _page(app, summary.xpath('./a/@href')[0])
    assert not restored.xpath('//p[@aria-label="Aktive Spielplanfilter"]')
    assert restored.xpath('//div[contains(@class,"game-meta") and contains(.,"Nr. 1005")]')


def test_past_schedule_date_stays_expanded_and_identified_outside_filter():
    app, _engine, _ids = _site()
    tree, text = _page(app, "/?date=05.09.2026")
    assert tree.xpath('//details[@class="past-details" and @open]'), (
        "a selected past date must remain visible independently of the filter disclosure"
    )
    assert "05.09.2026" in tree.xpath('//p[@aria-label="Aktive Spielplanfilter"]')[0].text_content()
    assert "Nr. 1001" in text and "Nr. 1005" not in text


def test_statistics_disclosures_keep_order_fallback_defaults_and_authoritative_counts():
    app, _engine, ids = _site()
    tree, text = _page(app, "/statistik", ids["seller"])
    sections = _stat_sections(tree)
    assert len(sections) == 3
    titles = [section.xpath('./summary/h3')[0].text_content() for section in sections]
    assert titles[0].startswith("Spiele pro Mannschaft") and "(3 Mannschaften)" in titles[0]
    assert titles[1].startswith("Dienste pro Person") and "(2 Personen mit Diensten)" in titles[1]
    assert titles[2].startswith("Offene Dienste, Altersanforderungen und Einrichtung")
    assert "(4 betroffene Spiele oder Tagesblöcke)" in titles[2]
    assert [section.get("data-desktop-open") for section in sections] == ["true", "false", "false"]
    assert [section.get("open") is not None for section in sections] == [True, False, False], (
        "team statistics alone must start open in the native desktop fallback"
    )
    assert all(not section.get("name") for section in sections), "sections must not form an exclusive accordion"
    rows = sections[1].xpath('.//tbody/tr')
    seller = next(row for row in rows if "Seller <Fixture>" in row.text_content())
    assert seller.xpath('./td[@data-label="Dienste"]/text()') == ["3"]
    assert "1× Verkauf" in seller.text_content() and "1× Kasse" in seller.text_content()
    assert "1× Kuchenlieferung" in seller.text_content(), "retained and day-block duties keep counting"
    gaps = sections[2].xpath('.//tbody/tr')
    assert len(gaps) == 4, "each affected game/day block contributes exactly one summary count"
    game = next(row for row in gaps if "TSV Übersee" in row.text_content())
    assert len(game.xpath('.//span[contains(@class,"chip-gap")]')) > 1, (
        "multiple vacancies and age deficiency must stay separate within the single game"
    )
    assert "Geburtsdatum" in game.text_content()
    assert "mobile-seller@fixture.invalid" not in text and "+491701234567" not in text
    assert "1990-01-01" not in text and "birth_date" not in text


def test_statistics_setup_needed_remains_separate_and_does_not_add_containers():
    app, engine, ids = _site()
    with h.Session(engine) as session:
        cake = next(block for block in db.get_day_blocks(session, h.SEASON, "01.11.2026")
                    if block.phase == db.BLOCK_CAKE_DELIVERY)
        cake.assignments.clear()
        cake.delivery_time = None
        cake.cake_quantity = None
        session.commit()
    tree, _text = _page(app, "/statistik", ids["admin"])
    outstanding = _stat_sections(tree)[2]
    assert "(4 betroffene Spiele oder Tagesblöcke)" in outstanding.xpath('./summary')[0].text_content()
    cake_row = outstanding.xpath('.//tbody/tr[contains(.,"Kuchenlieferung")]')[0]
    assert cake_row.xpath('.//div[@class="gap-list"]/span/text()') == ["Admin-Einrichtung erforderlich"]


def test_empty_statistics_keep_zero_headers_and_native_reachable_empty_states():
    app, _engine, ids = _site(populate=False)
    # No displayed teams exercises the template's otherwise uncommon all-zero state.
    with patch.object(webapp.db, "get_all_teams", return_value=[]):
        tree, _text = _page(app, "/statistik", ids["admin"])
    sections = _stat_sections(tree)
    assert len(sections) == 3, "empty datasets must not remove section identities"
    assert "(0 Mannschaften)" in sections[0].xpath('./summary')[0].text_content()
    assert "(0 Personen mit Diensten)" in sections[1].xpath('./summary')[0].text_content()
    assert "(0 betroffene Spiele oder Tagesblöcke)" in sections[2].xpath('./summary')[0].text_content()
    assert "Noch keine Mannschaften vorhanden." in sections[0].text_content()
    assert "Noch keine Dienste vergeben." in sections[1].text_content()
    assert "Alle anstehenden Spiele und Tagesblöcke sind vollständig besetzt" in sections[2].text_content()
    assert all(section.xpath('./div[@class="disclosure-content"]') for section in sections)
    intro = tree.xpath('//details[contains(@class,"page-introduction")]')[0]
    assert intro.get("open") is not None and intro.get("data-desktop-open") == "true"
    assert "Noch keine Spiele vorhanden." not in intro.text_content(), (
        "the actionable empty-season notice must remain outside general hints"
    )


if __name__ == "__main__":
    h.run_all(dict(globals()))
