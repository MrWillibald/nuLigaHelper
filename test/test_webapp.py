# ---------------------------------------------------------------
#                          nuLigaHelper – tests
# ---------------------------------------------------------------
# Web interface regression scenario against the sample game plan:
# schedule rendering, inline assignment API, person/team management
# and the statistics page. Uses a throwaway database (see helpers).
#
# Run standalone:  python test/test_webapp.py
# Or via pytest:   pytest test/test_webapp.py
#
# Note: the tests build on each other and run top to bottom.
# ---------------------------------------------------------------

import re

import helpers as h

import db
import webapp

client = webapp.app.test_client()

# ---------------------------------------------------------------------
# Scenario data (see helpers.sample_games)
# ---------------------------------------------------------------------

with h.Session(h.app_db_engine()) as session:
    games = h.sync_sample_games(session)

    # Game 1001 is our scenario game; its own age class identifies the
    # team that PLAYS it. Two further age classes serve as responsible
    # resp. unrelated teams.
    game_1001 = next(g for g in games if g["game_nr"] == 1001)
    AK_PLAYING = game_1001["ak"]
    AK_RESPONSIBLE = next(g["ak"] for g in games if g["ak"] != AK_PLAYING)
    AK_UNRELATED = next(g["ak"] for g in games
                        if g["ak"] not in (AK_PLAYING, AK_RESPONSIBLE))

    support = db.get_support_team(session)
    team_playing = session.query(db.Team).filter_by(name=AK_PLAYING).one()
    team_responsible = session.query(db.Team).filter_by(name=AK_RESPONSIBLE).one()
    team_unrelated = session.query(db.Team).filter_by(name=AK_UNRELATED).one()

    alice = db.get_or_create_person(session, "Alice Test", email="alice@x.de")
    alice.team_id = team_playing.id               # plays in game 1001
    bob = db.get_or_create_person(session, "Bob Test", phone="+491700000002")
    bob.team_id = team_responsible.id             # member of responsible team
    caro = db.get_or_create_person(session, "Caro Test")
    caro.team_id = support.id                     # support helper
    dora = db.get_or_create_person(session, "Dora Test", email="dora@x.de")
    dora.team_id = team_unrelated.id              # unrelated team
    session.commit()

    GAME_ID = session.query(db.Game).filter_by(game_nr=1001).one().id
    ALICE_ID, BOB_ID, CARO_ID, DORA_ID = alice.id, bob.id, caro.id, dora.id
    TEAM_PLAYING_ID, TEAM_RESPONSIBLE_ID = team_playing.id, team_responsible.id
    SUPPORT_ID = support.id


def _card_html(html: str) -> str:
    start = html.find(f'id="game-{GAME_ID}"')
    return html[start:html.find("</article>", start)]


# ---------------------------------------------------------------------


def test_schedule_renders_games_without_contact_data():
    html = client.get("/").get_data(as_text=True)
    assert html.count("<select") == len(games) * 7  # 6 role slots + team select
    assert "TuS Raubling" in html
    assert "alice@x.de" not in html and "+4917" not in html, \
        "contact data must never appear on the overview"


def test_games_start_without_a_responsible_team():
    with h.Session(h.app_db_engine()) as session:
        unlinked = session.query(db.Game).filter(db.Game.team_id.is_(None)).count()
        assert unlinked == len(games), "no team may be pre-assigned"


def test_responsible_team_api_sets_and_clears_the_team():
    r = client.post(f"/api/games/{GAME_ID}/team", json={"team_id": TEAM_RESPONSIBLE_ID})
    assert r.get_json() == {"ok": True}
    with h.Session(h.app_db_engine()) as session:
        game = session.get(db.Game, GAME_ID)
        assert game.judge_team_name == AK_RESPONSIBLE

    r = client.post(f"/api/games/{GAME_ID}/team", json={"team_id": None})
    assert r.get_json() == {"ok": True}
    with h.Session(h.app_db_engine()) as session:
        assert session.get(db.Game, GAME_ID).judge_team_name is None


def test_playing_and_outside_helpers_are_greyed_but_selectable():
    client.post(f"/api/games/{GAME_ID}/team", json={"team_id": TEAM_RESPONSIBLE_ID})
    card = _card_html(client.get("/").get_data(as_text=True))

    assert card.count('class="option-playing"') == 6, \
        "Alice belongs to the playing team and must be marked in all slots"
    assert "spielt in diesem Spiel selbst" in card, \
        "the play-hint tooltip must be present"

    assert card.count('class="foreign-option"') == 6, \
        "Dora belongs to an unrelated team and must be greyed in all slots"
    assert f"gehört zu {AK_UNRELATED}" in card

    # members of the responsible team / support team appear without marks;
    # only the person's own <option> tag is inspected
    for name in ("Bob Test", "Caro Test"):
        if f">{name}<" not in card:
            continue  # may already have been deleted by an earlier test
        pos = card.find(f">{name}<")
        tag_start = card.rfind("<option", 0, pos)
        tag = card[tag_start:card.find(">", tag_start)]
        assert "option-playing" not in tag and "foreign-option" not in tag, \
            f"{name} must not be marked"

    # all categories remain selectable
    for pid, label in ((ALICE_ID, "playing"), (CARO_ID, "support"), (DORA_ID, "outside")):
        r = client.post("/api/assignment", json={
            "game_id": GAME_ID, "role": "Zeitnehmer", "slot": 0, "person_id": pid})
        assert r.get_json() == {"ok": True}, f"{label} pick must be allowed"
    client.post("/api/assignment", json={
        "game_id": GAME_ID, "role": "Zeitnehmer", "slot": 0, "person_id": None})


def test_assignment_api_fills_both_sale_slots_and_rejects_duplicates():
    r = client.post("/api/assignment", json={
        "game_id": GAME_ID, "role": "Verkauf", "slot": 0, "person_id": ALICE_ID})
    assert r.get_json() == {"ok": True}
    r = client.post("/api/assignment", json={
        "game_id": GAME_ID, "role": "Verkauf", "slot": 1, "person_id": DORA_ID})
    assert r.get_json() == {"ok": True}

    with h.Session(h.app_db_engine()) as session:
        sales = [a.person.name
                 for a in session.get(db.Game, GAME_ID).assignments_by_role("Verkauf")]
        assert sales == ["Alice Test", "Dora Test"], sales

    r = client.post("/api/assignment", json={
        "game_id": GAME_ID, "role": "Verkauf", "slot": 1, "person_id": ALICE_ID})
    assert r.get_json()["ok"] is False, "duplicates must be rejected"

    r = client.post("/api/assignment", json={
        "game_id": GAME_ID, "role": "Unbekannt", "slot": 0, "person_id": None})
    assert r.status_code == 400


def test_person_management_crud():
    r = client.get("/personen")
    assert r.status_code == 200
    page = r.get_data(as_text=True)
    assert "Mannschaft" in page and "Support-Team" in page

    r = client.post("/personen/add", data={
        "name": "Ela Test", "email": "ela@x.de", "team_id": SUPPORT_ID},
        follow_redirects=True)
    assert "Ela Test" in r.get_data(as_text=True)
    with h.Session(h.app_db_engine()) as session:
        ela = session.query(db.Person).filter_by(name="Ela Test").one()
        ela_id = ela.id
        assert ela.team_id == SUPPORT_ID

    r = client.post(f"/personen/{ela_id}/edit", data={
        "name": "Ela Test", "email": "", "phone": "+49170999",
        "team_id": SUPPORT_ID}, follow_redirects=True)
    with h.Session(h.app_db_engine()) as session:
        ela = session.get(db.Person, ela_id)
        assert ela.phone == "+49170999" and ela.email is None
        assert ela.team_id == SUPPORT_ID


def test_team_editing_is_completely_disabled():
    with h.Session(h.app_db_engine()) as session:
        any_team = session.query(db.Team).first().id
    assert client.post(f"/teams/{any_team}/delete").status_code == 404
    assert client.post("/teams/add", data={"name": "Neues Team"}).status_code == 404


def test_team_mv_can_be_assigned_to_members_only():
    # UI offers an MV select per team
    page = client.get("/personen").get_data(as_text=True)
    assert "Mannschaftsverantwortlicher" in page
    assert 'class="mv-select"' in page

    # a member can become MV; the selection is reflected in the page
    r = client.post(f"/api/teams/{TEAM_PLAYING_ID}/mv", json={"person_id": ALICE_ID})
    assert r.get_json() == {"ok": True}
    page = client.get("/personen").get_data(as_text=True)
    row = re.search(
        r'team-mv-row[^>]*>\s*<span[^>]*>\s*BL mD.*?selected', page, re.S)
    assert row, "Alice must appear as selected MV of her team"

    # a person of another team is rejected
    r = client.post(f"/api/teams/{TEAM_PLAYING_ID}/mv", json={"person_id": DORA_ID})
    body = r.get_json()
    assert body["ok"] is False and "kein Mitglied" in body["error"]

    # clearing works
    r = client.post(f"/api/teams/{TEAM_PLAYING_ID}/mv", json={"person_id": None})
    assert r.get_json() == {"ok": True}
    with h.Session(h.app_db_engine()) as session:
        team = session.get(db.Team, TEAM_PLAYING_ID)
        assert team.mv_person_id is None


def test_statistics_page_with_bars_and_aggregated_gaps():
    stats = client.get("/statistik").get_data(as_text=True)
    assert "Spiele pro Mannschaft" in stats and "Dienste pro Person" in stats
    assert "Offene Dienste" in stats

    # one bar per team
    with h.Session(h.app_db_engine()) as session:
        n_teams = session.query(db.Team).count()
        total = session.query(db.Game).filter_by(season_year=h.SEASON).count()
    assert stats.count('class="bar-track"') == n_teams
    assert f"{total} Spielen" in stats

    # gaps aggregate both sale slots into a single chip
    assert 'chip-gap">Verkauf' in stats
    assert ">Verkauf 1<" not in stats and ">Verkauf 2<" not in stats

    with h.Session(h.app_db_engine()) as session:
        game = session.get(db.Game, GAME_ID)
        missing_sales = 2 - len(game.assignments_by_role("Verkauf"))
    if missing_sales > 0:
        assert f"Verkauf ({missing_sales}×)" in stats, "gap chip must show the count"


def test_deleting_a_person_removes_their_assignments():
    # give Bob a task first so the deletion has something to cascade
    r = client.post("/api/assignment", json={
        "game_id": GAME_ID, "role": "Sekretär", "slot": 0, "person_id": BOB_ID})
    assert r.get_json() == {"ok": True}

    r = client.post(f"/personen/{BOB_ID}/delete", follow_redirects=True)
    with h.Session(h.app_db_engine()) as session:
        assert session.get(db.Person, BOB_ID) is None
        game = session.get(db.Game, GAME_ID)
        assert game.assignment_by_role("Sekretär") is None, \
            "the deleted person's assignment must be removed as well"


if __name__ == "__main__":
    h.run_all(dict(globals()))
