"""Spielfest plan and administration presentation."""

import os
import tempfile

import helpers as h

import db
import webapp


def test_spielfest_is_one_labelled_plan_entry_without_fake_matchup():
    previous = os.environ["NULIGAHELPER_DB"]
    path = os.path.join(h._TEST_DIR, f"spf-web-{next(tempfile._get_candidate_names())}.db")
    os.environ["NULIGAHELPER_DB"] = path
    try:
        db.initialize_db(db.make_engine(path))
        app = webapp.create_app()
    finally:
        os.environ["NULIGAHELPER_DB"] = previous
    engine = db.make_engine(path)
    with h.Session(engine) as session:
        session.add(db.Game(
            season_year=h.SEASON,
            game_nr="SPF:2099-11-28:spf mini",
            day="Sa", date="28.11.2099", time="09:00", hall=280345,
            ak="SPF Mini", home="Spielfest", guest="", score="",
        ))
        session.commit()

    page = app.test_client().get("/").get_data(as_text=True)
    assert page.count("Spielfest SPF Mini") == 1
    assert "Spielfest &ndash;" not in page
    assert "SPF:2099-11-28:spf mini" in page


if __name__ == "__main__":
    h.run_all(dict(globals()))
