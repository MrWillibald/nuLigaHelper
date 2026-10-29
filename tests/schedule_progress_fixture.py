"""Loopback-only synthetic schedule for opt-in progress layout verification."""
import hashlib
import json
import logging
import os
from pathlib import Path

from tests import helpers as h
from nuligahelper import db
from nuligahelper import web as webapp
from sqlalchemy import select
from werkzeug.serving import make_server
from tests.db.test_game_duties import _site


def main():
    app, engine, _path, ids = _site()
    webapp.common.effective_today = lambda: h.date(2026, 9, 1)
    with h.Session(engine) as session:
        games = [session.get(db.Game, game_id) for game_id in ids["games"]]
        # Normal partial, all adult roles vacant, full SPF with a retained duty,
        # and exceptionally long identity plus unresolved age coverage.
        for game in games:
            game.hall = 1
        adult = games[0]
        for (role, slot), person_id in zip(db.BASELINE_REQUIRED_POSITIONS, ids["helpers"]):
            session.add(db.Assignment(game=adult, role=role, slot=slot,
                                      person=session.get(db.Person, person_id)))
        games[1].ak = "BL M"
        games[1].date = "21.10.2026"
        games[2].ak = "SPF Mini"
        games[2].date = "22.10.2026"
        for (role, slot), person_id in zip(db.BASELINE_REQUIRED_POSITIONS, ids["helpers"]):
            session.add(db.Assignment(game=games[2], role=role, slot=slot,
                                      person=session.get(db.Person, person_id)))
        session.add(db.Assignment(game=games[2], role=db.ROLE_CASH, slot=0,
                                  person=session.get(db.Person, ids["member"])))
        games[3].date = "23.10.2026"
        games[3].ak = "GE"
        games[3].home = "Handballspielgemeinschaft Oberbayerisches Alpenvorland Raubling"
        games[3].guest = "Spielgemeinschaft SehrlangerGastvereinsname Bergkirchen und Umgebung"
        long_team = db.Team(name="Verantwortliche Mannschaft Spielgemeinschaft Alpenvorland mit langem Namen")
        session.add(long_team)
        games[3].team = long_team
        legacy = db.Person(name="Legacy Fixture", birth_date=None)
        session.add(db.Assignment(game=games[3], person=legacy, role=db.ROLE_TIMEKEEPER, slot=0))
        past = db.Game(season_year=h.SEASON, game_nr="9600", date="30.08.2026", time="00:30",
                       ak="BL wD", home="TuS Raubling", guest="Past Opponent", hall=1)
        session.add(past)
        games.append(past)
        db.reconcile_day_blocks(session, h.SEASON, {game.date for game in games})
        for day, quantity in [("20.10.2026", 4), ("21.10.2026", 0)]:
            block = next(block for block in db.get_day_blocks(session, h.SEASON, day)
                         if block.phase == db.BLOCK_CAKE_DELIVERY)
            db.configure_cake_block(session, block, "10:00", quantity, None, None)
        for phase in (db.BLOCK_PREPARATION, db.BLOCK_CLEANUP, db.BLOCK_CAKE_DELIVERY):
            block = next(block for block in db.get_day_blocks(session, h.SEASON, "20.10.2026") if block.phase == phase)
            db.claim_block_slot(session, block, 0, None, session.get(db.Person, ids["member"]))
        session.commit()

    @app.get('/__snapshot')
    def snapshot():
        with engine.connect() as connection:
            values = {table.name: [list(row) for row in connection.execute(select(table))]
                      for table in db.Base.metadata.sorted_tables}
        return {"hash": hashlib.sha256(json.dumps(values, default=str, sort_keys=True).encode()).hexdigest()}

    before_dir = os.environ.get("PROGRESS_REFERENCE_DIR")
    if before_dir:
        @app.get('/__before')
        def before_page():
            return before

        @app.get('/__before.css')
        def before_css():
            return app.response_class(Path(before_dir, 'style.css').read_text(), mimetype='text/css')

        # Optional reference assets are rendered against the same synthetic data.
        from jinja2 import ChoiceLoader, DictLoader
        original_loader = app.jinja_loader
        app.jinja_loader = ChoiceLoader([DictLoader({"schedule.html":
            Path(before_dir, "schedule.html").read_text()}), original_loader])
        before = app.test_client().get('/').get_data(as_text=True).replace('/static/style.css', '/__before.css')
        app.jinja_loader = original_loader
        app.jinja_env.cache.clear()

    client = app.test_client()
    h.sign_in(client, ids["admin"])
    cookie = client.get_cookie(app.config['SESSION_COOKIE_NAME'])
    logging.getLogger('werkzeug').setLevel(logging.ERROR)
    server = make_server('127.0.0.1', 0, app, threaded=True)
    print('PROGRESS_FIXTURE ' + json.dumps({"port": server.server_port,
          "cookie": {"name": cookie.key, "value": cookie.value}, "ids": ids}), flush=True)
    server.serve_forever()


if __name__ == '__main__':
    main()
