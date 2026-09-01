"""Daily-job event filtering tests."""

import helpers as h
import db
import main


def test_new_game_filter_keeps_duplicate_numbers_independent():
    events = db.SyncEvents(new_games=[
        db.GameEvent(1, 555, "meeting:101", "GE"),
        db.GameEvent(2, 555, "meeting:102", "BL mD"),
    ])
    result = main.reportable_new_games(events)
    assert [event.game_id for event in result] == [2]


if __name__ == "__main__":
    h.run_all(dict(globals()))
