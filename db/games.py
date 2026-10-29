# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Game lookups, Spielfest identity, display names and date sorting.
# ---------------------------------------------------------------

from __future__ import annotations

import re
from datetime import date, datetime

from sqlalchemy import select
from sqlalchemy.orm import Session

import age_eligibility

from db.models import Game


def get_games_on_date(session: Session, date: str) -> list[Game]:
    """Return all games on the given date string (format dd.mm.yyyy), ordered by time."""
    return list(
        session.scalars(select(Game).where(Game.date == date).order_by(Game.time, Game.game_nr))
    )


def is_spielfest(game_or_age_group: Game | str | None) -> bool:
    """Return whether a game/age group represents a collapsed SPF event."""
    age_group = (
        game_or_age_group.ak
        if isinstance(game_or_age_group, Game)
        else game_or_age_group
    )
    return bool(re.search(r"(?<!\w)spf(?!\w)", age_group or "", re.IGNORECASE))


def game_display_name(game: Game) -> str:
    """Human-facing event name without exposing synthetic matchup placeholders."""
    if is_spielfest(game):
        return f"Spielfest {game.ak or ''}".strip()
    return f"{game.home or '?'} – {game.guest or '?'}"


def game_sort_key(game: "Game") -> tuple:
    """
    Chronological sort key for games with German dd.mm.yyyy date strings
    (plain string ordering would sort by day-of-month first).
    """
    try:
        d = datetime.strptime(game.date or "", "%d.%m.%Y").date()
    except ValueError:
        d = date.max
    raw_time = (game.time or "").split()
    time_parts = raw_time[0].split(":") if raw_time else []
    if len(time_parts) == 2 and all(p.isdigit() for p in time_parts):
        time_key = (int(time_parts[0]), int(time_parts[1]))
    else:
        time_key = (99, 99)
    number_key = (
        (0, int(game.game_nr))
        if str(game.game_nr).isdigit()
        else (1, str(game.game_nr).casefold())
    )
    return (d, time_key, number_key, game.id or 0)


classify_game_category = age_eligibility.classify_game_category
