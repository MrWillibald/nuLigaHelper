# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Shared constants and small helpers used across all modules
# ---------------------------------------------------------------

import datetime
import json
import os


def load_config(filename: str = "config.json") -> dict:
    """Load config.json located next to the project files."""
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
    with open(path, encoding="utf-8") as f:
        return json.load(f)

# Version string
VERSION = "0.30"

# Debug flag: disables all outbound mail/SMS and fixes "today" for testing
DEBUG_FLAG = False

# Change day flag: fixes "today" to the date below (for testing notifications)
CHANGE_DAY = False
DEBUG_TODAY = datetime.date(2025, 11, 21)


def season_year_for(today: datetime.date) -> int:
    """Return the season start year (seasons run from July to June)."""
    return today.year if today.month >= 7 else today.year - 1


def effective_today() -> datetime.date:
    """Return today's date, overridden in debug mode."""
    if DEBUG_FLAG or CHANGE_DAY:
        return DEBUG_TODAY
    return datetime.date.today()
