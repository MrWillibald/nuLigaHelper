# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Display helpers and constants shared by pages and templates.
# ---------------------------------------------------------------

from datetime import datetime


from nuligahelper import db



MONATE = [
    "Januar", "Februar", "März", "April", "Mai", "Juni",
    "Juli", "August", "September", "Oktober", "November", "Dezember",
]

SLOT_LABELS = [
    (db.position_label(role, slot), role, slot)
    for role, count in db.ROLE_SLOT_COUNT.items()
    for slot in range(count)
]

COUNTRY_CODES = [
    ("+49", "Deutschland (+49)"),
    ("+43", "Österreich (+43)"),
    ("+41", "Schweiz (+41)"),
    ("+39", "Italien (+39)"),
    ("+33", "Frankreich (+33)"),
    ("+420", "Tschechien (+420)"),
    ("custom", "Andere Ländervorwahl"),
]


def ak_color(ak: str | None) -> str:
    """Map an age class (e.g. 'BL mD') to the club's team colors."""
    parts = (ak or "").split()
    gender = parts[1][0].lower() if len(parts) > 1 else ""
    youth = len(parts) > 1 and len(parts[1]) > 1
    if gender == "w":
        return "#00C6D7" if youth else "#7CFFCB"
    if gender == "m":
        return "#6BB32C" if youth else "#DDFF00"
    return "#FFA01F"


def parse_date(date_str: str | None):
    try:
        return datetime.strptime(date_str or "", "%d.%m.%Y").date()
    except ValueError:
        return None


def display_time(time_str: str | None) -> str:
    parts = (time_str or "").split()
    return parts[0] if parts else ""
