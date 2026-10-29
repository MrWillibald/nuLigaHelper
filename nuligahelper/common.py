# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Shared constants and small helpers used across all modules
# ---------------------------------------------------------------

import datetime
import json
import os

from nuligahelper import PROJECT_DIR
from nuligahelper.jobs.message_config import referee_targets, warn_legacy_message_settings


PROVIDER_ENV = {
    "email": {
        "smtpserver": "NULIGAHELPER_EMAIL_SMTP_SERVER",
        "mail_ID": "NULIGAHELPER_EMAIL_USER",
        "mail_password": "NULIGAHELPER_EMAIL_PASSWORD",
        "mail_name": "NULIGAHELPER_EMAIL_SENDER_NAME",
        "mailAddrNewspaper": "NULIGAHELPER_EMAIL_NEWSPAPER_RECIPIENT",
        "mailAddrAdmin": "NULIGAHELPER_EMAIL_ADMIN_RECIPIENT",
        "mail_saleID": "NULIGAHELPER_EMAIL_SALE_USER",
        "mail_salePassword": "NULIGAHELPER_EMAIL_SALE_PASSWORD",
    },
    "twilio": {
        "twilio_sid": "NULIGAHELPER_TWILIO_ACCOUNT_SID",
        "twilio_token": "NULIGAHELPER_TWILIO_AUTH_TOKEN",
        "twilio_ID": "NULIGAHELPER_TWILIO_SENDER_ID",
        "twilio_service_ID": "NULIGAHELPER_TWILIO_SERVICE_SID",
    },
    "dropbox": {
        "dropbox_token": "NULIGAHELPER_DROPBOX_TOKEN",
        "dropbox_folder": "NULIGAHELPER_DROPBOX_FOLDER",
        "dated_retention": "NULIGAHELPER_DROPBOX_DATED_RETENTION",
    },
}


def load_config(filename: str | None = None) -> dict:
    """Load non-secret JSON settings and provider settings from the environment."""
    filename = filename or os.environ.get("NULIGAHELPER_CONFIG", "config.json")
    path = os.path.join(PROJECT_DIR, filename)
    with open(path, encoding="utf-8") as f:
        config = json.load(f)
    club = config["club"]
    warn_legacy_message_settings(club)
    production_mode = os.environ.get("NULIGAHELPER_ENV") == "production"
    if production_mode and any(section in club for section in PROVIDER_ENV):
        raise ValueError("NULIGAHELPER_CONFIG")
    for section, fields in PROVIDER_ENV.items():
        env_present = any(name in os.environ for name in fields.values())
        if not production_mode and not env_present:
            continue  # Legacy local configuration and synthetic test fixtures.
        values = {}
        for key, name in fields.items():
            value = os.environ.get(name)
            if not value or value.strip() != value or "\n" in value:
                raise ValueError(name)
            if production_mode:
                from nuligahelper.ops.settings import text
                text(value, name)
            values[key] = value
        if section == "dropbox":
            try:
                values["dated_retention"] = int(values["dated_retention"])
                if values["dated_retention"] < 1:
                    raise ValueError
            except ValueError:
                raise ValueError("NULIGAHELPER_DROPBOX_DATED_RETENTION") from None
        club[section] = values
    return config

# Version string
VERSION = "0.30"

# Local development overrides, read from the environment on every use and
# refused in production (see refuse_overrides_in_production()):
#   NULIGAHELPER_TODAY=YYYY-MM-DD  pins "today" for all date-dependent decisions
#   NULIGAHELPER_DRY_RUN=1         suppresses outbound mail and SMS
TODAY_ENV = "NULIGAHELPER_TODAY"
DRY_RUN_ENV = "NULIGAHELPER_DRY_RUN"
OVERRIDE_ENV = (TODAY_ENV, DRY_RUN_ENV)


class OverrideError(RuntimeError):
    """Raised for an invalid development override or one set in production."""


def season_year_for(today: datetime.date) -> int:
    """Return the season start year (seasons run from July to June)."""
    return today.year if today.month >= 7 else today.year - 1


def _production_environment() -> bool:
    return os.environ.get("NULIGAHELPER_ENV", "").strip().lower() == "production"


def pinned_today() -> datetime.date | None:
    """Return the NULIGAHELPER_TODAY override, or None when it is unset."""
    raw = os.environ.get(TODAY_ENV)
    if raw is None or raw == "":
        return None
    refuse_overrides_in_production()
    try:
        return datetime.date.fromisoformat(raw.strip())
    except ValueError:
        raise OverrideError(
            f"{TODAY_ENV} muss ein gültiges Datum im Format JJJJ-MM-TT sein."
        ) from None


def dry_run() -> bool:
    """Return True when NULIGAHELPER_DRY_RUN=1 suppresses outbound mail and SMS."""
    raw = os.environ.get(DRY_RUN_ENV)
    if raw is None or raw == "":
        return False
    refuse_overrides_in_production()
    if raw.strip() == "1":
        return True
    raise OverrideError(f"{DRY_RUN_ENV} darf nur 1 sein oder fehlen.")


def refuse_overrides_in_production() -> None:
    """Fail before any work when a development override reaches production."""
    if not _production_environment():
        return
    for name in OVERRIDE_ENV:
        if os.environ.get(name) not in (None, ""):
            raise OverrideError(f"{name} ist im Produktionsbetrieb nicht erlaubt.")


def effective_today() -> datetime.date:
    """Return today's date, or the NULIGAHELPER_TODAY override."""
    pinned = pinned_today()
    if pinned is not None:
        return pinned
    if os.environ.get("NULIGAHELPER_ENV") == "production":
        from zoneinfo import ZoneInfo
        return datetime.datetime.now(ZoneInfo("Europe/Berlin")).date()
    return datetime.date.today()
