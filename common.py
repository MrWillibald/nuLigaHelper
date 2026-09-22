# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Shared constants and small helpers used across all modules
# ---------------------------------------------------------------

import datetime
import json
import os


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
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
    with open(path, encoding="utf-8") as f:
        config = json.load(f)
    club = config["club"]
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
                from production import text
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
    if os.environ.get("NULIGAHELPER_ENV") == "production":
        from zoneinfo import ZoneInfo
        return datetime.datetime.now(ZoneInfo("Europe/Berlin")).date()
    return datetime.date.today()
