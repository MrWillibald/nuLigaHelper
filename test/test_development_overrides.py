# ---------------------------------------------------------------
#                          nuLigaHelper – tests
# ---------------------------------------------------------------
# Development overrides: NULIGAHELPER_TODAY pins the effective date,
# NULIGAHELPER_DRY_RUN suppresses providers, and production refuses
# both before any work.
# ---------------------------------------------------------------

import contextlib
import datetime
import io
import os
import tempfile
from email.message import EmailMessage
from pathlib import Path
from unittest.mock import patch

import helpers as h
from nuligahelper import common
from nuligahelper.jobs import daily as main
from nuligahelper.jobs import notifier
from nuligahelper.ops import monitoring as operations
from nuligahelper.ops import runtime_check
from nuligahelper import web as webapp

PRODUCTION = {
    "NULIGAHELPER_ENV": "production",
    "NULIGAHELPER_SECRET": "synthetic-stable-production-secret",
    "NULIGAHELPER_TRUSTED_HOSTS": "nuliga.example.invalid",
}


def _refused(action, variable):
    try:
        action()
    except common.OverrideError as exc:
        assert variable in str(exc), f"the diagnostic must name {variable}"
        return exc
    raise AssertionError(f"{variable} must be refused")


def _sender():
    sender = notifier.Notifier.__new__(notifier.Notifier)
    sender.smtpserver = "smtp.example.invalid"
    sender.twilio_sid = "synthetic-sid"
    sender.twilio_token = "synthetic-token"
    sender.twilio_service_ID = "synthetic-service"
    return sender


def test_today_override_pins_effective_date_and_unset_uses_real_date():
    with patch.dict(os.environ, {"NULIGAHELPER_TODAY": "2025-11-21"}):
        assert common.effective_today() == datetime.date(2025, 11, 21)
        assert common.season_year_for(common.effective_today()) == 2025
    with patch.dict(os.environ, {"NULIGAHELPER_TODAY": ""}):
        assert common.effective_today() == datetime.date.today(), (
            "an empty override behaves like an unset one"
        )
    assert "NULIGAHELPER_TODAY" not in os.environ
    assert common.effective_today() == datetime.date.today()


def test_invalid_override_values_stop_with_the_variable_name():
    for value in ("21.11.2025", "2025-02-30", "tomorrow"):
        with patch.dict(os.environ, {"NULIGAHELPER_TODAY": value}):
            exc = _refused(common.effective_today, "NULIGAHELPER_TODAY")
            assert value not in str(exc), "the raw value is not echoed"
    for value in ("true", "0", "yes"):
        with patch.dict(os.environ, {"NULIGAHELPER_DRY_RUN": value}):
            _refused(common.dry_run, "NULIGAHELPER_DRY_RUN")
    with patch.dict(os.environ, {"NULIGAHELPER_DRY_RUN": "1"}):
        assert common.dry_run() is True
    assert common.dry_run() is False


def test_dry_run_suppresses_providers_without_pinning_the_date():
    sender = _sender()
    with patch.dict(os.environ, {"NULIGAHELPER_DRY_RUN": "1"}), \
            patch("nuligahelper.jobs.notifier.smtplib.SMTP_SSL") as smtp, patch("nuligahelper.jobs.notifier.Client") as twilio:
        assert common.effective_today() == datetime.date.today(), (
            "dry run alone must not change the effective date"
        )
        assert sender.send_Mail(EmailMessage(), "user", "password") is None
        assert sender.send_SMS("+491700000000", "Synthetischer Text") is None
        smtp.assert_not_called()
        twilio.assert_not_called()


def test_date_override_alone_still_hands_messages_to_providers():
    sender = _sender()
    with patch.dict(os.environ, {"NULIGAHELPER_TODAY": "2025-11-21"}), \
            patch("nuligahelper.jobs.notifier.smtplib.SMTP_SSL") as smtp, patch("nuligahelper.jobs.notifier.Client") as twilio:
        sender.send_Mail(EmailMessage(), "user", "password")
        sender.send_SMS("+491700000000", "Synthetischer Text")
        smtp.assert_called_once_with("smtp.example.invalid")
        twilio.assert_called_once_with("synthetic-sid", "synthetic-token")


def test_production_refuses_overrides_in_every_entry_point():
    for variable, value in (("NULIGAHELPER_TODAY", "2025-11-21"), ("NULIGAHELPER_DRY_RUN", "1")):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "never-created.db"
            values = {**PRODUCTION, "NULIGAHELPER_DB": str(path), variable: value}
            with patch.dict(os.environ, values), \
                    patch("nuligahelper.db.make_engine", side_effect=AssertionError("database opened")):
                _refused(webapp.create_app, variable)
                _refused(main.main, variable)
                accessor = common.effective_today if variable == "NULIGAHELPER_TODAY" else common.dry_run
                _refused(accessor, variable)
                assert runtime_check.check() == variable, (
                    "ExecStartPre must name the override before any unit runs"
                )
            assert not path.exists(), "a refused start must not create the database"


def test_production_preflight_names_the_override():
    config = str(Path(h.PROJECT_DIR) / "test/fixtures/operations.json")
    for variable, value in (("NULIGAHELPER_TODAY", "2025-11-21"), ("NULIGAHELPER_DRY_RUN", "1")):
        for action in ("preflight", "launch-check"):
            output = io.StringIO()
            with patch.dict(os.environ, {**PRODUCTION, variable: value}), \
                    contextlib.redirect_stdout(output):
                code = operations.main([action, "--config", config])
            assert code == 2, f"{action} must fail with {variable} set"
            assert output.getvalue().strip() == f"operation=failed field={variable}"


def test_production_without_overrides_uses_berlin_date_and_sends():
    sender = _sender()
    with patch.dict(os.environ, {"NULIGAHELPER_ENV": "production"}), \
            patch("nuligahelper.jobs.notifier.smtplib.SMTP_SSL") as smtp:
        common.refuse_overrides_in_production()
        assert common.dry_run() is False
        from zoneinfo import ZoneInfo
        assert common.effective_today() == datetime.datetime.now(ZoneInfo("Europe/Berlin")).date()
        sender.send_Mail(EmailMessage(), "user", "password")
        smtp.assert_called_once()


if __name__ == "__main__":
    h.run_all(dict(globals()))
