"""Provider settings are supplied by the production environment only."""

import json
import os
import tempfile
from pathlib import Path
from unittest.mock import patch

import helpers as h
import common


def _values():
    values = {name: "synthetic-value" for fields in common.PROVIDER_ENV.values()
              for name in fields.values()}
    values["NULIGAHELPER_DROPBOX_DATED_RETENTION"] = "14"
    return values


def test_production_provider_settings_come_from_env_and_legacy_json_is_rejected():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "config.json"
        path.write_text(json.dumps({"club": {"texts": {}}}), encoding="utf-8")
        values = {"NULIGAHELPER_ENV": "production", **_values()}
        with patch.dict(os.environ, values, clear=True):
            club = common.load_config(str(path))["club"]
            assert club["email"]["mail_password"] == "synthetic-value"
            assert club["twilio"]["twilio_token"] == "synthetic-value"
            assert club["dropbox"]["dated_retention"] == 14
            del os.environ["NULIGAHELPER_TWILIO_AUTH_TOKEN"]
            try:
                common.load_config(str(path))
            except ValueError as error:
                assert str(error) == "NULIGAHELPER_TWILIO_AUTH_TOKEN"
            else:
                raise AssertionError("production accepted a missing Twilio credential")
            path.write_text(json.dumps({"club": {"email": {"mail_password": "legacy-secret"}}}), encoding="utf-8")
            try:
                common.load_config(str(path))
            except ValueError as error:
                assert str(error) == "NULIGAHELPER_CONFIG"
            else:
                raise AssertionError("production accepted provider credentials in JSON")


def test_production_rejects_example_provider_values_without_echoing_them():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "config.json"
        path.write_text(json.dumps({"club": {}}), encoding="utf-8")
        values = {"NULIGAHELPER_ENV": "production", **_values(),
                  "NULIGAHELPER_DROPBOX_TOKEN": "REPLACE_DROPBOX_TOKEN"}
        with patch.dict(os.environ, values, clear=True):
            try:
                common.load_config(str(path))
            except ValueError as error:
                assert str(error) == "NULIGAHELPER_DROPBOX_TOKEN"
            else:
                raise AssertionError("production accepted a placeholder Dropbox token")
