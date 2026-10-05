"""Offline, non-mutating migration checks; all configuration here is synthetic."""

from contextlib import redirect_stdout
from copy import deepcopy
import io
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
from unittest.mock import patch

import helpers as h
import common
import message_config


def test_preflight_classifies_known_defaults_customized_and_unused_without_values():
    club = {
        "texts": {
            "mailErrorSubject": "Benachrichtigung Fehler",
            "textBlockTask": "Hallo {}, morgen ({}) übernimmst du {} um {}.",
            "mailTask": "private-copy-canary {private-name-canary}",
            "mailEarlyTask": "private-unused-canary",
            "obsoleteCustomKey": "private-unknown-canary",
            "mailRefCoordTargets": [{"Name": "private-name", "Address": "private@contacts.invalid"}],
        }
    }
    before = deepcopy(club)
    report = message_config.preflight(club)
    entries = {row["key"]: row for row in report["legacy"]}
    assert entries["mailErrorSubject"]["status"] == "prior-default"
    assert entries["textBlockTask"]["status"] == "prior-default", "notifier fallback is also a prior default"
    assert entries["mailTask"]["status"] == "customized"
    assert entries["mailEarlyTask"]["destination"] == "removed: no supported consumer"
    assert entries["obsoleteCustomKey"]["status"] == "unrecognized"
    assert entries["mailRefCoordTargets"]["status"] == "recipient-metadata"
    assert club == before, "preflight must not rewrite text or metadata"
    output = json.dumps(report)
    assert "private" not in output, "migration reports must contain no private values"
    assert set(message_config.LEGACY_DESTINATIONS) == set(message_config.LEGACY_DEFAULT_HASHES)


def test_recipient_mapping_preserves_values_and_new_setting_wins_including_empty():
    old = [{"Name": "Old", "Address": "+49123456789"}]
    new = [{"Name": "New", "Address": "new@synthetic.invalid"}]
    club = {"texts": {"mailRefCoordTargets": old}}
    assert common.referee_targets(club) == old
    assert message_config.preflight(club)["referee_source"] == "club.texts.mailRefCoordTargets"
    club["notifications"] = {"referee_targets": new}
    assert common.referee_targets(club) == new
    report = message_config.preflight(club)
    assert report["recipient_conflict"]
    assert report["referee_source"] == "club.notifications.referee_targets"
    club["notifications"]["referee_targets"] = []
    assert common.referee_targets(club) == [], "an explicit empty list must not revive old recipients"
    club["notifications"]["referee_targets"] = old
    assert not message_config.preflight(club)["recipient_conflict"]
    assert common.referee_targets({}) == []


def test_legacy_load_is_readable_and_warnings_identify_keys_without_values():
    club = {
        "texts": {"mailTask": "private-copy-canary", "mailRefCoordTargets": [{"Address": "old-private-canary"}]},
        "notifications": {"referee_targets": [{"Address": "new-private-canary"}]},
    }
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "synthetic.json"
        path.write_text(json.dumps({"club": club}), encoding="utf-8")
        before = path.read_bytes()
        with patch.dict(os.environ, {}, clear=True), patch("message_config.logging.warning") as warning:
            loaded = common.load_config(str(path))
        assert loaded["club"] == club, "presence of legacy text alone must not reject startup"
        assert path.read_bytes() == before
        warnings = " ".join(call.args[0] % call.args[1:] for call in warning.call_args_list)
        assert "club.texts" in warnings and "mailTask" in warnings
        assert "takes precedence" in warnings
        assert "private" not in warnings


def test_preflight_cli_needs_no_database_or_provider_settings_and_changes_nothing():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "synthetic.json"
        path.write_text(json.dumps({"club": {"texts": {"mailTask": "private-copy-canary"}}}), encoding="utf-8")
        before = path.read_bytes()
        env = {"PATH": os.environ.get("PATH", ""), "NULIGAHELPER_DB": str(Path(directory) / "must-not-exist.db")}
        result = subprocess.run(
            [sys.executable, "-m", "message_config", "--config", str(path)],
            cwd=h.PROJECT_DIR, env=env, capture_output=True, text=True,
        )
        assert result.returncode == 0, result.stderr
        assert json.loads(result.stdout)["legacy"][0]["status"] == "customized"
        assert "private" not in result.stdout + result.stderr
        assert path.read_bytes() == before
        assert not Path(env["NULIGAHELPER_DB"]).exists(), "message migration is independent of schema or providers"


def test_bad_configuration_preflight_omits_exception_details_and_contents():
    with tempfile.TemporaryDirectory() as directory:
        path = Path(directory) / "synthetic.json"
        path.write_text("private-invalid-json-canary", encoding="utf-8")
        output = io.StringIO()
        with redirect_stdout(output):
            assert message_config.main(["--config", str(path)]) == 2
        assert "private" not in output.getvalue()


if __name__ == "__main__":
    h.run_all(dict(globals()))
