"""Read-only migration inventory for legacy notification settings.

Fingerprints describe the last positional example and notifier fallback defaults;
they are comparison data only, never a source of renderable message wording.
"""

import argparse
import hashlib
import json
import logging
import os
from pathlib import Path


LEGACY_DESTINATIONS = {
    "mailErrorSubject": "game.new.subject",
    "mailError": "game.new.email",
    "mailNewspaperSubject": "newspaper.article.subject (dormant)",
    "mailNewspaper": "newspaper.article.email (dormant)",
    "mailMVSubject": "mv.game.day_before / mv.spielfest.day_before.subject",
    "mailTask": "game.day_before.email",
    "textTask": "game.day_before.sms",
    "mailMV": "mv.game.day_before.email",
    "textMV": "mv.game.day_before.sms",
    "mailPreTask": "game.weekly.email",
    "textPreTask": "game.weekly.sms",
    "mailEarlyTask": "removed: no supported consumer",
    "textEarlyTask": "removed: no supported consumer",
    "mailPreparationTask": "block.preparation.weekly.email",
    "textPreparationTask": "block.preparation.weekly.sms",
    "mailBlockPreTask": "block.weekly.email",
    "textBlockPreTask": "block.weekly.sms",
    "mailBlockTask": "block.day_before.email",
    "textBlockTask": "block.day_before.sms",
    "mailShifted": "game.shifted.email",
    "textShifted": "game.shifted.sms",
    "mailRefCoordSubject": "referee.missing.subject",
    "mailRefCoord": "referee.missing.email / sms (positional mismatch corrected)",
}

# Populated with immutable SHA-256 comparison data from the prior public example.
LEGACY_DEFAULT_HASHES = {
    'mailErrorSubject': frozenset(['06a9e77dfbdf3c5122b3b8481d01352cbc6a011798b09bfdfdfdd97f7d0b4087']),
    'mailError': frozenset(['42923dbbf95f2d677c8c5f8a2186c63db5ec36463b8834f6180cd58c5e66fbdf']),
    'mailNewspaperSubject': frozenset(['745163aaf1db864f9cd5b912c569dfefda2767f33937af5b2d857de39eb9ef99']),
    'mailNewspaper': frozenset(['b8dd5148c1125b332f3d1dc2ede9c0d60f3e16022d26d3206d1d1aa3fdef7876']),
    'mailMVSubject': frozenset(['179949b2f9307b8408d68eff71888aa37a4a8c0360ad596447cf3feee168fbf0']),
    'mailTask': frozenset(['2089cdd4aabb13850174a9f98ef2eac796370d8e0261653b3c82a097ec1d7b02']),
    'textTask': frozenset(['205dbc246d4a536cdc0a1672c290bcb86e314b29c36b70f46486d6cc4eef28e5']),
    'mailMV': frozenset(['bd6ad028eff970abe1b2dae4ec2eeaa8563bd57fc27862e71d2b150745fdb720']),
    'textMV': frozenset(['70d3431478d2b47d863d8b0e0ecb261170e9aa327f91b39d35b6c7c4d7a499e8']),
    'mailPreTask': frozenset(['1f8ddb84ad47035048147b5344347325045322f009470017c5624a9494e39f93']),
    'textPreTask': frozenset(['c8718d773acebdf59021711985d69c7062b1f7df23f5d0dce4ddc65e003eb6de']),
    'mailEarlyTask': frozenset(['4b59eb45529640e88c4be0a4b015de369b8caa0bc867369deee4ff89dce66bf2']),
    'textEarlyTask': frozenset(['a51dd291a31f63bb1bb4b71a7ab3dad9d51f9c7c597f0e6751f5bf2e1f8a65e5']),
    'mailPreparationTask': frozenset(['3ed6e1806c689d35e1c49ea97307e2fdeb352158bd895954558c4effe0c0b26a', '785de423404493f3242be325367753e164c5a22e9c678c4205ae4aa4275085d8']),
    'textPreparationTask': frozenset(['6aa22222baaa3db22268cd417ccfd818e3e5bff4178b9156a7ac8ba8c7fd5de1', 'd910c3cafbd70680d6f55f04492a7aef2d0f4483f9f83d362e37c6a250de039f']),
    'mailBlockPreTask': frozenset(['2dd44e53bc39b8f8c4362eeecfd3b6cac0496e79211adaf8a69c46dd3b3f4025', '49a40d4eae9c5e079d24fb327ee9117c874fad3ce821066e5e2b0b97c6793aca']),
    'textBlockPreTask': frozenset(['5204926cc49c5ead08f3a9f84e284b0054c4e74b410c31971d988b9b844ae73a', 'add49cce05dd6e6e904bd2cb4464423302f9b25057b501a9dd29b951331ae7d5']),
    'mailBlockTask': frozenset(['4cfa3bd0b43695fd2d30aa05fbe04e3c74a4cbd2160c994b1181cda3c7f8acc1', 'e063dba9dddbaf25f16ba39e1fb05d61350b56f5ed8761bea7be17b6a94bba4f']),
    'textBlockTask': frozenset(['25ca9e9ff3db93a92f8a9a08e7f270a6af578ae126f8b40190b771072f2836bc', '73f505c7e0ad2bc3834822697b610f6a81a21ad11f7283c18b2dba88dc1b5aa4']),
    'mailShifted': frozenset(['5155140cc80768c470faf1ba9cd87fddb252af5882c890b4784df9ce7b190c1f']),
    'textShifted': frozenset(['9f2dde5e61571efa1f69fbe3fad5e86f44344db8edde257ecd87c261bda5f113']),
    'mailRefCoordSubject': frozenset(['ea35f8b53787436f9117b964a15b1d3326e0817b6b0ee10f29f215515ddd538d']),
    'mailRefCoord': frozenset(['884dad7f3538ce35b04b01d903a6d8ba84dcc71efe873e65ca9f6a48e9eef901']),
}


def preflight(club: dict) -> dict:
    """Return only key names, dispositions and statuses; never setting values."""
    legacy = club.get("texts", {})
    if not isinstance(legacy, dict):
        raise ValueError("club.texts")
    notifications = club.get("notifications", {})
    if not isinstance(notifications, dict):
        raise ValueError("club.notifications")
    entries = []
    for key in sorted(legacy):
        if key == "mailRefCoordTargets":
            status = "recipient-metadata"
            destination = "club.notifications.referee_targets"
        else:
            value = legacy[key]
            fingerprint = (
                hashlib.sha256(value.encode("utf-8")).hexdigest()
                if isinstance(value, str) else None
            )
            status = (
                "prior-default" if fingerprint in LEGACY_DEFAULT_HASHES.get(key, ())
                else "customized" if key in LEGACY_DESTINATIONS
                else "unrecognized"
            )
            destination = LEGACY_DESTINATIONS.get(key, "developer review required")
        entries.append({"key": key, "status": status, "destination": destination})
    has_new = "referee_targets" in notifications
    has_old = "mailRefCoordTargets" in legacy
    return {
        "legacy": entries,
        "referee_source": (
            "club.notifications.referee_targets" if has_new else
            "club.texts.mailRefCoordTargets" if has_old else "unset"
        ),
        "recipient_conflict": (
            has_new and has_old
            and notifications["referee_targets"] != legacy["mailRefCoordTargets"]
        ),
    }


def warn_legacy_message_settings(club: dict) -> None:
    """Warn about obsolete keys without logging wording or recipient values."""
    report = preflight(club)
    if "texts" in club:
        logging.warning(
            "Deprecated message settings club.texts keys=%s; wording comes from "
            "messages.py. Run python -m message_config before deployment.",
            json.dumps([item["key"] for item in report["legacy"]]),
        )
    if report["recipient_conflict"]:
        logging.warning(
            "Conflicting referee recipient metadata: "
            "club.notifications.referee_targets takes precedence over "
            "club.texts.mailRefCoordTargets."
        )


def referee_targets(club: dict) -> list:
    """Resolve recipient settings with explicit new-key precedence, even if empty."""
    notifications = club.get("notifications", {})
    if "referee_targets" in notifications:
        return list(notifications["referee_targets"])
    return list(club.get("texts", {}).get("mailRefCoordTargets", []))


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--config", default=os.environ.get("NULIGAHELPER_CONFIG", "config.json"),
        help="Configuration to inspect (no providers or database are opened)",
    )
    args = parser.parse_args(argv)
    path = Path(args.config)
    if not path.is_absolute():
        path = Path(__file__).resolve().parent / path
    try:
        config = json.loads(path.read_text(encoding="utf-8"))
        report = preflight(config["club"])
    except (OSError, ValueError, KeyError, TypeError, AttributeError):
        print("Message preflight: unreadable configuration or invalid settings shape")
        return 2
    print(json.dumps(report, indent=2, ensure_ascii=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
