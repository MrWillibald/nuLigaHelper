# ---------------------------------------------------------------
#                          nuLigaHelper – tests
# ---------------------------------------------------------------
# Project layout guards: code imports the nuligahelper package, test
# patches target package modules, and only the transition shims stay
# at the repository root.
# ---------------------------------------------------------------

import re
from pathlib import Path

from tests import helpers as h

PROJECT = Path(h.PROJECT_DIR)
OLD_MODULES = (
    "common", "db", "age_eligibility", "contact_validation", "task_labels", "schema_migrations",
    "webapp", "auth_abuse", "main", "daily_lock", "scraper", "notifier", "messages",
    "message_config", "backup", "production", "production_logging", "operations",
    "monitor_watchdog", "runtime_check", "privacy", "manage_db",
)
OLD = "|".join(OLD_MODULES)
ROOT_SHIMS = {"main.py", "manage_db.py", "operations.py", "privacy.py", "runtime_check.py",
              "monitor_watchdog.py", "webapp.py", "schema_migrations.py"}


def _sources():
    for folder in ("nuligahelper", "tests"):
        for path in sorted((PROJECT / folder).rglob("*")):
            if path.suffix in {".py", ".mjs"} and "__pycache__" not in path.parts:
                yield path


def test_code_and_tests_import_package_modules():
    old_import = re.compile(rf"^\s*(?:import ({OLD})\b(?!\.)|from ({OLD})(?:\.\w+)* import )", re.M)
    offenders = [f"{path.relative_to(PROJECT)}: {match.group(0).strip()}"
                 for path in _sources() for match in old_import.finditer(path.read_text())]
    assert not offenders, "use nuligahelper.* imports: " + "; ".join(offenders)


def test_patch_targets_name_package_modules():
    stale = re.compile(rf"""patch(?:\.dict)?\(\s*['"]({OLD})[.'"]""")
    offenders = [f"{path.relative_to(PROJECT)}: {match.group(0)}"
                 for path in _sources() if path.name != Path(__file__).name
                 for match in stale.finditer(path.read_text())]
    assert not offenders, "patch nuligahelper.<module>.<name>: " + "; ".join(offenders)


def test_repository_root_keeps_only_transition_shims():
    assert {path.name for path in PROJECT.glob("*.py")} == ROOT_SHIMS, (
        "application code belongs in nuligahelper/; the root only keeps the transition shims"
    )
    assert sorted(path.name for path in (PROJECT / "test").iterdir()) == ["run_tests.sh"], (
        "test/ only keeps the deploy-command shim; tests live in tests/"
    )


if __name__ == "__main__":
    h.run_all(dict(globals()))
