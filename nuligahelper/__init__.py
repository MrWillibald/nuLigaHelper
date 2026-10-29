# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Application package:
#   db/     domain layer and database (models, rules, migrations)
#   web/    Flask web interface (templates and static files inside)
#   jobs/   daily job: scraping, sync, notifications and wording
#   ops/    production operations: backup, monitoring, privacy
#   common  configuration loading and the effective date
#   cli     operator command line (python -m nuligahelper.cli)
# ---------------------------------------------------------------

from pathlib import Path

# Repository root: config.json, alembic.ini and relative database paths
# resolve here, independent of the module that asks.
PROJECT_DIR = Path(__file__).resolve().parent.parent
