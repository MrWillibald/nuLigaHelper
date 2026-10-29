# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Transition shim: the installed deploy command probes the schema
# head with "import schema_migrations". The module now lives in
# nuligahelper.db.schema_upgrade. Remove together with the other
# root shims once hosts run the updated deploy command.
# ---------------------------------------------------------------

from nuligahelper.db.schema_upgrade import *  # noqa: F401,F403
from nuligahelper.db.schema_upgrade import head_revisions, inspect_schema  # noqa: F401
