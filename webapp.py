# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Transition shim: the installed web unit still starts Gunicorn with
# "webapp:app". New units use "nuligahelper.web:app".
# Remove after every host runs the updated units (docs/betrieb/deployment.md).
# ---------------------------------------------------------------

from nuligahelper.web import app  # noqa: F401
