# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Local development server: ./venv/bin/python -m webapp
# Production uses Gunicorn with webapp:app (see release-assets/).
# ---------------------------------------------------------------

from webapp import app

app.run(host="0.0.0.0", port=8080)
