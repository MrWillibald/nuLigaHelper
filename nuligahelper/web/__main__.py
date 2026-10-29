# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Local development server: ./venv/bin/python -m nuligahelper.web
# Production uses Gunicorn with nuligahelper.web:app (see release-assets/).
# ---------------------------------------------------------------

from nuligahelper.web import app

app.run(host="0.0.0.0", port=8080)
