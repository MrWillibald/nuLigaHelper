#!/bin/bash
set -e

if [ -z "${NULIGAHELPER_SECRET:-}" ] && [ -f .nuligahelper_secret ]; then
    export NULIGAHELPER_SECRET="$(<.nuligahelper_secret)"
fi
: "${NULIGAHELPER_SECRET:?NULIGAHELPER_SECRET muss gesetzt sein}"

# Activate the virtual environment
source venv/bin/activate

# Local/development server only (trusted LAN, port 8080). For production use the
# documented Caddy -> loopback Gunicorn -> systemd deployment in deploy/PRODUCTION.md.
echo "Lokaler Entwicklungsserver; nicht öffentlich ins Internet stellen." >&2
python webapp.py
