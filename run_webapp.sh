#!/bin/bash
set -e

if [ -z "${NULIGAHELPER_SECRET:-}" ] && [ -f .nuligahelper_secret ]; then
    export NULIGAHELPER_SECRET="$(<.nuligahelper_secret)"
fi
: "${NULIGAHELPER_SECRET:?NULIGAHELPER_SECRET muss gesetzt sein}"

# Activate the virtual environment
source venv/bin/activate

# Start the web interface (reachable in local network on port 8080)
python webapp.py
