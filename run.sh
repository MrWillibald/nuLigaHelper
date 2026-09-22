#!/bin/bash
set -e

if [ -z "${NULIGAHELPER_SECRET:-}" ] && [ -f .nuligahelper_secret ]; then
    export NULIGAHELPER_SECRET="$(<.nuligahelper_secret)"
fi
: "${NULIGAHELPER_SECRET:?NULIGAHELPER_SECRET muss gesetzt sein}"

# Activate the virtual environment
source venv/bin/activate

# Python owns the per-database overlap lock; exec preserves its status for cron.
exec python main.py
