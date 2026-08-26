#!/bin/bash

# Activate the virtual environment
source venv/bin/activate

# Start the web interface (reachable in local network on port 8080)
python webapp.py

# Deactivate the virtual environment
deactivate
