#!/bin/bash

# Activate the virtual environment
source venv/bin/activate

# Run the nuLigaHelper daily job
python main.py

# Deactivate the virtual environment
deactivate
