#!/bin/bash

# Run the full nuLigaHelper test suite (standalone runner)
cd "$(dirname "$0")/.."

if [ ! -d venv ]; then
    echo "Virtual environment not found. Create it first:" >&2
    echo "  python3 -m venv venv && ./venv/bin/pip install -r requirements.txt" >&2
    exit 1
fi

source venv/bin/activate

status=0
for test_file in test/test_*.py; do
    echo "=== $test_file ==="
    python "$test_file" || status=1
done

deactivate
exit $status
