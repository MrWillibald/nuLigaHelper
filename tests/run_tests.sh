#!/bin/bash

# Run the full nuLigaHelper test suite (standalone runner). Every test module
# runs as part of the tests package from the project root:
#   python -m tests.db.test_db
cd "$(dirname "$0")/.."

if [ ! -d venv ]; then
    echo "Virtual environment not found. Create it first:" >&2
    echo "  python3 -m venv venv && ./venv/bin/pip install -r requirements.txt" >&2
    exit 1
fi

source venv/bin/activate

status=0
for test_file in tests/*/test_*.py; do
    module="${test_file%.py}"
    module="${module//\//.}"
    echo "=== $test_file ==="
    python -m "$module" || status=1
done

deactivate
exit $status
