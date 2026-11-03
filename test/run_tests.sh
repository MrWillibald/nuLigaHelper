#!/bin/bash

# Transition shim: the installed deploy command runs test/run_tests.sh while
# preparing a release. The suite now lives in tests/. Remove together with
# the root shims once hosts run the updated deploy command (docs/betrieb/deployment.md).
exec bash "$(dirname "$0")/../tests/run_tests.sh" "$@"
