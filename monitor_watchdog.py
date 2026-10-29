# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Transition shim: already installed systemd units and cron entries
# still call this path. New units run "python -m nuligahelper.ops.watchdog".
# Remove after every host runs the updated units (docs/betrieb/deployment.md).
# ---------------------------------------------------------------

import runpy

if __name__ == "__main__":
    runpy.run_module("nuligahelper.ops.watchdog", run_name="__main__", alter_sys=True)
