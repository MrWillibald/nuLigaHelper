# ---------------------------------------------------------------
#                          nuLigaHelper
# ---------------------------------------------------------------
# Transition shim: already installed systemd units and cron entries
# still call this path. New units run "python -m nuligahelper.jobs.daily".
# Remove after every host runs the updated units (DEPLOYMENT.md).
# ---------------------------------------------------------------

import runpy

if __name__ == "__main__":
    runpy.run_module("nuligahelper.jobs.daily", run_name="__main__", alter_sys=True)
