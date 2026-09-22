"""Independent root cron watchdog. Sends only through the fixed alert unit."""
import subprocess
import production as p


def main():
    try:
        healthy = p.success_age('monitor', '/var/lib/nuligahelper') <= 1800
    except (OSError, ValueError):
        healthy = False
    if not healthy:
        result = subprocess.run(['systemctl', 'start', 'nuligahelper-alert@monitor.service'],
                                capture_output=True, timeout=45, check=False)
        return 0 if result.returncode == 0 else 2
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
