# nuLigaHelper

nuLigaHelper organizes the home games of the TuS Raubling handball department. A daily
job reads the home games from the BHV/nuLiga Hallenspielplan, stores them with all
helper duties in a SQLite database, sends reminders by e-mail or SMS and backs the
database up to Dropbox. A web interface lets members, team managers (MV) and admins
staff the duties themselves.

Host-specific deployment configuration and operator records are maintained privately
outside this repository.

# Features

- Reads all home games from the nuLiga Hallenspielplan and keeps track of time shifts,
  missing referees ("§77") and unknown new games.
- Game duties Zeitnehmer, Sekretär, Verkauf (2), Ordner, Kasse and Wischer: seven required
  positions for adult games, five for youth and unresolved categories, with age rules.
- Dated day blocks per home-game date: Vorbereitung and Aufräumen (three places each) and
  Kuchenlieferung (quantity and delivery time set by admins).
- Binding duties: helpers seek a replacement instead of withdrawing; occupied duties lock
  at 09:00 (Europe/Berlin) on the day before.
- Reminders one week and one day ahead, weekly vacancy invitations to the responsible
  team, MV follow-up for open positions, referee-coordinator alerts and admin info mails.
  E-mail is preferred, SMS (Twilio) is the fallback.
- Passwordless web interface (e-mail or SMS code) with access tiers derived on every request:
  - **Guest**: public read-only schedule with assigned helper names, no contact data.
  - **Member**: roster, statistics, own profile and own duties.
  - **MV**: additionally staffing and roster maintenance for the teams they manage.
  - **Admin**: all administration, registration approval and the audit trail.
- Statistics per team and person, open duties, append-only assignment audit.
- Validated online SQLite snapshots to Dropbox after every daily run, with dated retention.
- A newspaper-article notification exists but is currently disabled.

# Documentation

| Topic | Where |
| --- | --- |
| User manuals (German) | [Mitgliederhandbuch](docs/MITGLIEDERHANDBUCH.md) ([PDF](output/pdf/nuLigaHelper_Mitgliederhandbuch.pdf)) for members, [Benutzerhandbuch](docs/BENUTZERHANDBUCH.md) ([PDF](output/pdf/nuLigaHelper_Benutzerhandbuch.pdf)) for all roles; maintenance notes in [docs/README.md](docs/README.md) |
| Behavior contract | [`openspec/specs/`](openspec/specs/) – one specification per capability; planned changes in [`openspec/changes/`](openspec/changes/) |
| Operations | [docs/betrieb/](docs/betrieb/README.md): SQLite profile and backups, schema migration, management CLI, notification wording, private deployment |
| Releases to production | [DEPLOYMENT.md](DEPLOYMENT.md) (routine updates), [release-assets/README.md](release-assets/README.md) (host setup) |
| Data inventory | [DATA-INVENTORY.md](DATA-INVENTORY.md) |
| Development | [AGENTS.md](AGENTS.md) (rules and module map), [database schema](docs/entwicklung/database-schema.md), [tests](test/README.md), [refactoring proposal](docs/REFACTORING-VORSCHLAG.md) |

# Getting started

```bash
python3 -m venv venv && ./venv/bin/pip install -r requirements.txt
```

Copy `config_template.json` to `config.json` and adjust the club and recipient settings.
Notification wording lives in `messages.py`. E-mail, Twilio and Dropbox settings come from
environment variables mapped in `common.PROVIDER_ENV`; production refuses provider sections
in `config.json`.

`NULIGAHELPER_SECRET` is mandatory for the web app and the daily job. Set it in the
environment or store it in the gitignored `.nuligahelper_secret` file read by both launch
scripts. Keep it stable: changing it invalidates every session. `NULIGAHELPER_DB` optionally
overrides the SQLite path locally; production requires `NULIGAHELPER_ENV=production`, an
absolute `NULIGAHELPER_DB` and exact `NULIGAHELPER_TRUSTED_HOSTS` entries.

For local tests of date-dependent notifications, `NULIGAHELPER_TODAY=YYYY-MM-DD` pins the
effective date and `NULIGAHELPER_DRY_RUN=1` suppresses outbound mail and SMS:

```bash
NULIGAHELPER_TODAY=2025-11-21 NULIGAHELPER_DRY_RUN=1 ./run.sh
```

Production refuses to start while either variable is set.

```bash
python3 -c 'import secrets; print(secrets.token_urlsafe(48))' > .nuligahelper_secret
./venv/bin/python manage_db.py init           # create an empty database
./venv/bin/python manage_db.py grant-admin 1  # after adding the first person
./run.sh                                      # daily job: scrape, sync, notify, back up
./run_webapp.sh                               # local web interface on http://<ip>:8080
```

Games are inserted automatically by the first daily run. The daily job logs to standard
error (the systemd journal in production). Run it through one scheduler only. All CLI
commands are described in [docs/betrieb/cli.md](docs/betrieb/cli.md).

`run_webapp.sh` starts the Flask development server for local or trusted-LAN use only.
Public hosting uses the accepted production deployment (Caddy with TLS, Gunicorn and
systemd); see [private deployment](docs/betrieb/private-deployment.md) and
[DEPLOYMENT.md](DEPLOYMENT.md).

# Testing

The `test/` folder contains offline tests for the database, notifications, authentication,
authorization, CLI, concurrent claims, migrations, deployment assets and the web flow.
Every test uses synthetic secrets and throwaway databases:

```bash
test/run_tests.sh
```

Single files run with `./venv/bin/python test/test_db.py`, or use pytest
(`./venv/bin/python -m pytest test/ -v`). See [test/README.md](test/README.md) for details.
