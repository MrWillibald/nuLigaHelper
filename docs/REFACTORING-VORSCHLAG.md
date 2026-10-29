# Vorschlag: Refaktorisierung und neue Ordnerstruktur

Stand: 29.10.2026, Basis `master` @ `dbaff44`. Dies ist ein **Vorschlag**, noch keine
Umsetzung. Nichts davon ändert Fachregeln aus `AGENTS.md`.

## 1. Ausgangslage in Zahlen

| Bereich | Heute |
|---|---|
| Python-Dateien im Wurzelordner | **25** Module + 2 Shell-Skripte + 4 Requirements/Config-Dateien |
| `webapp.py` | 2 789 Zeilen, davon ~2 550 Zeilen **innerhalb von `create_app()`** (alle Routen sind verschachtelte Funktionen) |
| `db.py` | 2 434 Zeilen: Modelle, SQLite-Setup, Sync, Personen, Dienste, Tagesblöcke, Sperren, Audit |
| `static/app.js` | 1 205 Zeilen in einer Datei |
| `README.MD` | 758 Zeilen / 64 KB, Deutsch und Englisch gemischt, überwiegend Betriebsdoku |
| `test/` | 80 Dateien flach in einem Ordner (Python + `.mjs`) |
| Doku-Dateien im Wurzelordner | `README.MD`, `AGENTS.md`, `CLAUDE.md`, `DEPLOYMENT.md`, `DATA-INVENTORY.md` |
| Testlauf (`test/run_tests.sh`) | ~3 min, 486 PASS, **1 FAIL** (siehe 5.4, umgebungsabhängig) |

Was heute **gut** ist und bleiben soll: Fachlogik konsequent in `db.py` statt in den
Aufrufern, Offline-Tests mit synthetischen Daten, zentraler Textkatalog
`messages.py`, explizite Alembic-Migrationen, dokumentierte Fallstricke in `AGENTS.md`.

Die Hauptprobleme:

1. **Kein Einstiegspunkt erkennbar.** 25 gleichrangige Dateien im Wurzelordner; man
   sieht nicht, was Web, täglicher Job, Betrieb oder einmalige Migration ist.
2. **Zwei Riesen-Dateien.** `webapp.py` und `db.py` sind je ~2 500 Zeilen. In
   `create_app()` teilen sich alle Routen Closures (`get_session`, `tier_for`, …) –
   einzelne Seiten lassen sich weder isoliert lesen noch testen.
3. **Doku ist verstreut und teils veraltet** (Details in 5.1).
4. **Kleinere Altlasten** (Duplikate, toter Code, Debug-Schalter im Code).

## 2. Ziel-Struktur

Leitidee: **Ein Ordner pro Aufgabe**, Namen sagen, wofür etwas da ist. Der Wurzelordner
enthält nur noch Startskripte, Abhängigkeiten und die wichtigsten Doku-Dateien.

```
nuLigaHelper/
├── README.md                    Kurz: Was ist das? Schnellstart. Links in docs/
├── AGENTS.md  CLAUDE.md         Regeln für Mitwirkende/Agenten (bleiben im Root)
├── config_template.json
├── requirements.txt             (siehe 4.6)
├── requirements-production.txt
├── requirements-test.txt
├── alembic.ini                  zeigt auf nuligahelper/db/migrations
├── run.sh                       täglicher Job        → python -m nuligahelper.jobs.daily
├── run_webapp.sh                lokaler Webserver    → python -m nuligahelper.web
├── manage.sh (optional)         CLI                  → python -m nuligahelper.cli
│
├── nuligahelper/                ◀ gesamter Anwendungscode
│   ├── config.py                ← common.py (Konfig laden, effective_today, Saison)
│   ├── cli.py                   ← manage_db.py
│   │
│   ├── db/                      ◀ Fachlogik + Datenbank  (heute: db.py)
│   │   ├── __init__.py          öffentliche API, re-exportiert alles → "from nuligahelper import db" bleibt
│   │   ├── models.py            ORM-Klassen (Team, Person, Game, Assignment, DayBlock, Audit, …)
│   │   ├── engine.py            SQLite-Engine, PRAGMAs, initialize_db/verify_db
│   │   ├── roles.py             Rollen, Slot-Anzahl, Labels, Aufgabenbeschreibungen ← task_labels.py
│   │   ├── clock.py             duty_now, Cutoff-Berechnung (09:00 Vortag)
│   │   ├── sync.py              sync_games + SyncEvents
│   │   ├── people.py            Personen, Teams, Mitgliedschaft, Registrierung, MV
│   │   ├── duties.py            Spieldienste: claim/release/replacement/takeover/lock
│   │   ├── day_blocks.py        Vorbereitung/Kuchen/Abbau
│   │   ├── staffing.py          missing_slots, staffing_status, claim_eligibility
│   │   ├── visibility.py        Spielplan-Sichtbarkeit
│   │   ├── audit.py             Audit-Snapshots
│   │   ├── age_eligibility.py   (unverändert verschoben)
│   │   ├── contacts.py          ← contact_validation.py
│   │   ├── schema_upgrade.py    ← schema_migrations.py
│   │   └── migrations/          ← migrations/ (Alembic: env.py, versions/)
│   │
│   ├── web/                     ◀ Weboberfläche  (heute: webapp.py)
│   │   ├── __init__.py          create_app(), app  (gunicorn: nuligahelper.web:app)
│   │   ├── __main__.py          lokaler Entwicklungsserver
│   │   ├── security.py          Proxy-Grenze, Host-Check, Header, CSRF, Zugriffsstufen
│   │   ├── auth.py              Login, Registrierung, Codes/Challenges
│   │   ├── auth_abuse.py        (unverändert verschoben)
│   │   ├── schedule.py          Spielplan "/", Sichtbarkeit, Kandidaten-API
│   │   ├── schedule_view.py     build_schedule()/block_view() als reine Funktionen
│   │   ├── assignments.py       /api/assignment/*, /api/block-assignment/*, /dienst, Kuchen, Team
│   │   ├── persons.py           /personen*, /api/teams/<id>/mv
│   │   ├── statistics.py        /statistik, /audit
│   │   ├── formatting.py        ak_color, parse_date, display_time, MONATE, SLOT_LABELS
│   │   ├── templates/
│   │   └── static/
│   │
│   ├── jobs/                    ◀ täglicher Cron-Job
│   │   ├── daily.py             ← main.py
│   │   ├── daily_lock.py
│   │   ├── scraper.py
│   │   ├── notifier.py
│   │   ├── messages.py
│   │   └── message_config.py
│   │
│   └── ops/                     ◀ Produktionsbetrieb
│       ├── backup.py            Snapshot, Dropbox, Restore
│       ├── settings.py          ← production.py (Validierung der Betriebs-Konfig)
│       ├── journal_logging.py   ← production_logging.py
│       ├── monitoring.py        ← operations.py
│       ├── watchdog.py          ← monitor_watchdog.py
│       ├── runtime_check.py
│       └── privacy.py
│
├── tests/                       ◀ gleiche Gliederung wie der Code
│   ├── helpers.py  fixtures/
│   ├── db/   web/   jobs/   ops/   migrations/
│   └── js/                      alle *.mjs (Node/Playwright)
│
├── release-assets/              bleibt (Pfad ist in Host-Installation und Deploy-Skript verdrahtet)
│
├── docs/
│   ├── README.md                Inhaltsverzeichnis
│   ├── benutzer/                BENUTZERHANDBUCH.md, MITGLIEDERHANDBUCH.md, screenshots/, pdf/ (← output/pdf)
│   ├── betrieb/                 deployment.md, datenschutz-inventar.md, datenbank-migration.md, backup.md
│   ├── entwicklung/             architektur.md, tests.md
│   └── build_handbook.py
│
└── openspec/                    unverändert
```

**Ergebnis:** Wurzelordner von ~45 auf ~15 Einträge; jede Python-Datei liegt in genau
einem von vier sprechenden Ordnern (`db`, `web`, `jobs`, `ops`).

### Bewusste Entscheidungen

- **Paketname `nuligahelper/`** statt `src/`-Layout: kein Packaging/Installieren nötig,
  `python -m nuligahelper…` läuft direkt aus dem Checkout.
- **`db` bleibt als Name** (statt `domain`): alle Aufrufer und ~80 Testdateien sprechen
  heute `db.claim_slot(...)` usw. – mit `db/__init__.py` als Fassade bleibt das so.
- **`release-assets/` wird nicht umbenannt**: Pfade wie
  `/opt/nuligahelper/current/release-assets/recovery_check.py` stehen im Deploy-Skript,
  in installierten systemd-Units und in Tests; der Gewinn wäre rein kosmetisch.
  `deploy/` ist außerdem schon als privater, gitignorierter Ordner belegt.
- **Keine neuen Frameworks**, keine neue Abstraktionsebene (kein Repository-/Service-
  Layer). Es wird nur aufgeteilt und gruppiert.

## 3. Umsetzung in Phasen

Jede Phase ist ein eigener PR, verhaltensneutral, Testsuite grün. Phasen 1–2 ändern
**keine** Importpfade nach außen und sind damit ohne Deploy-Risiko.

### Phase 0 – Doku-Hygiene (klein, sofort möglich)
- Veraltete Aussagen korrigieren (Liste in 5.1).
- `AGENTS.md`-Modultabelle um die 16 fehlenden Module ergänzen.
- `README.MD` → `README.md`: kurzer Überblick + Schnellstart; Betriebs-, Migrations-
  und Backup-Abschnitte nach `docs/betrieb/` verschieben (Inhalt unverändert).

### Phase 1 – `webapp.py` in Themen-Module zerlegen
> **Umgesetzt** mit `split-webapp-into-modules`: Paket `webapp/` mit `context`, `security`,
> `auth`, `challenges`, `schedule`, `views`, `assignments`, `persons`, `statistics` und
> `formatting`; `create_app()` hat noch rund 70 Zeilen, kein Modul mehr als 500.

- `create_app()` behält nur: Konfiguration, Session-Handling, Registrierung der
  Themen-Module. Gemeinsame Closures (`get_session`, `tier_for`, `current_mv_team_ids`,
  `api_error`, …) werden Modulfunktionen, die über `flask.g` / `current_app` arbeiten.
- Routen wandern in Module (`auth`, `schedule`, `assignments`, `persons`,
  `statistics`), die je eine Funktion `register(app)` haben. Bewusst **keine Flask-
  Blueprints**: Blueprints setzen immer ein Präfix vor die Endpunkt-Namen
  (`schedule.schedule`), dann müssten alle ~40 `url_for()`-Aufrufe in Templates und
  Tests angepasst werden. So bleiben die Endpunkt-Namen gleich.
- `build_schedule()` (~190 Zeilen) und die Statistik-Berechnung (~125 Zeilen) werden
  reine Funktionen `(session, viewer, …) -> dict` → direkt testbar ohne HTTP.
- Technisch wird `webapp.py` zum Paket `webapp/` (gleicher Importname, `create_app`
  und `app` in `webapp/__init__.py`). Damit laufen `gunicorn webapp:app`, die systemd-
  Units und alle Tests unverändert; die Umbenennung zu `nuligahelper/web/` folgt erst
  in Phase 3.

### Phase 2 – `db.py` in Teilmodule zerlegen
> **Umgesetzt** mit `split-domain-module`: Paket `db/` mit 16 Modulen in fester Schichtung
> (`roles` … `sync`); `import db` und alle `db.<name>` bleiben gültig. Tests patchen
> `db.clock.duty_now` und `db.storage.SQLITE_TIMEOUT_SECONDS`.

- Aufteilung wie in Abschnitt 2, zunächst als Paket `db/` im Wurzelordner;
  `db/__init__.py` re-exportiert die heutige öffentliche API vollständig (`import db`
  und `migrations/env.py` funktionieren weiter).
- **Stolperstein Monkeypatching:** Tests patchen `db.duty_now` (17×) und
  `common.effective_today` (52×). Nach einer Aufteilung greift ein Patch auf die
  Fassade nicht mehr, wenn Teilmodule `duty_now` direkt importiert haben. Regel daher:
  Teilmodule rufen `clock.duty_now()` / `config.effective_today()` **immer über das
  Modul** auf (nie `from … import duty_now`), und die Patch-Ziele in den Tests werden
  in derselben Änderung auf `…clock.duty_now` umgestellt.
- Reihenfolge der Extraktion (vom Rand zur Mitte): `roles` → `engine` → `models` →
  `visibility` → `sync` → `people` → `clock`/`audit` → `day_blocks` → `duties` → `staffing`.

### Phase 3 – Paket `nuligahelper/` + Tests gruppieren
- Module in die Zielordner verschieben (`git mv`, damit die Historie erhalten bleibt),
  Importe umstellen, `alembic.ini` `script_location` anpassen.
- `tests/` nach Bereichen gliedern; `run_tests.sh` sucht rekursiv
  (`tests/**/test_*.py`), `helpers.py` setzt `sys.path` weiterhin auf das Projekt.
- **Deploy-Übergang** (einziger Teil mit Betriebsrisiko):
  - systemd-Units in `release-assets/systemd/` auf `python -m nuligahelper.…` und
    `gunicorn … nuligahelper.web:app` umstellen; Watchdog-Cron ebenso.
  - Für **ein Release** bleiben im Wurzelordner dünne Weiterleitungen (`main.py`,
    `webapp.py`, `operations.py`, `privacy.py`, `runtime_check.py`,
    `monitor_watchdog.py`, `manage_db.py`, `schema_migrations.py`), damit bereits
    installierte Units und das installierte Deploy-Skript (es importiert
    `schema_migrations` im neuen Release) weiter funktionieren, bis die neuen Units
    mit `DEPLOYMENT.md` installiert sind. Danach entfernen.
  - `test_deployment_assets.py` / `test_release_prepare.py` prüfen diese Pfade und
    werden mit angepasst.

### Phase 4 – Doku-Ordner und Frontend (optional)
- `docs/` wie oben gliedern; `output/pdf` → `docs/benutzer/pdf` (Pfad in
  `build_handbook.py` und README-Links anpassen).
- `static/app.js` in thematische Dateien teilen (Feedback/Task-Hilfe,
  Dienst-Auswahl & Bestätigung, Kandidaten-Laden, Tagesblöcke, Personen-Formular).
  Ohne Build-Schritt per `<script type="module">`. Nur sinnvoll, wenn die
  `js_*.mjs`-Tests vorher angeschaut werden, da sie `app.js` direkt laden.

## 4. Inhaltliche Refaktorisierungen (unabhängig von der Struktur)

| # | Thema | Vorschlag | Aufwand |
|---|---|---|---|
| 4.1 | **Debug-Schalter im Code** (`DEBUG_FLAG`, `CHANGE_DAY`, `DEBUG_TODAY` in `common.py`, müssen vor jedem Commit zurückgesetzt werden) | **Umgesetzt** (`move-debug-switches-to-environment`): `NULIGAHELPER_TODAY` und `NULIGAHELPER_DRY_RUN`, in Produktion abgewiesen | – |
| 4.2 | **Doppelte SQLite-Hilfsfunktionen**: `_backup_path`/`_create_backup` in `schema_migrations.py` **und** `game_identity_migration.py`; `_readonly_uri` in `backup.py` **und** `schema_migrations.py` | **Umgesetzt** (`tidy-maintenance-helpers`): `backup.readonly_uri()` ersetzt beide Kopien; die Backup-Duplikate entfielen mit 4.4 | – |
| 4.3 | **Ruhende Zeitungsartikel-Funktion**: `Notifier.send_article`, auskommentierter Aufruf in `main.py` (mit altem `strGameDateSa`-Stil), 5 `dormant`-Texte in `messages.py` | **Entschieden: bleibt vorerst deaktiviert erhalten**, wird nicht angefasst | – |
| 4.4 | **Einmal-Migration `game_identity_migration.py`** (Legacy `games.source_key`) | **Umgesetzt** (`remove-legacy-game-identity-migration`): entfernt, Altschema wird weiterhin erkannt und abgewiesen | – |
| 4.5 | **Unbenutzte/irreführende Konstanten**: `ROLE_MV` (MV ist keine Dienstrolle mehr), `ROLE_SUPPORT = ROLE_CASH` (Alias aus der Kasse/Reinigung-Umbenennung) | **Umgesetzt** (`tidy-maintenance-helpers`): beide Konstanten entfernt | – |
| 4.6 | **Drei Requirements-Dateien mit unterschiedlichen Versionen**: `requirements.txt` ungepinnt, Produktion/CI gepinnt | **Umgesetzt** (`tidy-maintenance-helpers`): `requirements.txt` = `-r requirements-test.txt` | – |
| 4.7 | **Zwei Test-Runner** (eigener Mini-Runner in `helpers.py` + pytest) | Mittelfristig nur pytest (CI ruft dann `pytest`), Mini-Runner entfernen. Spart Code, liefert bessere Fehlermeldungen | mittel |
| 4.8 | **Benennung**: `production.py` wird als `p` bzw. `production_data` importiert, `common.py` sagt nichts über den Inhalt | `ops/settings.py`, `config.py` (siehe Zielstruktur) | in Phase 3 |
| 4.9 | **Veralteter Kopfkommentar** in `manage_db.py` („until the web interface is available“) | Aktualisieren | trivial |

## 5. Gefundene Unstimmigkeiten

### 5.1 Veraltete Doku (Phase 0)
- `CLAUDE.md`: „No migration layer by intent: after a schema change delete the `.db`
  file“ – **falsch**, es gibt Alembic (`migrations/`, `schema_migrations.py`,
  `manage_db.py migrate-schema`).
- `CLAUDE.md`: Routen `/api/assignment`, `/api/game/<id>/team` – tatsächlich
  `/api/assignment/claim|release|…` und `/api/games/<id>/team`; außerdem fehlen
  `/dienst`, `/audit`, Login/Registrierung, Blöcke.
- `CLAUDE.md`: „Pages: schedule, statistics, persons“ – es gibt zusätzlich Login,
  Registrierung, Dienst-Formular, Audit, Rechtliches.
- `AGENTS.md`-Modultabelle nennt 9 Module, es gibt 25 (fehlen u. a. `backup`,
  `auth_abuse`, `privacy`, `operations`, `production`, `schema_migrations`,
  `age_eligibility`, `contact_validation`, `daily_lock`).
- `AGENTS.md`: „local network, port 8080 … Production serving, TLS and public exposure
  remain separate work“ – inzwischen gibt es das Caddy/Gunicorn/systemd-Deployment.
- `run_webapp.sh` verweist auf `deploy/PRODUCTION.md` – dieser Ordner ist gitignoriert;
  gemeint ist vermutlich `DEPLOYMENT.md` / `release-assets/README.md`.

### 5.2 Sprachmischung
README und Doku wechseln zwischen Deutsch (Benutzer) und Englisch (Betrieb).
Vorschlag: Benutzer-Doku Deutsch, Entwickler-/Betriebsdoku Englisch – aber je Datei
**eine** Sprache.

### 5.3 Reihenfolgeabhängige Webapp-Tests
`AGENTS.md` dokumentiert, dass Webapp-Tests aufeinander aufbauen. Beim Aufteilen der
Testdateien (Phase 3) ist das ein Risiko → beim Verschieben jeweils einzeln laufen
lassen; langfristig eigene Fixtures pro Test.

### 5.4 Umgebungsabhängiger Test (behoben in `tidy-maintenance-helpers`)
`test_operations.py::test_monitor_detects_staleness_missing_timers_clock_disk_and_tls`
schlug im Basislauf fehl, weil er echtes `os.statvfs()` des Temp-Ordners nutzt und
die Plattenbelegung der Testmaschine mitbewertet (`app_disk: 1`). Der Test sollte ein
festes `statvfs`-Ergebnis vorgeben.

### 5.5 Überflüssige bzw. doppelte Markdown-Inhalte (Prüfung 29.10.2026)
Keine der 11 Markdown-Dateien außerhalb von `openspec/` ist verwaist; alle werden verlinkt.
Doppelt gepflegt sind aber:
- `CLAUDE.md` wiederholt Befehle und Regeln aus `AGENTS.md` (und ist dabei veraltet) →
  auf kurze Orientierung plus `@AGENTS.md` reduzieren.
- `README.MD` enthält ~360 Zeilen englische Verhaltensbeschreibung (Weboberfläche,
  Sichtbarkeit, Aufgaben, Dienstpositionen, verbindliche Dienste, Altersgrenzen,
  Kuchen), die schon in den deutschen Handbüchern und in `openspec/specs/` stehen →
  durch kurze Funktionsliste mit Links ersetzen.
- `test/README.md` listet in seiner Tabelle nur 24 von 58 Testdateien → Tabelle
  durch Beschreibung der Testgruppen ersetzen.
- `docs/REFACTORING-VORSCHLAG.md` (diese Datei) überschneidet sich mit den
  OpenSpec-Changes → nach Abschluss von `restructure-project-layout` löschen.

Bewusst behalten: `DEPLOYMENT.md` (Routine-Update) und `release-assets/README.md`
(Ersteinrichtung, von `test_deployment_assets.py` geprüft) haben verschiedene Zwecke;
`DATA-INVENTORY.md` ist das Datenschutz-Inventar; `openspec/changes/archive/` ist die
Änderungshistorie. Umsetzung im Change `refresh-developer-documentation`.

## 6. Entscheidungen und Umsetzung über OpenSpec

Entscheidungen vom 29.10.2026:
- `game_identity_migration.py` darf weg – alle Datenbanken sind migriert.
- Die Zeitungsartikel-Funktion bleibt vorerst deaktiviert erhalten.
- Die Umsetzung läuft über OpenSpec-Changes in `openspec/changes/`.

Changes in empfohlener Reihenfolge (je ein PR):

| # | Change | Inhalt | Spec-Deltas | Stand |
|---|---|---|---|---|
| 1 | `refresh-developer-documentation` | Phase 0, 4.9 | keine (nur Doku) | umgesetzt |
| 2 | `remove-legacy-game-identity-migration` | 4.4 | `game-identity` (REMOVED), `schema-migrations` (MODIFIED) | umgesetzt |
| 3 | `move-debug-switches-to-environment` | 4.1 | `production-operations` (ADDED) | umgesetzt |
| 4 | `tidy-maintenance-helpers` | 4.2, 4.5, 4.6, 5.4 | keine (Refactoring) | umgesetzt |
| 5 | `split-webapp-into-modules` | Phase 1 | keine (Refactoring) | umgesetzt |
| 6 | `split-domain-module` | Phase 2 | keine (Refactoring) | umgesetzt |
| 7 | `restructure-project-layout` | Phase 3 + Doku-Teil aus Phase 4 | keine (Refactoring) | offen |

1–4 sind klein und inhaltlich unabhängig (sie berühren teils dieselben Doku-Zeilen); 7 setzt 2, 5 und 6 voraus und braucht ein
geplantes Wartungsfenster für die systemd-Units. Später und optional: `app.js`
aufteilen (Phase 4), nur noch pytest als Test-Runner (4.7).
