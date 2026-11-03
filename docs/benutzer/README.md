# Benutzerdokumentation

Es gibt zwei eigenständige deutschsprachige Ausgaben:

| Zielgruppe | Bearbeitbare Quelle | Druckbare Fassung |
| --- | --- | --- |
| Mitglieder | [`MITGLIEDERHANDBUCH.md`](MITGLIEDERHANDBUCH.md) | [`pdf/nuLigaHelper_Mitgliederhandbuch.pdf`](pdf/nuLigaHelper_Mitgliederhandbuch.pdf) |
| Alle Rollen | [`BENUTZERHANDBUCH.md`](BENUTZERHANDBUCH.md) | [`pdf/nuLigaHelper_Benutzerhandbuch.pdf`](pdf/nuLigaHelper_Benutzerhandbuch.pdf) |

Gemeinsame Abläufe bei Änderungen in beiden Quellen pflegen. Die Mitgliederfassung
enthält ausschließlich die Funktionen für Mitglieder; darin keine Beschreibungen
oder Screenshots mit zusätzlichen MV- oder Adminfunktionen ergänzen. Die
Kuchenabbildung `16-kuchen-mitglied.png` zeigt ausdrücklich die Mitgliederansicht.

Die markierten Screenshots unter `screenshots/` wurden mit der unveränderten
Weboberfläche von Version 0.30 und ausschließlich erfundenen Beispieldaten
erstellt. Die Beispiele sind auf den 25. Oktober 2026 bezogen. Orangefarbene
Nummern passen zu den Erläuterungen unter den jeweiligen Abbildungen. Für neue
Aufnahmen eine getrennte Testdatenbank verwenden und Mail- sowie SMS-Versand
abschalten. Keine Produktionsdaten, Anmeldecodes oder Zugangsdaten aufnehmen.

Die PDFs lassen sich mit Python 3 sowie den optionalen Dokumentationspaketen
`reportlab` und `Pillow` gemeinsam neu erzeugen:

```bash
python docs/benutzer/build_handbook.py
```

Mit `--edition members` beziehungsweise `--edition all` wird nur eine Ausgabe
erzeugt. Der Generator übernimmt die Farbwerte direkt aus `nuligahelper/web/static/style.css`:
Dunkelblau, orangefarbene Akzente, die siebenfarbige Kopfleiste, helle Flächen und
abgerundete Elemente entsprechen dadurch der Weboberfläche. Die Schriften
Liberation Sans aus `/usr/share/fonts/truetype/liberation` bieten einen zu Arial
passenden Schriftsatz. Alle Ressourcen sind lokal; der Generator liest keine
Anwendungskonfiguration und greift auf keine Datenbank oder Netzwerkadresse zu.

Diese Dokumentationswerkzeuge werden für den Betrieb des Webservice nicht
benötigt. Die erzeugten PDFs anschließend rendern und alle Seiten auf lesbare
Screenshots, passende Markierungen, vollständige Tabellen und gute Umbrüche
prüfen. Die Mitgliederfassung zusätzlich auf versehentlich enthaltene
Verwaltungsfunktionen kontrollieren. Inhaltsverzeichnisse, Sprunglinks auf dem
Deckblatt und PDF-Lesezeichen werden beim Erzeugen aktualisiert.

---
Zurück zur [Dokumentationsübersicht](../README.md)
