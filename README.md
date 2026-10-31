# Personalplanung (Master-File)

Baukasten für ein gemeinsames Excel-File zur Arbeitsplanung (Office 2016, ohne Makros).
Die Blätter lehnen sich an die bestehenden Dateien an.

| Blatt | Für wen | Inhalt |
|---|---|---|
| Planung | Personalchefin | Eine Zeile pro Person mit Hauptbereich. Pro Woche aus der Liste: EIC, JUC, FEL, FRE, TOM, BLU, LOG, F (Ferien), K (Krank), U (Unfall) oder zwei Bereiche wie `BLU/FRE` (zählt je halb). |
| Übersicht | Personalchefin | Pro Bereich IST, SOLL und IST - SOLL, getrennt nach Frauen und Männern. Dazu „Wen anstellen?“. |
| Grafik | Personalchefin | Bereich wählen: Frauen, Männer und Total – Säulen = eingeplant, Linie = Bedarf. |
| Kräuter Eichhof | Kräuter-Leitung | Neue Leute ohne Gruppe, Gruppen mit Namen (automatisch), Bedarf / Ist / Differenz je Gruppe, Total und „Gruppen zuteilen“ (Name und Gruppe per Auswahlliste). |
| Gemüse Freiland | Gemüse-Leitung | Stundenplanung wie bisher. Darunter Soll, Ist und Differenz (ohne Namen) und eine Grafik. |
| Anleitung, Listen | alle / Admin | Kurze Anleitung, Jahr und Auswahllisten. |

## File erzeugen

```bash
python3 tools/build_personalplanung.py \
  --gesamt   "Arbeitsplanung gesamtbetrieblich _2026.xlsx" \
  --kraeuter "Wochenplan Kräuterteam Eichhof-Luchs 2026 Kopie offline.xlsx" \
  --freiland "Arbeitsplanung_Freiland_2026.xlsx" \
  --out      Personalplanung_2026_Test.xlsx
```

Die Quelldateien und das erzeugte File enthalten Personendaten und werden nicht ins Repository
eingecheckt (siehe `.gitignore`). Benötigt: Python 3 mit `openpyxl`.

## Web-Prototyp

`tools/web/personalplanung.html` ist ein klickbarer Prototyp der Planung als Webseite (Excel-Optik,
Drag & Drop, Rückgängig, Rollen Personalchefin / Kräuter Eichhof / Gemüse Freiland).
`tools/build_web_demo.py` füllt ihn mit der Planungsstruktur aus den Excel-Dateien und ersetzt
alle Namen durch erfundene Namen:

```bash
python3 tools/build_web_demo.py --gesamt G.xlsx --kraeuter K.xlsx --freiland F.xlsx \
  --vorlage tools/web/personalplanung.html --out personalplanung.html
```

Der Prototyp speichert Änderungen nur im eigenen Browser.

## Arbeitsversion (HTML + Daten-Datei auf dem Netzlaufwerk)

`tools/web/personalplanung_app.html` ist das fertige Programm. Es braucht keinen Server und kein
Internet: Programm und Daten liegen als zwei Dateien im selben Ordner auf dem Netzlaufwerk.

| Datei | Inhalt |
|---|---|
| `personalplanung.html` | das Programm (Kopie von `tools/web/personalplanung_app.html`) |
| `personalplanung-daten.json` | alle Daten: Personen, Wochenplan 2026/2027, Bedarf, Stunden Freiland, Mutationen, Notizen, Traktanden, Aufgaben, Verlauf |
| `Sicherungen/` | legt das Programm selbst an: eine Kopie der Daten pro Tag |

- Öffnen mit Microsoft Edge oder Google Chrome, Namen eintragen, Ordner wählen. Beim nächsten Mal
  genügt „Weiter mit Ordner“.
- Jede Änderung wird sofort in die Daten-Datei geschrieben, mit Name, Rolle und Zeit im Verlauf.
- Speichern zwei Personen gleichzeitig, werden die Änderungen zusammengeführt (Versionsnummer in der
  Datei). Änderungen der anderen lädt das Programm alle 20 Sekunden nach.
- Andere Browser (z.B. Firefox) können die Daten nur als Arbeitskopie öffnen und herunterladen.

Die Daten-Datei mit den echten Namen erzeugt `tools/build_daten.py` aus den bisherigen Excel-Dateien:

```bash
python3 tools/build_daten.py --gesamt G.xlsx --kraeuter K.xlsx --freiland F.xlsx \
  --out personalplanung-daten.json
```

Die Daten-Datei enthält Personendaten und gehört nicht ins Repository (siehe `.gitignore`).
