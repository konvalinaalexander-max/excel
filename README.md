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

## Arbeitsversion: eine HTML-Datei mit den Daten

`tools/web/personalplanung_app.html` ist das Programm (Vorlage ohne Daten). Für den Betrieb wird daraus eine
einzige Datei `personalplanung.html` gebaut, in der Programm und Daten zusammen stehen. Sie liegt in einem
Ordner auf dem Netzlaufwerk; es braucht keinen Server und kein Internet.

- **Öffnen** (Doppelklick) zeigt sofort alle Daten, zum Ansehen in jedem Browser.
- **Speichern** geht mit Microsoft Edge oder Google Chrome: Beim ersten Ändern gibt man einmal den Ordner frei, in
  dem die Datei liegt. Danach schreibt jede Änderung die Datei neu (nur der Datenblock wird ersetzt).
- **Mehrere gleichzeitig:** Vor dem Speichern wird die Datei gelesen; hat jemand anderes gespeichert, werden die
  eigenen Änderungen darübergelegt. Nach dem Speichern wird kontrolliert, ob die eigene Version wirklich in der Datei
  steht (sonst wird nochmals zusammengeführt). Offene Programme laden fremde Änderungen alle 20 Sekunden nach.
- **Sicherung:** einmal pro Tag eine Kopie der ganzen Datei im Unterordner `Sicherungen/`.
- **Datei-Menü:** Kopie herunterladen, Daten übernehmen (aus `personalplanung-daten.json` oder einer älteren
  `personalplanung.html`), Programm aktualisieren (neue Version einspielen, die Daten bleiben), Kurzanleitung.
- Alle können alles ändern. Der Verlauf hält fest, wer was wann geändert hat.

Blätter: **Personal** (Wochenplan, Personen mit Verträgen, «Prüfen» mit offenen Datenfragen), **Übersicht**
(Überblick, pro Bereich, Bedarf eingeben), **Personalsitzung** (Wo fehlen Leute? · Mutationen · Anstellen · Notizen
und Protokoll), **Kräuter Eichhof** (Brett, Jahresübersicht, Aufgaben), **Gemüse Freiland** (Stunden, Bedarf, Team).

### Daten nur aus den Originalen

`tools/build_daten.py` liest die bisherigen Excel-Dateien und schreibt die Daten (JSON). Der Bedarf steht dort nur als
Gesamtzahl; «Bedarf Frauen/Männer» bleibt leer, bis er im Programm erfasst wird. 2027 nur, soweit im Original bewusst
eingetragen. Was beim Übernehmen bereinigt wurde, zeigt das Programm unter Personal → Prüfen.

```bash
python3 tools/build_daten.py --gesamt G.xlsx --kraeuter K.xlsx --freiland F.xlsx --out personalplanung-daten.json
python3 tools/build_html.py --vorlage tools/web/personalplanung_app.html --daten personalplanung-daten.json \
  --out personalplanung.html
```

Die Daten-Datei und die fertige `personalplanung.html` enthalten Personendaten und gehören nicht ins Repository
(siehe `.gitignore`).
