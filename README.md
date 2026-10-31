# Personalplanung (Master-File)

Baukasten für ein gemeinsames Excel-File zur Arbeitsplanung (Office 2016, ohne Makros).
Die Blätter lehnen sich an die bestehenden Dateien an.

| Blatt | Für wen | Inhalt |
|---|---|---|
| Planung | Personalchefin | Eine Zeile pro Person, pro Woche ein Kürzel: EIC, JUC, FEL, FRE, TOM, BLU, LOG, F (Ferien), K (Krank), U (Unfall). Zwei Bereiche in einer Woche: `BLU/EIC` (zählt je halb). |
| Übersicht | Personalchefin | Pro Bereich IST, SOLL und IST - SOLL, getrennt nach Frauen und Männern. Dazu „Wen anstellen?“ und Grafiken. |
| Kräuter Eichhof | Kräuter-Leitung | Namen erscheinen automatisch nach Gruppe (Steckling, Fax 10.5er, Fax 13er, Verteilen / Topfmaschine, Luchs). Pro Gruppe: Bedarf (Eingabe), Ist, Differenz. |
| Gemüse Freiland | Gemüse-Leitung | Stundenplanung wie bisher. Darunter Soll, Ist und Differenz, ohne Namen. |
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
