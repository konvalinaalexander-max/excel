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
