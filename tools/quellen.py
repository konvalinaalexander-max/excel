"""Liest die bestehenden Planungsdateien (Gesamtplanung, Kräuter, Freiland) ein.

Enthält keine Personendaten: alle Namen kommen zur Laufzeit aus den Quelldateien.
"""
import difflib
import re
import unicodedata
import warnings

import openpyxl

warnings.filterwarnings("ignore")

NW = 53  # Anzahl Kalenderwochen im Planungsjahr (2026 hat 53 ISO-Wochen)

BEREICH_NAMEN = {
    "Kräuter Eichhof": "Kräuter Eichhof", "Kräuter Juch": "Kräuter Juch",
    "Kräuter Feld": "Kräuter Feld", "Gemüse Freiland": "Gemüse Freiland",
    "Tomaten": "Tomaten", "Gurken": "Tomaten", "Blumen": "Blumen", "Logistik": "Logistik",
}
CODE_MAP = {"DEF": "F", "PRO": "F", "SC": "F"}  # alte Abwesenheits-Kürzel -> Ferien

# Eichhof-Gruppen: Name, Geschlecht, Zeile "Bedarf" in der Kräuter-Datei, Zeilen der Personen
GRUPPEN = [
    ("Steckling", "F", 15, range(2, 15)),
    ("Fax 10.5er", "F", 28, range(16, 28)),
    ("Fax 13er", "F", 35, range(29, 35)),
    ("Verteilen / Topfmaschine", "M", 48, range(36, 48)),
    ("Luchs", "M", 55, range(49, 55)),
]
FUNKTION_KEYWORDS = [  # Reihenfolge = Priorität
    ("luchs", "Luchs"), ("13er", "Fax 13er"), ("10.5", "Fax 10.5er"), ("10,5", "Fax 10.5er"),
    ("steckling", "Steckling"), ("nützling", "Steckling"), ("topf", "Verteilen / Topfmaschine"),
    ("verteil", "Verteilen / Topfmaschine"), ("rampe", "Verteilen / Topfmaschine"),
    ("psm", "Verteilen / Topfmaschine"), ("fax", "Fax 10.5er"),
]


def norm(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z ]", " ", s).split()


def lese_gesamt(pfad):
    wb = openpyxl.load_workbook(pfad, data_only=True)
    ws = wb["Planung"]
    hdr = [c.value for c in ws[2]]
    ix = {h: i for i, h in enumerate(hdr)}
    personen = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row[ix["Vorname"]]:
            continue
        wochen = []
        for w in range(1, NW + 1):
            v = row[ix[f"26_{w:02d}"]]
            v = "" if v is None else str(v).strip().upper()
            wochen.append(CODE_MAP.get(v, v))
        bem = str(row[ix["Bemerkung"]] or "").strip()
        wechsel = str(row[ix["Wechselt?"]] or "").strip()
        if wechsel:
            bem = (bem + " | " if bem else "") + "Wechsel: " + wechsel
        personen.append({
            "badge": row[ix["Badge"]], "vorname": str(row[ix["Vorname"]]).strip(),
            "nachname": str(row[ix["Nachname"]] or "").strip(), "fm": row[ix["F/ M"]],
            "partner": row[ix["Partner"]], "fschein": row[ix["FSchein"]],
            "bereich": BEREICH_NAMEN.get(row[ix["Bereich, Haupt-"]], row[ix["Bereich, Haupt-"]]),
            "zustand": row[ix["Zustand"]], "eintritt": row[ix["Eintritt"]],
            "austritt": row[ix["Austritt"]], "funktion": str(row[ix["Funktion"]] or "").strip(),
            "bemerkung": bem, "wochen": wochen,
        })
    ub = wb["Übersicht"]

    def ub_row(r):
        return [ub.cell(r, 4 + w).value or 0 for w in range(1, NW + 1)]

    kz = wb["korrg. Zahlen"]
    juch_plan = {kz.cell(5, c).value: kz.cell(6, c).value for c in range(2, 28)}
    soll = {
        "VJ": ub_row(4), "KRANK": ub_row(16), "BLU": ub_row(22), "FEL": ub_row(47),
        "TOM": ub_row(53), "JUC_ERST": ub_row(68),
        "LOG": {lab: ub_row(r) for lab, r in [("Fixe Fahrer", 32), ("Fahrer Springer", 33),
                                              ("Rampe", 34), ("Stapler", 35), ("Werkstatt", 36)]},
    }
    juc = []
    for w in range(1, NW + 1):
        # «korrg. Zahlen» reicht bis KW 52; KW 53 kommt aus der Übersicht
        juc.append(juch_plan[w] if w >= 27 and juch_plan.get(w) is not None else soll["JUC_ERST"][w - 1])
    soll["JUC"] = juc
    return personen, soll


def lese_kraeuter(pfad):
    ws = openpyxl.load_workbook(pfad, data_only=True)["Gruppen Kostenstelle 2026"]
    bedarf, namen = {}, {}
    for name, _g, r_bedarf, r_pers in GRUPPEN:
        werte = [ws.cell(r_bedarf, 17 + min(w, 52)).value for w in range(1, NW + 1)]
        bedarf[name] = [v if isinstance(v, (int, float)) else 0 for v in werte]
        for r in r_pers:
            v = ws.cell(r, 5).value
            if v:
                namen.setdefault(name, []).append(str(v))
    return bedarf, namen


def gruppen_zuordnung(personen, kr_namen):
    """Ordnet Eichhof-Leuten eine Gruppe zu: zuerst über 'Funktion', dann über den Namen."""
    tokens = []
    for gruppe, liste in kr_namen.items():
        for n in liste:
            for t in norm(n):
                if t not in ("mann", "fahrer", "ab", "kw", "s", "g"):
                    tokens.append((t, gruppe))
    for p in personen:
        p["gruppe"] = ""
        if "EIC" not in "".join(p["wochen"]):
            continue
        f = p["funktion"].lower()
        for kw, gruppe in FUNKTION_KEYWORDS:
            if kw in f:
                p["gruppe"] = gruppe
                break
        if p["gruppe"]:
            continue
        kandidaten = norm(p["vorname"]) + norm(p["nachname"])
        best = (0, "")
        for k in kandidaten:
            for t, gruppe in tokens:
                s = difflib.SequenceMatcher(None, k, t).ratio()
                if s > best[0]:
                    best = (s, gruppe)
        if best[0] >= 0.85:
            p["gruppe"] = best[1]
    return personen


def lese_freiland(pfad):
    ws = openpyxl.load_workbook(pfad)["Freiland 2024"]
    zeilen = []
    for r in range(5, 45):
        a = ws.cell(r, 1)
        werte = [ws.cell(r, 2 + w).value if w <= 52 else None for w in range(1, NW + 1)]
        werte = [v if isinstance(v, (int, float)) else None for v in werte]
        gruen = bool(a.fill and a.fill.fill_type and a.fill.fgColor.rgb == "FF92D050")
        zeilen.append({"zeile": r, "label": (a.value or ""), "notiz": ws.cell(r, 2).value,
                       "werte": werte, "gruen": gruen, "fett": bool(a.font.b)})
    return zeilen


def ist_nach_geschlecht(personen, code):
    """Eingeplante Frauen/Männer pro Woche für einen Bereich (2 Bereiche in einer Woche = je 0.5)."""
    f, m = [0.0] * NW, [0.0] * NW
    for p in personen:
        for i, v in enumerate(p["wochen"]):
            if code in v:
                anteil = 1 / (1 + v.count("/"))
                if p["fm"] == "F":
                    f[i] += anteil
                elif p["fm"] == "M":
                    m[i] += anteil
    return f, m
