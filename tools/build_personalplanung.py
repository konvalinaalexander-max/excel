#!/usr/bin/env python3
"""Baut das Master-File "Personalplanung" (Excel 2016, ohne Makros).

Blätter:
  Planung          Personalchefin: alle Mitarbeiter, pro Woche Bereich / F / K / U
  Übersicht        Personalchefin: IST und SOLL pro Bereich, getrennt nach Frauen und Männern
  Kräuter Eichhof  Kräuter-Leitung: Gruppen mit Namen (automatisch), Bedarf / Ist / Differenz
  Gemüse Freiland  Gemüse-Leitung: Stundenplanung wie bisher, darunter Soll / Ist / Differenz
  Anleitung, Listen

Aufruf:
  python3 build_personalplanung.py --gesamt G.xlsx --kraeuter K.xlsx --freiland F.xlsx --out OUT.xlsx
"""
import argparse
import datetime

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import SeriesLabel
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import Outline

import quellen as q

JAHR = 2026
NW = q.NW

AREAS = [  # Code, Name (Hauptbereich), Farbe
    ("EIC", "Kräuter Eichhof", "DDEBF7"), ("JUC", "Kräuter Juch", "EDEDED"),
    ("FEL", "Kräuter Feld", "C6E0B4"), ("FRE", "Gemüse Freiland", "E2EFDA"),
    ("TOM", "Tomaten", "FBE5D6"), ("BLU", "Blumen", "FFF2CC"), ("LOG", "Logistik", "D6DCE4"),
]
STATUS = [("F", "Ferien", "B4A7D6"), ("K", "Krank", "FF7C80"), ("U", "Unfall", "FF99FF")]
FARBE = {c: f for c, _n, f in AREAS + STATUS}
FARBE_KOMBI = "E4DFEC"

# ---------------------------------------------------------------- Stil
CALIBRI = "Calibri"
F_NORM = Font(name=CALIBRI, size=11)
F_BOLD = Font(name=CALIBRI, size=11, bold=True)
F_TITEL = Font(name=CALIBRI, size=14, bold=True)
F_KLEIN = Font(name=CALIBRI, size=9, color="808080")
F_GRAU = Font(name=CALIBRI, size=11, color="595959")
F_LINK = Font(name=CALIBRI, size=11, color="0563C1", underline="single")
FILL = lambda c: PatternFill("solid", fgColor=c)  # noqa: E731
FILL_KOPF = FILL("FFC000")      # Kopfzeile Personen (wie bisher)
FILL_DATUM = FILL("FFFF00")     # Eintritt / Austritt (wie bisher)
FILL_KW = FILL("FFF2CC")        # Wochen-Kopf (wie bisher)
FILL_RECHNET = FILL("F2F2F2")   # rechnet Excel
FILL_SOLL_GEM = FILL("FFFF00")  # "Anzahl Personen" im Freiland (wie bisher)
DUENN = Side(style="thin", color="BFBFBF")
MITTEL = Side(style="medium", color="000000")
DOPPELT = Side(style="double", color="000000")
B_UNTEN_MITTEL = Border(bottom=MITTEL)
B_OBEN_MITTEL = Border(top=MITTEL)
B_UNTEN_DOPPELT = Border(bottom=DOPPELT)
B_RASTER = Border(left=DUENN, right=DUENN, top=DUENN, bottom=DUENN)
ZENTRIERT = Alignment(horizontal="center", vertical="center")
ZAHL = 'General;-General;"-"'
ROT = dict(fill=FILL("FFC7CE"), font=Font(color="9C0006"))
GELB = dict(fill=FILL("FFEB9C"), font=Font(color="9C5700"))

JAHR_REF = "Listen!$B$3"
STD_REF = "Listen!$B$4"
HAT53_REF = "Listen!$B$5"

# ---------------------------------------------------------------- Planung-Layout
P1, P2 = 3, 402                 # Datenzeilen (400 Plätze)
PCOL = dict(badge=1, vorname=2, nachname=3, fm=4, partner=5, fschein=6, bereich=7, gruppe=8,
            gruppe2=9, zustand=10, eintritt=11, austritt=12, kwe=13, kwa=14, funktion=15,
            bemerkung=16)
P_W0 = 18                       # Spalte R = KW 1
P_KEY1, P_KEY2 = P_W0 + NW + 1, P_W0 + NW + 2   # versteckte Sortierschlüssel Eichhof


def pw(w):
    return CL(P_W0 + w - 1)


def prng(col_letter):
    return f"Planung!{col_letter}${P1}:{col_letter}${P2}"


def pfix(key):
    c = CL(PCOL[key])
    return f"Planung!${c}${P1}:${c}${P2}"


def anteil(col_letter, code_ref):
    r = prng(col_letter)
    return f'ISNUMBER(SEARCH({code_ref},{r}))/(1+LEN({r})-LEN(SUBSTITUTE({r},"/","")))'


def f_ist(col_letter, code_ref, geschlecht=None, extra=""):
    teil = anteil(col_letter, code_ref)
    if geschlecht:
        teil = f'({pfix("fm")}="{geschlecht}")*{teil}'
    return f"=SUMPRODUCT({teil}{extra})"


def montag(kw_cell):
    return (f"DATE({JAHR_REF},1,4)-WEEKDAY(DATE({JAHR_REF},1,4),2)+1+({kw_cell}-1)*7")


MONATE = '"Jan","Feb","Mär","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez"'


def kw_kopf(ws, zeile_kw, zeile_oben, col0, oben="datum", fmt="0"):
    """Wochenkopf: Zeile mit KW-Nummern und darüber Montag (Datum) oder Monat."""
    for w in range(1, NW + 1):
        col = col0 + w - 1
        c = ws.cell(zeile_kw, col, w)
        c.font, c.fill, c.alignment, c.number_format = F_BOLD, FILL_KW, ZENTRIERT, fmt
        c.border = B_UNTEN_MITTEL
        k = f"{CL(col)}${zeile_kw}"
        o = ws.cell(zeile_oben, col)
        if oben == "datum":
            o.value = f"={montag(k)}"
            o.number_format, o.font, o.alignment = "dd.mm.", F_KLEIN, ZENTRIERT
        else:  # Monat des Donnerstags (ISO-Regel), nur beim Monatswechsel anzeigen
            do = f"({montag(k)}+3)"
            if w == 1:
                o.value = f"=CHOOSE(MONTH({do}),{MONATE})"
            else:
                kp = f"{CL(col - 1)}${zeile_kw}"
                o.value = (f'=IF(MONTH({do})<>MONTH(({montag(kp)}+3)),'
                           f'CHOOSE(MONTH({do}),{MONATE}),"")')
            o.font, o.fill = F_BOLD, FILL_KW
    # aktuelle Woche rot, KW 53 grau falls es sie im Jahr nicht gibt
    rng = f"{CL(col0)}{zeile_kw}:{CL(col0 + NW - 1)}{zeile_kw}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f"AND({CL(col0)}${zeile_kw}=_xlfn.ISOWEEKNUM(TODAY()),YEAR(TODAY())={JAHR_REF})"],
        fill=FILL("C00000"), font=Font(color="FFFFFF", bold=True)))
    return rng


def grau_kw53(ws, col, r1, r2):
    ws.conditional_formatting.add(f"{CL(col)}{r1}:{CL(col)}{r2}", FormulaRule(
        formula=[f'{HAT53_REF}="nein"'], fill=FILL("D9D9D9"), font=Font(color="A6A6A6")))


def diff_farben(ws, rng):
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["-0.5"], **ROT))
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0.5"], **GELB))


# ================================================================ Listen
def blatt_listen(wb):
    ws = wb.create_sheet("Listen")
    ws["A1"], ws["A1"].font = "Einstellungen und Auswahllisten", F_TITEL
    ws["A3"], ws["B3"] = "Jahr", JAHR
    ws["A4"], ws["B4"] = "Stunden pro Person und Woche (Gemüse)", 52
    ws["A5"], ws["B5"] = "Gibt es KW 53 in diesem Jahr?", f'=IF(_xlfn.ISOWEEKNUM(DATE(B3,12,28))=53,"ja","nein")'
    for c in ("B3", "B4"):
        ws[c].border = B_RASTER
    ws["B5"].fill = FILL_RECHNET
    kopf = {"A7": "Kürzel", "B7": "Bedeutung", "D7": "Hauptbereich", "F7": "Gruppen Eichhof",
            "G7": "Frauen/Männer", "I7": "Zustand", "K7": "F/M"}
    for k, v in kopf.items():
        ws[k], ws[k].font, ws[k].fill = v, F_BOLD, FILL_KOPF
    for i, (code, name, farbe) in enumerate(AREAS + STATUS):
        ws.cell(8 + i, 1, code).fill = FILL(farbe)
        ws.cell(8 + i, 2, name)
    for i, (_c, name, _f) in enumerate(AREAS):
        ws.cell(8 + i, 4, name)
    for i, (name, g, *_r) in enumerate(q.GRUPPEN):
        ws.cell(8 + i, 6, name)
        ws.cell(8 + i, 7, g)
    for i, v in enumerate(["ok", "FA"]):
        ws.cell(8 + i, 9, v)
    for i, v in enumerate(["F", "M"]):
        ws.cell(8 + i, 11, v)
    ws["A20"] = "Zwei Bereiche in einer Woche: beide Kürzel mit / trennen, z.B. BLU/EIC (zählt je halb)."
    ws["A20"].font = F_GRAU
    for col, wdt in zip("ABCDEFGHIJK", [38, 18, 2, 18, 2, 26, 14, 2, 10, 2, 6]):
        ws.column_dimensions[col].width = wdt
    n_codes, n_gr = len(AREAS) + len(STATUS), len(q.GRUPPEN)
    return {
        "codes": f"Listen!$A$8:$A${7 + n_codes}",
        "bereiche": f"Listen!$D$8:$D${7 + len(AREAS)}",
        "gruppen": f"Listen!$F$8:$F${7 + n_gr}",
        "zustand": "Listen!$I$8:$I$9", "fm": "Listen!$K$8:$K$9",
    }


# ================================================================ Planung
def blatt_planung(wb, personen, listen):
    ws = wb.create_sheet("Planung")
    ws["A1"] = f'="Planung "&{JAHR_REF}'
    ws["A1"].font = F_TITEL
    titel = {"badge": "Badge", "vorname": "Vorname", "nachname": "Nachname", "fm": "F/ M",
             "partner": "Partner", "fschein": "FSchein", "bereich": "Bereich, Haupt-",
             "gruppe": "Gruppe Eichhof", "gruppe2": "2. Gruppe (50%)", "zustand": "Zustand",
             "eintritt": "Eintritt", "austritt": "Austritt", "kwe": "KW Eintritt",
             "kwa": "KW Austritt", "funktion": "Funktion", "bemerkung": "Bemerkung"}
    for key, col in PCOL.items():
        c = ws.cell(2, col, titel[key])
        c.font, c.border = F_BOLD, B_UNTEN_MITTEL
        c.fill = FILL_DATUM if key in ("eintritt", "austritt", "kwe", "kwa") else FILL_KOPF
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.cell(2, PCOL["gruppe"]).fill = FILL(FARBE["EIC"])
    ws.cell(2, PCOL["gruppe2"]).fill = FILL(FARBE["EIC"])
    kw_kopf(ws, 2, 1, P_W0, oben="datum")
    ws.row_dimensions[2].height = 30

    for i, p in enumerate(personen):
        r = P1 + i
        werte = dict(badge=p["badge"], vorname=p["vorname"], nachname=p["nachname"], fm=p["fm"],
                     partner=p["partner"], fschein=p["fschein"], bereich=p["bereich"],
                     gruppe=p.get("gruppe") or None, zustand=p["zustand"],
                     eintritt=p["eintritt"] if isinstance(p["eintritt"], datetime.datetime) else None,
                     austritt=p["austritt"] if isinstance(p["austritt"], datetime.datetime) else None,
                     funktion=p["funktion"] or None, bemerkung=p["bemerkung"] or None)
        for key, v in werte.items():
            if v not in (None, ""):
                ws.cell(r, PCOL[key], v)
        for w, v in enumerate(p["wochen"], start=1):
            if v:
                ws.cell(r, P_W0 + w - 1, v)
    for r in range(P1, P2 + 1):
        for key in ("eintritt", "austritt"):
            ws.cell(r, PCOL[key]).number_format = "dd.mm.yyyy"
        e, a = CL(PCOL["eintritt"]), CL(PCOL["austritt"])
        for key, src in (("kwe", e), ("kwa", a)):
            c = ws.cell(r, PCOL[key], f'=IF({src}{r}="","",_xlfn.ISOWEEKNUM({src}{r}))')
            c.fill, c.font = FILL_RECHNET, F_GRAU
        # versteckte Sortierschlüssel für das Blatt Kräuter Eichhof
        w1, wN = pw(1), pw(NW)
        g1, g2 = CL(PCOL["gruppe"]), CL(PCOL["gruppe2"])
        ws.cell(r, P_KEY1, f'=IF(COUNTIF(${w1}{r}:${wN}{r},"*EIC*")=0,"",'
                           f'IFERROR(MATCH(${g1}{r},{listen["gruppen"]},0),9)*1000+ROW())')
        ws.cell(r, P_KEY2, f'=IF(OR(${CL(P_KEY1)}{r}="",${g2}{r}=""),"",'
                           f'IFERROR(MATCH(${g2}{r},{listen["gruppen"]},0)*1000+ROW(),""))')
        for col in range(P_W0, P_W0 + NW):
            ws.cell(r, col).alignment = ZENTRIERT

    # Auswahllisten
    def dv(formula, rng, prompt=None, style="stop", error=None):
        d = DataValidation(type="list", formula1=f"={formula}", allow_blank=True,
                           showErrorMessage=True, errorStyle=style)
        if error:
            d.error, d.errorTitle = error, "Planung"
        if prompt:
            d.prompt, d.promptTitle, d.showInputMessage = prompt, "Planung", True
        d.add(rng)
        ws.add_data_validation(d)

    col = lambda k: CL(PCOL[k])  # noqa: E731
    dv(listen["fm"], f"{col('fm')}{P1}:{col('fm')}{P2}")
    dv(listen["bereiche"], f"{col('bereich')}{P1}:{col('bereich')}{P2}")
    dv(listen["gruppen"], f"{col('gruppe')}{P1}:{col('gruppe2')}{P2}",
       prompt="Nur für Leute bei Kräuter Eichhof. 2. Gruppe = je 50%.")
    dv(listen["zustand"], f"{col('zustand')}{P1}:{col('zustand')}{P2}")
    wrng = f"{pw(1)}{P1}:{pw(NW)}{P2}"
    dv(listen["codes"], wrng, style="information",
       error="Dieses Kürzel ist nicht in der Liste. Für zwei Bereiche in einer Woche z.B. BLU/EIC "
             "schreiben. Mit OK wird der Eintrag übernommen.")

    # Farben im Wochenraster (wie bisher je Bereich), zuerst Kombinationen
    tl = f"{pw(1)}{P1}"
    ws.conditional_formatting.add(wrng, FormulaRule(
        formula=[f'ISNUMBER(SEARCH("/",{tl}))'], fill=FILL(FARBE_KOMBI), font=Font(italic=True)))
    for code, _n, farbe in AREAS:
        ws.conditional_formatting.add(wrng, FormulaRule(
            formula=[f'ISNUMBER(SEARCH("{code}",{tl}))'], fill=FILL(farbe)))
    for code, _n, farbe in STATUS:
        ws.conditional_formatting.add(wrng, CellIsRule(operator="equal", formula=[f'"{code}"'],
                                                       fill=FILL(farbe), font=Font(bold=True)))
    grau_kw53(ws, P_W0 + NW - 1, P1, P2)

    breiten = dict(badge=7, vorname=13, nachname=16, fm=5, partner=7, fschein=7, bereich=15,
                   gruppe=19, gruppe2=12, zustand=8, eintritt=10.5, austritt=10.5, kwe=8.5, kwa=8.5,
                   funktion=20, bemerkung=26)
    for key, wdt in breiten.items():
        ws.column_dimensions[CL(PCOL[key])].width = wdt
    ws.column_dimensions[CL(17)].width = 1.5
    for w in range(1, NW + 1):
        ws.column_dimensions[pw(w)].width = 5.6
    ws.column_dimensions[CL(P_KEY1 - 1)].width = 1.5
    for c in (P_KEY1, P_KEY2):
        ws.column_dimensions[CL(c)].hidden = True
    ws.freeze_panes = "D3"
    ws.auto_filter.ref = f"A2:{CL(P_KEY2)}{P2}"
    return ws


# ================================================================ Kräuter Eichhof
K_W0 = 8  # Spalte H = KW 1


def kw_k(w):
    return CL(K_W0 + w - 1)


def blatt_kraeuter(wb, personen, bedarf):
    ws = wb.create_sheet("Kräuter Eichhof")
    kopf = ["Nr", "f/m", "Gruppe", "Name", "Funktion", "Anteil"]
    for i, t in enumerate(kopf, start=1):
        c = ws.cell(2, i, t)
        c.font, c.border = F_BOLD, B_UNTEN_MITTEL
    ws["A1"] = "Kräuter Eichhof – Gruppen"
    ws["A1"].font = F_TITEL
    kw_kopf(ws, 2, 1, K_W0, oben="datum")
    keys = f"Planung!${CL(P_KEY1)}${P1}:${CL(P_KEY2)}${P2}"
    zaehl = {}
    for p in personen:
        if "EIC" in "".join(p["wochen"]):
            zaehl[p.get("gruppe") or ""] = zaehl.get(p.get("gruppe") or "", 0) + 1

    r = 3
    zeilen = {}
    gruppen = [(i + 1, n, g) for i, (n, g, *_x) in enumerate(q.GRUPPEN)] + [(9, "ohne Gruppe", "")]
    for gi, name, geschl in gruppen:
        platz = max(6, zaehl.get(name if gi != 9 else "", 0) + (4 if gi != 9 else 6))
        start = f'COUNTIF({keys},"<"&{gi}*1000)'
        r_first = r
        for j in range(1, platz + 1):
            kth = f"SMALL({keys},{start}+{j})"
            ws.cell(r, 7, f'=IFERROR(IF(INT({kth}/1000)={gi},MOD({kth},1000),""),"")')
            ws.cell(r, 1, f'=IF($G{r}="","",{j})')
            ws.cell(r, 2, f'=IF($G{r}="","",LOWER(INDEX(Planung!$D:$D,$G{r})))')
            ws.cell(r, 3, f'=IF($G{r}="","","{name}")')
            ws.cell(r, 4, f'=IF($G{r}="","",TRIM(INDEX(Planung!$B:$B,$G{r})&" "&INDEX(Planung!$C:$C,$G{r})))')
            if gi == 9:
                ws.cell(r, 5, f'=IF($G{r}="","",HYPERLINK("#Planung!H"&$G{r},"Gruppe in Planung wählen"))').font = F_LINK
                ws.cell(r, 6, f'=IF($G{r}="","",1)')
            else:
                ws.cell(r, 5, f'=IF($G{r}="","",INDEX(Planung!$O:$O,$G{r})&"")')
                ws.cell(r, 6, f'=IF($G{r}="","",IF(INDEX(Planung!$I:$I,$G{r})="",1,0.5))')
            ws.cell(r, 6).number_format = '[=1]"";0%'
            for w in range(1, NW + 1):
                c = f"INDEX(Planung!{pw(w)}$1:{pw(w)}${P2},$G{r})"
                teil = f'$F{r}/(1+LEN({c})-LEN(SUBSTITUTE({c},"/","")))'
                ws.cell(r, K_W0 + w - 1,
                        f'=IF($G{r}="","",IF(OR({c}={{"F","K","U"}}),UPPER({c}),'
                        f'IF(ISNUMBER(SEARCH("EIC",{c})),IF({teil}=1,"x",{teil}),"")))').alignment = ZENTRIERT
            r += 1
        grid = f"{kw_k(1)}{r_first}:{kw_k(NW)}{r - 1}"
        zeilen[name] = {"personen": (r_first, r - 1)}
        # Hinweis, wenn die Plätze nicht reichen
        voll = f'COUNTIFS({keys},">="&{gi}*1000,{keys},"<"&({gi}+1)*1000)'
        if gi != 9:
            ws.cell(r, 3, name).font = F_BOLD
            ws.cell(r, 4, "Bedarf").font = F_BOLD
            ws.cell(r, 5, f'=IF({voll}>{platz},"Achtung: "&({voll}-{platz})&" Personen ohne Platz","")').font = Font(color="C00000")
            for w in range(1, NW + 1):
                c = ws.cell(r, K_W0 + w - 1, bedarf[name][w - 1])
                c.font, c.alignment, c.number_format = F_BOLD, ZENTRIERT, ZAHL
            r_bed = r
            r += 1
            ws.cell(r, 4, "Ist (eingeplant)")
            g_ref = f'"{name}"'
            extra = (f'*(({pfix("gruppe")}={g_ref})+({pfix("gruppe2")}={g_ref}))'
                     f'/(1+({pfix("gruppe2")}<>""))')
            for w in range(1, NW + 1):
                c = ws.cell(r, K_W0 + w - 1, f_ist(pw(w), '"EIC"', extra=extra))
                c.fill, c.alignment, c.number_format = FILL_RECHNET, ZENTRIERT, ZAHL
            r_ist = r
            r += 1
            ws.cell(r, 4, "Differenz (Ist − Bedarf)")
            for w in range(1, NW + 1):
                c = ws.cell(r, K_W0 + w - 1, f"=ROUND({kw_k(w)}{r_ist}-{kw_k(w)}{r_bed},1)")
                c.fill, c.alignment, c.number_format = FILL_RECHNET, ZENTRIERT, ZAHL
            diff_farben(ws, f"{kw_k(1)}{r}:{kw_k(NW)}{r}")
            for col in range(1, K_W0 + NW):
                ws.cell(r, col).border = B_UNTEN_DOPPELT
            zeilen[name].update(bedarf=r_bed, ist=r_ist, diff=r, geschlecht=geschl)
            r += 1
        else:
            ws.cell(r, 3, "ohne Gruppe").font = F_BOLD
            ws.cell(r, 4, "Ist (eingeplant)")
            ws.cell(r, 5, f'=IF({voll}>{platz},"Achtung: "&({voll}-{platz})&" Personen ohne Platz","")').font = Font(color="C00000")
            for w in range(1, NW + 1):
                c = ws.cell(r, K_W0 + w - 1, f_ist(pw(w), '"EIC"', extra=f'*({pfix("gruppe")}="")'))
                c.fill, c.alignment, c.number_format = FILL_RECHNET, ZENTRIERT, ZAHL
            for col in range(1, K_W0 + NW):
                ws.cell(r, col).border = B_UNTEN_DOPPELT
            zeilen[name].update(ist=r)
            r += 1
        # Farben der Personen-Wochen
        tl = f"{kw_k(1)}{r_first}"
        ws.conditional_formatting.add(grid, CellIsRule(operator="equal", formula=['"x"'], fill=FILL(FARBE["EIC"])))
        ws.conditional_formatting.add(grid, FormulaRule(formula=[f"ISNUMBER({tl})"], fill=FILL("EEF5FB")))
        for code, _n, farbe in STATUS:
            ws.conditional_formatting.add(grid, CellIsRule(operator="equal", formula=[f'"{code}"'], fill=FILL(farbe)))
        r += 1  # Leerzeile

    # Zusammenfassung (wie bisher "Ein-Austritt")
    ws.cell(r, 3, "Total Eichhof").font = F_TITEL
    r += 1
    summen = {}
    gr = [n for n, *_x in q.GRUPPEN]
    for titel, geschl in (("Total", None), ("Frauen", "F"), ("Männer", "M")):
        bed_rows = [zeilen[n]["bedarf"] for n in gr if geschl is None or zeilen[n]["geschlecht"] == geschl]
        ws.cell(r, 3, titel).font = F_BOLD
        ws.cell(r, 4, "Bedarf").font = F_BOLD
        ws.cell(r + 1, 4, "Ist (eingeplant)")
        ws.cell(r + 2, 4, "Differenz (Ist − Bedarf)")
        for w in range(1, NW + 1):
            col = kw_k(w)
            c = ws.cell(r, K_W0 + w - 1, "=ROUND(" + "+".join(f"{col}{x}" for x in bed_rows) + ",1)")
            c.font, c.fill, c.alignment, c.number_format = F_BOLD, FILL_RECHNET, ZENTRIERT, ZAHL
            c = ws.cell(r + 1, K_W0 + w - 1, f_ist(pw(w), '"EIC"', geschlecht=geschl))
            c.fill, c.alignment, c.number_format = FILL_RECHNET, ZENTRIERT, ZAHL
            c = ws.cell(r + 2, K_W0 + w - 1, f"=ROUND({col}{r + 1}-{col}{r},1)")
            c.fill, c.alignment, c.number_format = FILL_RECHNET, ZENTRIERT, ZAHL
        diff_farben(ws, f"{kw_k(1)}{r + 2}:{kw_k(NW)}{r + 2}")
        for col in range(1, K_W0 + NW):
            ws.cell(r + 2, col).border = B_UNTEN_DOPPELT
        summen[titel] = {"bedarf": r, "ist": r + 1, "diff": r + 2}
        r += 4
    ws.cell(r, 4, "Weisse Zahlen in den Zeilen \"Bedarf\" tippt die Kräuter-Leitung selbst. "
                  "Graue Felder und Namen rechnet Excel aus dem Blatt Planung.").font = F_GRAU

    grau_kw53(ws, K_W0 + NW - 1, 3, r)
    for col, wdt in zip("ABCDEFG", [4, 4, 22, 24, 22, 7, 4]):
        ws.column_dimensions[col].width = wdt
    ws.column_dimensions["G"].hidden = True
    for w in range(1, NW + 1):
        ws.column_dimensions[kw_k(w)].width = 5.2
    ws.freeze_panes = f"{kw_k(1)}3"
    return ws, summen


# ================================================================ Gemüse Freiland
G_W0 = 3  # Spalte C = KW 1


def kw_g(w):
    return CL(G_W0 + w - 1)


def blatt_gemuese(wb, zeilen_freiland):
    ws = wb.create_sheet("Gemüse Freiland")
    ws["A1"] = f'="Arbeitsplanung "&{JAHR_REF}&" Freiland"'
    ws["A1"].font = F_TITEL
    ws["A2"], ws["A3"] = "Monat", "KW"
    ws["A2"].font = ws["A3"].font = F_BOLD
    kw_kopf(ws, 3, 2, G_W0, oben="monat")
    summe_col = G_W0 + NW
    ws.cell(2, summe_col, "Summe").font = F_BOLD
    for z in zeilen_freiland:
        r = z["zeile"]
        a = ws.cell(r, 1, z["label"] or None)
        a.font = F_BOLD
        if z["gruen"]:
            a.fill = FILL("92D050")
            ws.cell(r, 2).fill = FILL("92D050")
        if z["notiz"] not in (None, ""):
            ws.cell(r, 2, z["notiz"])
        for w, v in enumerate(z["werte"], start=1):
            c = ws.cell(r, G_W0 + w - 1, v)
            c.border = B_RASTER
            c.alignment = ZENTRIERT
        if z["label"]:
            s = ws.cell(r, summe_col, f"=SUM({kw_g(1)}{r}:{kw_g(NW)}{r})")
            s.number_format, s.fill = "0", FILL_RECHNET
    r_sum, r_soll, r_ist, r_diff, r_jaet = 46, 47, 48, 49, 51
    ws.cell(r_sum, 1, "Summe Wochenstunden").font = F_BOLD
    ws.cell(r_soll, 1, "Soll: Anzahl Personen").font = F_BOLD
    ws.cell(r_soll, 2, f'=" (Stunden / "&{STD_REF}&")"').font = F_GRAU
    ws.cell(r_ist, 1, "Ist: eingeplante Personen").font = F_BOLD
    ws.cell(r_ist, 2, "(aus Planung)").font = F_GRAU
    ws.cell(r_diff, 1, "Differenz (Ist − Soll)").font = F_BOLD
    ws.cell(r_jaet, 1, "Jätstunden (automatisierbar)").font = F_BOLD
    for w in range(1, NW + 1):
        col = kw_g(w)
        c = ws.cell(r_sum, G_W0 + w - 1, f"=SUM({col}4:{col}45)")
        c.fill, c.number_format, c.border = FILL_RECHNET, "0", Border(top=MITTEL)
        c = ws.cell(r_soll, G_W0 + w - 1, f"={col}{r_sum}/{STD_REF}")
        c.fill, c.number_format, c.font = FILL_SOLL_GEM, "0.0", F_BOLD
        c = ws.cell(r_ist, G_W0 + w - 1, f_ist(pw(w), '"FRE"'))
        c.fill, c.number_format = FILL_RECHNET, ZAHL
        c = ws.cell(r_diff, G_W0 + w - 1, f"=ROUND({col}{r_ist}-{col}{r_soll},1)")
        c.fill, c.number_format = FILL_RECHNET, ZAHL
        c.border = B_UNTEN_MITTEL
        c = ws.cell(r_jaet, G_W0 + w - 1, f"={col}8+{col}17+{col}21+{col}25+{col}28")
        c.fill, c.number_format = FILL_RECHNET, "0"
        for rr in (r_sum, r_soll, r_ist, r_diff, r_jaet):
            ws.cell(rr, G_W0 + w - 1).alignment = ZENTRIERT
    diff_farben(ws, f"{kw_g(1)}{r_diff}:{kw_g(NW)}{r_diff}")
    for rr in (r_sum, r_soll, r_jaet):
        s = ws.cell(rr, summe_col, f"=SUM({kw_g(1)}{rr}:{kw_g(NW)}{rr})")
        s.number_format, s.fill = "0", FILL_RECHNET
    ws.cell(r_diff + 1, 1, "Rot = es fehlen Leute, gelb = mehr Leute eingeplant als gebraucht.").font = F_GRAU

    chart = BarChart()
    chart.title = "Freiland: Soll und Ist (Personen pro Woche)"
    chart.y_axis.title, chart.x_axis.title = "Personen", "KW"
    chart.add_data(Reference(ws, min_col=G_W0, max_col=G_W0 + NW - 1, min_row=r_soll), from_rows=True)
    chart.series[0].tx = SeriesLabel(v="Soll")
    chart.series[0].graphicalProperties.solidFill = "92D050"
    chart.set_categories(Reference(ws, min_col=G_W0, max_col=G_W0 + NW - 1, min_row=3))
    line = LineChart()
    line.add_data(Reference(ws, min_col=G_W0, max_col=G_W0 + NW - 1, min_row=r_ist), from_rows=True)
    line.series[0].tx = SeriesLabel(v="Ist")
    line.series[0].graphicalProperties.line.solidFill = "1F4E78"
    line.series[0].graphicalProperties.line.width = 28000
    line.series[0].smooth = False
    chart += line
    _achsen(chart)
    chart.width, chart.height = 32, 9
    ws.add_chart(chart, f"A{r_jaet + 2}")

    grau_kw53(ws, G_W0 + NW - 1, 3, r_jaet)
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 16
    for w in range(1, NW + 1):
        ws.column_dimensions[kw_g(w)].width = 5.2
    ws.column_dimensions[CL(summe_col)].width = 9
    ws.freeze_panes = f"{kw_g(1)}4"
    return ws, {"soll": r_soll}


def _achsen(chart):
    chart.x_axis.number_format = "0"
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    chart.legend.position = "b"


# ================================================================ Übersicht
U_W0 = 5  # Spalte E = KW 1


def kw_u(w):
    return CL(U_W0 + w - 1)


def runde(x, schritt=0.5):
    return round(x / schritt) * schritt


def aufteilen(total, f_ist, m_ist):
    """Startwert: SOLL nach dem Frauen-Anteil der Planung 2026 aufteilen (pro Woche)."""
    jahr_f, jahr_t = sum(f_ist), sum(f_ist) + sum(m_ist)
    anteil_jahr = jahr_f / jahr_t if jahr_t else 0.5
    f_out, m_out = [], []
    for t, f, m in zip(total, f_ist, m_ist):
        anteil_f = f / (f + m) if f + m else anteil_jahr
        fv = runde((t or 0) * anteil_f)
        f_out.append(fv)
        m_out.append(round((t or 0) - fv, 1))
    return f_out, m_out


def blatt_uebersicht(wb, personen, soll, kr_summen, gem):
    ws = wb.create_sheet("Übersicht")
    ws.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=True)
    ws["B1"] = f'="Übersicht "&{JAHR_REF}'
    ws["B1"].font = F_TITEL
    ws["B2"], ws["C2"] = "Bereich", "Beschreibung"
    ws["B2"].font = ws["C2"].font = F_BOLD
    ws["B2"].border = ws["C2"].border = B_UNTEN_MITTEL
    kw_kopf(ws, 2, 1, U_W0, oben="monat", fmt='"KW "0')

    def zeile(r, label, text, werte=None, formel=None, eingabe=False, fett=False, farbe=None,
              diff=False, unten=None, einzug=False):
        b = ws.cell(r, 2, ("   " if einzug else "") + label)
        c = ws.cell(r, 3, text)
        b.font = F_BOLD if (fett or not einzug) else F_NORM
        c.font = F_GRAU if einzug else F_NORM
        if farbe:
            b.fill = c.fill = FILL(farbe)
        for w in range(1, NW + 1):
            cell = ws.cell(r, U_W0 + w - 1)
            if formel:
                f = formel(w)
                cell.value = f"=ROUND({f[1:]},1)" if not f.startswith("=COUNT") else f
            elif werte is not None:
                cell.value = werte[w - 1]
            cell.number_format = ZAHL
            cell.alignment = ZENTRIERT
            if fett:
                cell.font = F_BOLD
            if not eingabe:
                cell.fill = FILL_RECHNET
            else:
                cell.border = B_RASTER
        if diff:
            diff_farben(ws, f"{kw_u(1)}{r}:{kw_u(NW)}{r}")
        if unten:
            for col in range(2, U_W0 + NW):
                ws.cell(r, col).border = unten

    def gruppiere(r1, r2, level=1):
        for rr in range(r1, r2 + 1):
            ws.row_dimensions[rr].outline_level = level

    # ---------- Gesamtbetrieb (Formeln werden unten nachgetragen, wenn alle Bereiche stehen)
    R = {"g_vertraege": 4, "g_ist": 5, "g_ist_f": 6, "g_ist_m": 7, "g_soll": 8, "g_soll_f": 9,
         "g_soll_m": 10, "g_diff": 11, "g_diff_f": 12, "g_diff_m": 13, "g_abw": 14,
         "g_reserve": 15, "g_diff_res": 16, "g_vj": 17}
    ws.cell(3, 2, "Gesamtbetrieb").font = F_BOLD
    for col in range(2, U_W0 + NW):
        ws.cell(3, col).border = B_OBEN_MITTEL

    # ---------- Bereiche
    bereiche = [
        ("BLU", "Blumen", "eingabe_fm", "Summe Erststempelung Vorjahr (Startwert, bitte anpassen)"),
        ("FRE", "Freiland", "freiland", "aus Blatt Gemüse Freiland (Stunden / 52)"),
        ("LOG", "Logistik", "logistik", "fixe Posten gemäss Logistik-Leitung"),
        ("FEL", "Kräuter Feld", "eingabe_fm", "Minimum Teamgrösse gemäss Teamleiter Feld"),
        ("TOM", "Tomaten", "eingabe_fm", "Planzahlen Tomaten (Eingabe)"),
        ("EIC", "Kräuter Eichhof, Stecklinge, Weihnachtssterne", "eichhof",
         "aus Blatt Kräuter Eichhof (Bedarf je Gruppe)"),
        ("JUC", "Kräuter Juch", "eingabe_fm", "Erststempelung Vorjahr / Plan Juch (Startwert, bitte anpassen)"),
    ]
    r = 26
    blocks = {}
    for code, name, art, quelle in bereiche:
        farbe = FARBE[code]
        f_ist26, m_ist26 = q.ist_nach_geschlecht(personen, code)
        ws.cell(r, 1, code)
        ws.cell(r, 2, name).font = F_BOLD
        for col in range(2, U_W0 + NW):
            ws.cell(r, col).border = B_OBEN_MITTEL
        code_ref = f"$A${r}"
        b = {"kopf": r}
        r += 1
        zeile(r, "IST", "eingeplant gemäss Planung", formel=lambda w: f_ist(pw(w), code_ref), farbe=farbe, fett=True)
        b["ist"] = r
        zeile(r + 1, "Frauen", "", formel=lambda w: f_ist(pw(w), code_ref, "F"), einzug=True, farbe=farbe)
        zeile(r + 2, "Männer", "", formel=lambda w: f_ist(pw(w), code_ref, "M"), einzug=True, farbe=farbe)
        b["ist_f"], b["ist_m"] = r + 1, r + 2
        gruppiere(r + 1, r + 2)
        r += 3
        b["soll"], b["soll_f"], b["soll_m"] = r, r + 1, r + 2
        if art == "eingabe_fm":
            if code == "FEL":
                f_start, m_start = [0] * NW, list(soll["FEL"])
            else:
                f_start, m_start = aufteilen(soll[code], f_ist26, m_ist26)
            zeile(r, "SOLL", quelle, formel=lambda w: f"={kw_u(w)}{r + 1}+{kw_u(w)}{r + 2}", farbe=farbe, fett=True)
            zeile(r + 1, "Frauen", "Eingabe", werte=f_start, eingabe=True, einzug=True, farbe=farbe)
            zeile(r + 2, "Männer", "Eingabe", werte=m_start, eingabe=True, einzug=True, farbe=farbe)
            gruppiere(r + 1, r + 2)
            r += 3
        elif art == "freiland":
            zeile(r, "SOLL", quelle, formel=lambda w: f"='Gemüse Freiland'!{kw_g(w)}{gem['soll']}", farbe=farbe, fett=True)
            zeile(r + 1, "Frauen", "Eingabe: fix angestellte Frauen", werte=[round(x, 1) for x in f_ist26],
                  eingabe=True, einzug=True, farbe=farbe)
            zeile(r + 2, "Männer", "SOLL minus Frauen",
                  formel=lambda w: f"=MAX(0,{kw_u(w)}{r}-{kw_u(w)}{r + 1})", einzug=True, farbe=farbe)
            gruppiere(r + 1, r + 2)
            r += 3
        elif art == "logistik":
            details = list(soll["LOG"].items())
            n = len(details)
            zeile(r, "SOLL", quelle, formel=lambda w: f"={kw_u(w)}{r + 1}+{kw_u(w)}{r + 2}", farbe=farbe, fett=True)
            zeile(r + 1, "Frauen", "Eingabe", werte=[0] * NW, eingabe=True, einzug=True, farbe=farbe)
            zeile(r + 2, "Männer", "Summe der Posten",
                  formel=lambda w: f"=SUM({kw_u(w)}{r + 3}:{kw_u(w)}{r + 2 + n})", einzug=True, farbe=farbe)
            for i, (lab, werte) in enumerate(details):
                zeile(r + 3 + i, "  " + lab, "Eingabe", werte=werte, eingabe=True, einzug=True)
            gruppiere(r + 1, r + 2)
            gruppiere(r + 3, r + 2 + n, level=2)
            r += 3 + n
        elif art == "eichhof":
            k = "'Kräuter Eichhof'!"
            zeile(r, "SOLL", quelle, formel=lambda w: f"={k}{kw_k(w)}{kr_summen['Total']['bedarf']}", farbe=farbe, fett=True)
            zeile(r + 1, "Frauen", "Gruppen Steckling, Fax", formel=lambda w: f"={k}{kw_k(w)}{kr_summen['Frauen']['bedarf']}",
                  einzug=True, farbe=farbe)
            zeile(r + 2, "Männer", "Gruppen Verteilen, Luchs", formel=lambda w: f"={k}{kw_k(w)}{kr_summen['Männer']['bedarf']}",
                  einzug=True, farbe=farbe)
            gruppiere(r + 1, r + 2)
            r += 3
        b["diff"], b["diff_f"], b["diff_m"] = r, r + 1, r + 2
        zeile(r, "IST - SOLL", "minus = es fehlen Leute",
              formel=lambda w: f"={kw_u(w)}{b['ist']}-{kw_u(w)}{b['soll']}", farbe=farbe, fett=True, diff=True)
        zeile(r + 1, "Frauen", "", formel=lambda w: f"={kw_u(w)}{b['ist_f']}-{kw_u(w)}{b['soll_f']}",
              einzug=True, farbe=farbe, diff=True)
        zeile(r + 2, "Männer", "", formel=lambda w: f"={kw_u(w)}{b['ist_m']}-{kw_u(w)}{b['soll_m']}",
              einzug=True, farbe=farbe, diff=True, unten=B_UNTEN_MITTEL)
        gruppiere(r + 1, r + 2)
        r += 4
        blocks[code] = b

    def summe(key):
        return lambda w: "=" + "+".join(f"{kw_u(w)}{blocks[c][key]}" for c in blocks)

    wr = lambda w: prng(pw(w))  # noqa: E731
    zeile(R["g_vertraege"], "Verträge", "Personen mit Eintrag in dieser Woche (inkl. Abwesende)",
          formel=lambda w: f"=COUNTA({wr(w)})", fett=True)
    zeile(R["g_ist"], "IST", "eingeplant (ohne Ferien / Krank / Unfall)", formel=summe("ist"), fett=True)
    zeile(R["g_ist_f"], "Frauen", "", formel=summe("ist_f"), einzug=True)
    zeile(R["g_ist_m"], "Männer", "", formel=summe("ist_m"), einzug=True)
    zeile(R["g_soll"], "SOLL", "Bedarf (Summe Bereiche)", formel=summe("soll"), fett=True)
    zeile(R["g_soll_f"], "Frauen", "", formel=summe("soll_f"), einzug=True)
    zeile(R["g_soll_m"], "Männer", "", formel=summe("soll_m"), einzug=True)
    zeile(R["g_diff"], "IST - SOLL", "minus = es fehlen Leute", formel=summe("diff"), fett=True, diff=True)
    zeile(R["g_diff_f"], "Frauen", "", formel=summe("diff_f"), einzug=True, diff=True)
    zeile(R["g_diff_m"], "Männer", "", formel=summe("diff_m"), einzug=True, diff=True)
    zeile(R["g_abw"], "Abwesend", "Ferien (F), Krank (K), Unfall (U) gemäss Planung",
          formel=lambda w: f'=COUNTIF({wr(w)},"F")+COUNTIF({wr(w)},"K")+COUNTIF({wr(w)},"U")')
    zeile(R["g_reserve"], "Reserve Krank", "Durchschnitt Krank 2023–2025 (Eingabe)",
          werte=[round(x, 1) for x in soll["KRANK"]], eingabe=True)
    zeile(R["g_diff_res"], "IST - SOLL", "inkl. Reserve Krank",
          formel=lambda w: f"={kw_u(w)}{R['g_diff']}-{kw_u(w)}{R['g_reserve']}", fett=True, diff=True)
    zeile(R["g_vj"], "VJ", "Verträge Vorjahr gemäss Mobatime (Eingabe)",
          werte=[round(x, 1) for x in soll["VJ"]], eingabe=True, unten=B_UNTEN_MITTEL)
    for a, bb in ((R["g_ist_f"], R["g_ist_m"]), (R["g_soll_f"], R["g_soll_m"]), (R["g_diff_f"], R["g_diff_m"])):
        gruppiere(a, bb)

    # ---------- Wen anstellen? (Zeilen 19-24)
    ws.cell(19, 2, "Wen anstellen?").font = F_BOLD
    ws.cell(20, 2, "von KW")
    ws.cell(21, 2, "bis KW")
    ws["C20"], ws["C21"] = 23, 35
    for c in ("C20", "C21"):
        ws[c].border, ws[c].alignment = B_RASTER, Alignment(horizontal="left")
    kw_rng = f"${kw_u(1)}$2:${kw_u(NW)}$2"

    def check(r_diff):
        d = f"${kw_u(1)}${r_diff}:${kw_u(NW)}${r_diff}"
        n = f'COUNTIFS({d},"<-0.5",{kw_rng},">="&$C$20,{kw_rng},"<="&$C$21)'
        mx = f"-SUMPRODUCT(MIN({d}+(({kw_rng}<$C$20)+({kw_rng}>$C$21))*1000))"
        return n, mx

    nf, mf = check(R["g_diff_f"])
    nm, mm = check(R["g_diff_m"])
    wochen = "($C$21-$C$20+1)"
    ws.cell(22, 2, "Frauen")
    ws.cell(22, 3, f'=IF({nf}=0,"genug Frauen eingeplant","in "&{nf}&" von "&{wochen}&" Wochen fehlen Frauen, höchstens "&ROUND({mf},1))')
    ws.cell(23, 2, "Männer")
    ws.cell(23, 3, f'=IF({nm}=0,"genug Männer eingeplant","in "&{nm}&" von "&{wochen}&" Wochen fehlen Männer, höchstens "&ROUND({mm},1))')
    ws.cell(24, 2, "Empfehlung").font = F_BOLD
    ws.cell(24, 3, f'=IF(AND({nf}>0,{nm}>0),"Frauen und Männer fehlen – ein Paar passt",'
                   f'IF({nm}>0,"Es fehlen Männer – eher Männer anstellen",'
                   f'IF({nf}>0,"Es fehlen Frauen – eher Frauen anstellen","In diesem Zeitraum fehlt niemand")))').font = F_BOLD
    for rr in (22, 23, 24):
        ws.cell(rr, 3).fill = FILL_RECHNET

    # ---------- Grafiken
    r += 1
    ws.cell(r, 2, "Grafik").font = F_TITEL
    ws["C19"] = f'=HYPERLINK("#\'Übersicht\'!B{r}","Grafik Frauen / Männer ansehen")'
    ws["C19"].font = F_LINK
    for i, (titel, r_soll, r_ist) in enumerate((("Frauen gesamt: SOLL und IST", R["g_soll_f"], R["g_ist_f"]),
                                                 ("Männer gesamt: SOLL und IST", R["g_soll_m"], R["g_ist_m"]))):
        ch = LineChart()
        ch.title = titel
        ch.y_axis.title, ch.x_axis.title = "Personen", "KW"
        for rr, name, farbe in ((r_soll, "SOLL", "C00000"), (r_ist, "IST", "1F4E78")):
            ch.add_data(Reference(ws, min_col=U_W0, max_col=U_W0 + NW - 1, min_row=rr), from_rows=True)
            s = ch.series[-1]
            s.tx = SeriesLabel(v=name)
            s.graphicalProperties.line.solidFill = farbe
            s.graphicalProperties.line.width = 28000
            s.smooth = False
        ch.set_categories(Reference(ws, min_col=U_W0, max_col=U_W0 + NW - 1, min_row=2))
        _achsen(ch)
        ch.width, ch.height = 30, 8
        ws.add_chart(ch, f"B{r + 1 + i * 17}")

    ws.column_dimensions["A"].hidden = True
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 50
    ws.column_dimensions["D"].width = 1.5
    for w in range(1, NW + 1):
        ws.column_dimensions[kw_u(w)].width = 5.8
    ws.freeze_panes = f"{kw_u(1)}3"
    grau_kw53(ws, U_W0 + NW - 1, 3, r)
    ws.grafik_zeile = r
    return ws


# ================================================================ Anleitung
ANLEITUNG = [
    ("Personalplanung – kurze Anleitung", "titel"),
    ("", None),
    ("Grundregeln", "fett"),
    ("1. Jeder schreibt nur in sein Blatt: Personalchefin = Planung und Übersicht, "
     "Kräuter-Leitung = Kräuter Eichhof, Gemüse-Leitung = Gemüse Freiland.", None),
    ("2. Weisse Felder werden von Hand ausgefüllt. Graue Felder rechnet Excel – dort nichts eintippen.", None),
    ("3. Nichts nach rechts ziehen. Alle Wochen des Jahres sind schon vorbereitet.", None),
    ("4. Das File liegt auf dem Netzlaufwerk: nach dem Eintragen speichern und schliessen, "
     "damit die anderen schreiben können.", None),
    ("", None),
    ("Blatt Planung (Personalchefin)", "fett"),
    ("Eine Zeile pro Person. In jede Woche kommt, wo die Person arbeitet:", None),
    ("   EIC Kräuter Eichhof · JUC Kräuter Juch · FEL Kräuter Feld · FRE Gemüse Freiland · "
     "TOM Tomaten · BLU Blumen · LOG Logistik", None),
    ("   F = Ferien · K = Krank · U = Unfall (immer ganze Wochen)", None),
    ("Zwei Bereiche in derselben Woche: beide Kürzel mit / trennen, z.B. BLU/EIC. Die Person zählt dann je halb.", None),
    ("Mehrere Wochen auf einmal: Wochen markieren, Kürzel tippen, Ctrl+Enter.", None),
    ("Spalte \"Gruppe Eichhof\": nur für Leute, die bei Kräuter Eichhof arbeiten (Steckling, Fax, Verteilen, Luchs). "
     "Mit \"2. Gruppe\" arbeitet die Person je zur Hälfte in beiden Gruppen.", None),
    ("Neue Person: in der nächsten leeren Zeile eintragen. Es hat Platz für 400 Personen.", None),
    ("", None),
    ("Blatt Übersicht (Personalchefin)", "fett"),
    ("Pro Bereich: IST (eingeplant), SOLL (Bedarf) und IST - SOLL, je getrennt nach Frauen und Männern.", None),
    ("Minus heisst: es fehlen Leute (rot). Plus heisst: mehr Leute eingeplant als gebraucht (gelb).", None),
    ("Mit den Knöpfen 1 / 2 / 3 oben links lassen sich die Zeilen Frauen/Männer ein- und ausblenden.", None),
    ("Der SOLL von Blumen, Tomaten, Kräuter Feld, Juch und Logistik wird hier eingetragen. "
     "Eichhof und Freiland kommen automatisch aus ihren Blättern.", None),
    ("\"Wen anstellen?\": Zeitraum eingeben – Excel zeigt, ob Frauen, Männer oder beide fehlen.", None),
    ("", None),
    ("Blatt Kräuter Eichhof (Kräuter-Leitung)", "fett"),
    ("Die Namen erscheinen automatisch, sobald die Personalchefin in der Planung EIC einträgt – sortiert nach Gruppe.", None),
    ("x = ganze Woche bei Eichhof, 0.5 = halbe Woche, F/K/U = abwesend, leer = in dieser Woche woanders.", None),
    ("Pro Gruppe die Zeile \"Bedarf\" ausfüllen. Darunter rechnet Excel Ist und Differenz.", None),
    ("Leute ohne Gruppe stehen unten bei \"ohne Gruppe\". Mit dem Link springt man direkt in die Planung, "
     "um die Gruppe zu wählen.", None),
    ("", None),
    ("Blatt Gemüse Freiland (Gemüse-Leitung)", "fett"),
    ("Wie bisher: Stunden pro Tätigkeit und Woche. Soll = Summe Stunden / 52.", None),
    ("Darunter neu: Ist (wie viele Leute eingeplant sind) und die Differenz.", None),
    ("", None),
    ("Neues Jahr", "fett"),
    ("File kopieren, im Blatt Listen das Jahr ändern, in der Planung die Wochen leeren. "
     "Der Bedarf des alten Jahres ist der Startwert für das neue.", None),
]


def blatt_anleitung(wb):
    ws = wb.create_sheet("Anleitung")
    for i, (t, art) in enumerate(ANLEITUNG, start=1):
        c = ws.cell(i, 1, t)
        c.font = F_TITEL if art == "titel" else (F_BOLD if art == "fett" else F_NORM)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.column_dimensions["A"].width = 120
    return ws


# ================================================================ main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--gesamt", required=True)
    ap.add_argument("--kraeuter", required=True)
    ap.add_argument("--freiland", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--vorschau", action="store_true", help="nur für Bildvorschau: Druckbereiche setzen")
    a = ap.parse_args()

    personen, soll = q.lese_gesamt(a.gesamt)
    bedarf, kr_namen = q.lese_kraeuter(a.kraeuter)
    q.gruppen_zuordnung(personen, kr_namen)
    freiland = q.lese_freiland(a.freiland)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    listen = blatt_listen(wb)
    blatt_planung(wb, personen, listen)
    _kr, kr_summen = blatt_kraeuter(wb, personen, bedarf)
    _gm, gem = blatt_gemuese(wb, freiland)
    ueb = blatt_uebersicht(wb, personen, soll, kr_summen, gem)
    blatt_anleitung(wb)
    wb._sheets = [wb[n] for n in ("Planung", "Übersicht", "Kräuter Eichhof", "Gemüse Freiland",
                                  "Anleitung", "Listen")]
    wb.active = 0
    for ws in wb.worksheets:
        ws.sheet_view.zoomScale = 90
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
    if a.vorschau:
        g = ueb.grafik_zeile
        bereiche = {"Planung": "A1:AH40", "Übersicht": f"B{g}:AH{g + 35}", "Kräuter Eichhof": "A1:AH60",
                    "Gemüse Freiland": "A1:AF75", "Anleitung": "A1:A40", "Listen": "A1:K20"}
        for name, rng in bereiche.items():
            ws = wb[name]
            ws.print_area = rng
            ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 1
            ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.calculation.fullCalcOnLoad = True
    wb.save(a.out)
    print("gespeichert:", a.out)


if __name__ == "__main__":
    main()
