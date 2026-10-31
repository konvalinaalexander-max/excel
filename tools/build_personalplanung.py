#!/usr/bin/env python3
"""Baut das Master-File "Personalplanung" (Excel 2016, ohne Makros).

Blätter:
  Planung          Personalchefin: alle Mitarbeiter, pro Woche ein oder zwei Bereiche / F / K / U
  Übersicht        Personalchefin: IST und SOLL pro Bereich, getrennt nach Frauen und Männern
  Grafik           Bereich wählen: Frauen / Männer / Total, eingeplant gegen Bedarf
  Kräuter Eichhof  Kräuter-Leitung: neue Leute, Gruppen mit Namen, Bedarf / Ist / Differenz,
                   Gruppen zuteilen
  Gemüse Freiland  Gemüse-Leitung: Stundenplanung wie bisher, darunter Soll / Ist / Differenz
  Anleitung, Listen

Aufruf:
  python3 build_personalplanung.py --gesamt G.xlsx --kraeuter K.xlsx --freiland F.xlsx --out OUT.xlsx
"""
import argparse
import datetime
from itertools import combinations

import openpyxl
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.chart.series import SeriesLabel
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import Outline

import quellen as q

JAHR = 2026
NW = q.NW

# ============================================================ Farben (Pastell wie in den Originalen)
AREAS = [  # Code, Hauptbereich, Kurzname, Farbe
    ("EIC", "Kräuter Eichhof", "Eichhof", "DDEBF7"),
    ("JUC", "Kräuter Juch", "Juch", "E7E6E6"),
    ("FEL", "Kräuter Feld", "Feld", "D5E8E4"),
    ("FRE", "Gemüse Freiland", "Freiland", "E2EFDA"),
    ("TOM", "Tomaten", "Tomaten", "FBE5D6"),
    ("BLU", "Blumen", "Blumen", "FFF2CC"),
    ("LOG", "Logistik", "Logistik", "EFE6D8"),
]
STATUS = [("F", "Ferien", "DCD3EC"), ("K", "Krank", "F4CCCC"), ("U", "Unfall", "E6B8B7")]
FARBE = {a[0]: a[3] for a in AREAS} | {s[0]: s[2] for s in STATUS}
KURZ = {a[0]: a[2] for a in AREAS}

C_TITEL, C_TEXT, C_GRAU, C_GRAU_D = "1F3864", "262626", "7F7F7F", "595959"
C_LINIE, C_LINIE_D, C_KOPF, C_RECHNET = "D9D9D9", "A6A6A6", "F2F2F2", "F5F5F5"
C_BAND, C_BANDING = "1F3864", "F7F9FC"
C_FEHLT, C_FEHLT_T, C_ZUVIEL, C_ZUVIEL_T = "F9D0D0", "9C0006", "FFE8A3", "7F6000"
C_IST_BALKEN, C_SOLL_LINIE = "9DC3E6", "1F3864"


def font(size=10, bold=False, color=C_TEXT, italic=False, underline=None):
    return Font(name="Calibri", size=size, bold=bold, color=color, italic=italic, underline=underline)


def fill(c):
    return PatternFill("solid", fgColor=c)


F_TITEL = font(16, True, C_TITEL)
F_UNTERTITEL = font(10, color=C_GRAU, italic=True)
F_BAND = font(11, True, "FFFFFF")
F_ABSCHNITT = font(12, True)
F_KOPF = font(10, True)
F_GRUPPE = font(9, True, C_GRAU_D)
F_TEXT = font(10)
F_TEXT_B = font(10, True)
F_TEXT_G = font(10, color=C_GRAU_D)
F_KLEIN = font(9, color=C_GRAU)
F_DATUM = font(8, color=C_GRAU)
F_LINK = font(10, color="0563C1", underline="single")

S_HAAR = Side(style="thin", color=C_LINIE)
S_DUNKEL = Side(style="thin", color=C_LINIE_D)
S_MITTEL = Side(style="medium", color=C_GRAU_D)
RAHMEN = Border(left=S_HAAR, right=S_HAAR, top=S_HAAR, bottom=S_HAAR)
RAHMEN_EINGABE = Border(left=S_DUNKEL, right=S_DUNKEL, top=S_DUNKEL, bottom=S_DUNKEL)
UNTEN_HAAR = Border(bottom=S_HAAR)
UNTEN_DUNKEL = Border(bottom=S_DUNKEL)
UNTEN_MITTEL = Border(bottom=S_MITTEL)

LINKS = Alignment(horizontal="left", vertical="center")
MITTE = Alignment(horizontal="center", vertical="center")
MITTE_SCHRUMPF = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
UMBRUCH = Alignment(horizontal="left", vertical="center", wrap_text=True)
ZAHL = 'General;-General;"-"'

JAHR_REF, STD_REF, HAT53_REF = "Listen!$C$4", "Listen!$C$5", "Listen!$C$6"
MONATE = '"Jan","Feb","Mär","Apr","Mai","Jun","Jul","Aug","Sep","Okt","Nov","Dez"'


def zelle(ws, r, c, wert=None, f=None, bg=None, rand=None, ausr=None, fmt=None):
    cell = ws.cell(r, c)
    if wert is not None:
        cell.value = wert
    if f:
        cell.font = f
    if bg:
        cell.fill = fill(bg)
    if rand:
        cell.border = rand
    if ausr:
        cell.alignment = ausr
    if fmt:
        cell.number_format = fmt
    return cell


def blatt_grundlage(ws, titel, untertitel, tab):
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_properties.tabColor = tab
    zelle(ws, 1, 1, titel, F_TITEL, ausr=Alignment(vertical="center"))
    zelle(ws, 2, 1, untertitel, F_UNTERTITEL, ausr=Alignment(vertical="center"))
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 18


def legende(ws, zeile, col0, eintraege, breite=3):
    """Farbige Legende: Liste (Text, Farbe, Schrift) als zusammengefasste Felder."""
    col = col0
    for text, farbe, fnt in eintraege:
        for c in range(col, col + breite):
            zelle(ws, zeile, c, bg=farbe, rand=RAHMEN)
        zelle(ws, zeile, col, text, fnt or font(9), ausr=MITTE)
        ws.merge_cells(start_row=zeile, start_column=col, end_row=zeile, end_column=col + breite - 1)
        col += breite + 1


def band(ws, zeile, c1, c2, text, dunkel=True, farbe=None, hoehe=20, linie_bis=None):
    bg = C_BAND if dunkel else farbe
    for c in range(c1, c2 + 1):
        zelle(ws, zeile, c, bg=bg)
    zelle(ws, zeile, c1, text, F_BAND if dunkel else F_ABSCHNITT, ausr=LINKS)
    if linie_bis:
        for c in range(c2 + 1, linie_bis + 1):
            zelle(ws, zeile, c, rand=UNTEN_DUNKEL)
    ws.row_dimensions[zeile].height = hoehe


def abstand(ws, zeile, hoehe=7):
    ws.row_dimensions[zeile].height = hoehe


def montag(kw_cell):
    return f"DATE({JAHR_REF},1,4)-WEEKDAY(DATE({JAHR_REF},1,4),2)+1+({kw_cell}-1)*7"


def kw_kopf(ws, zeile_kw, zeile_oben, col0, oben="datum", fmt="0"):
    """Wochenkopf: KW-Nummern, darüber Datum (Montag) oder Monat. Monate abwechselnd hinterlegt,
    aktuelle Woche dunkel."""
    for w in range(1, NW + 1):
        col = col0 + w - 1
        zelle(ws, zeile_kw, col, w, F_KOPF, rand=UNTEN_MITTEL, ausr=MITTE, fmt=fmt)
        k = f"{CL(col)}${zeile_kw}"
        if oben == "datum":
            zelle(ws, zeile_oben, col, f"={montag(k)}", F_DATUM, ausr=MITTE, fmt="dd.mm.")
        else:
            do = f"({montag(k)}+3)"
            if w == 1:
                v = f"=CHOOSE(MONTH({do}),{MONATE})"
            else:
                kp = f"{CL(col - 1)}${zeile_kw}"
                v = f'=IF(MONTH({do})<>MONTH(({montag(kp)}+3)),CHOOSE(MONTH({do}),{MONATE}),"")'
            zelle(ws, zeile_oben, col, v, font(9, True, C_GRAU_D), ausr=LINKS)
    c0, c1 = CL(col0), CL(col0 + NW - 1)
    rng = f"{c0}{zeile_kw}:{c1}{zeile_kw}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f"AND({c0}${zeile_kw}=_xlfn.ISOWEEKNUM(TODAY()),YEAR(TODAY())={JAHR_REF})"],
        fill=fill(C_BAND), font=Font(color="FFFFFF", bold=True)))
    ws.conditional_formatting.add(f"{c0}{zeile_oben}:{c1}{zeile_kw}", FormulaRule(
        formula=[f"ISODD(MONTH({montag(f'{c0}${zeile_kw}')}+3))"], fill=fill(C_KOPF)))
    ws.conditional_formatting.add(f"{c1}{zeile_oben}:{c1}{zeile_kw}", FormulaRule(
        formula=[f'{HAT53_REF}="nein"'], font=Font(color="BFBFBF")))


def diff_farben(ws, rng):
    ws.conditional_formatting.add(rng, CellIsRule(operator="lessThan", formula=["-0.5"],
                                                  fill=fill(C_FEHLT), font=Font(color=C_FEHLT_T, bold=True)))
    ws.conditional_formatting.add(rng, CellIsRule(operator="greaterThan", formula=["0.5"],
                                                  fill=fill(C_ZUVIEL), font=Font(color=C_ZUVIEL_T)))


def kombi_grafik(titel, ist_ref, soll_ref, kat_ref, farbe_ist=C_IST_BALKEN, farbe_soll=C_SOLL_LINIE,
                 name_ist="eingeplant (IST)", name_soll="Bedarf (SOLL)", breite=28, hoehe=8.5):
    """Säulen = eingeplant, Linie = Bedarf. Legende rechts, keine Achsentitel (keine Überlappung)."""
    bar = BarChart()
    bar.type, bar.grouping, bar.gapWidth = "col", "clustered", 40
    bar.add_data(ist_ref, from_rows=True)
    s = bar.series[0]
    s.tx = SeriesLabel(v=name_ist)
    s.graphicalProperties.solidFill = farbe_ist
    s.graphicalProperties.line.solidFill = farbe_ist
    bar.set_categories(kat_ref)
    line = LineChart()
    line.add_data(soll_ref, from_rows=True)
    s = line.series[0]
    s.tx = SeriesLabel(v=name_soll)
    s.graphicalProperties.line.solidFill = farbe_soll
    s.graphicalProperties.line.width = 25000
    s.smooth = False
    s.marker = Marker(symbol="none")
    line.set_categories(kat_ref)
    bar += line
    bar.title = titel
    bar.legend.position = "r"
    bar.legend.overlay = False
    bar.x_axis.number_format = "0"
    bar.x_axis.tickLblPos = "low"
    for ax in (bar.x_axis, bar.y_axis):
        ax.delete = False
    bar.y_axis.number_format = "0"
    bar.width, bar.height = breite, hoehe
    return bar


# ============================================================ Planung: Lage der Spalten
P1, P2 = 5, 404                       # Datenzeilen (400 Plätze)
PCOL = dict(badge=1, vorname=2, nachname=3, fm=4, partner=5, fschein=6, bereich=7, gruppe=8,
            gruppe2=9, zustand=10, eintritt=11, austritt=12, kwe=13, kwa=14, funktion=15,
            bemerkung=16)
P_W0 = 18                             # Spalte R = KW 1
P_KEY1 = P_W0 + NW + 1                # versteckt: Sortierschlüssel Eichhof (1. Gruppe)
P_KEY2, P_NAME, P_NEU = P_KEY1 + 1, P_KEY1 + 2, P_KEY1 + 3


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


def kuerzel_liste():
    """Auswahlliste für die Wochen: F, K, U, dann Bereiche und Paare alphabetisch."""
    codes = sorted(a[0] for a in AREAS)
    out = [(s[0], s[1]) for s in STATUS]
    for c in codes:
        out.append((c, dict((a[0], a[1]) for a in AREAS)[c]))
        for a, b in combinations(codes, 2):
            if a == c:
                out.append((f"{a}/{b}", f"je halb {KURZ[a]} und {KURZ[b]}"))
    return out


# ============================================================ Listen
def blatt_listen(wb):
    ws = wb.create_sheet("Listen")
    blatt_grundlage(ws, "Einstellungen und Listen", "Hier nur das Jahr ändern. Die Listen werden für die Auswahlfelder gebraucht.", "A6A6A6")
    for r, (lab, wert) in enumerate([("Jahr", JAHR), ("Stunden pro Person und Woche (Gemüse)", 52)], start=4):
        zelle(ws, r, 2, lab, F_TEXT)
        zelle(ws, r, 3, wert, F_TEXT_B, rand=RAHMEN_EINGABE, ausr=MITTE)
    zelle(ws, 6, 2, "Gibt es KW 53 in diesem Jahr?", F_TEXT_G)
    zelle(ws, 6, 3, '=IF(_xlfn.ISOWEEKNUM(DATE(C4,12,28))=53,"ja","nein")', F_TEXT_G, bg=C_RECHNET, ausr=MITTE)

    kuerzel = kuerzel_liste()
    kopf = {2: "Auswahl Woche", 3: "Bedeutung", 5: "Hauptbereich", 7: "Gruppen Eichhof", 8: "F/M",
            10: "Zustand", 12: "F/M", 14: "Neu bei Eichhof (automatisch)"}
    for c, t in kopf.items():
        zelle(ws, 8, c, t, F_KOPF, bg=C_KOPF, rand=UNTEN_MITTEL)
    for i, (code, text) in enumerate(kuerzel):
        haupt = code.split("/")[0]
        zelle(ws, 9 + i, 2, code, F_TEXT_B, bg=FARBE[haupt] if "/" not in code else None, rand=UNTEN_HAAR)
        zelle(ws, 9 + i, 3, text, F_TEXT, rand=UNTEN_HAAR)
    for i, a in enumerate(AREAS):
        zelle(ws, 9 + i, 5, a[1], F_TEXT, bg=a[3], rand=UNTEN_HAAR)
    for i, (name, g, *_r) in enumerate(q.GRUPPEN):
        zelle(ws, 9 + i, 7, name, F_TEXT, rand=UNTEN_HAAR)
        zelle(ws, 9 + i, 8, g, F_TEXT, rand=UNTEN_HAAR, ausr=MITTE)
    for i, v in enumerate(["ok", "FA"]):
        zelle(ws, 9 + i, 10, v, F_TEXT, rand=UNTEN_HAAR)
    for i, v in enumerate(["F", "M"]):
        zelle(ws, 9 + i, 12, v, F_TEXT, rand=UNTEN_HAAR, ausr=MITTE)
    neu_key = f"Planung!${CL(P_NEU)}${P1}:${CL(P_NEU)}${P2}"
    for k in range(1, 151):
        zelle(ws, 8 + k, 14, f'=IFERROR(INDEX(Planung!${CL(P_NAME)}:${CL(P_NAME)},SMALL({neu_key},{k})),"")', F_TEXT_G)
    for col, wdt in zip(range(1, 15), [2, 36, 34, 2, 18, 2, 24, 6, 2, 9, 2, 6, 2, 30]):
        ws.column_dimensions[CL(col)].width = wdt
    n = len(kuerzel)
    wb.defined_names["NeueEichhof"] = DefinedName(
        "NeueEichhof", attr_text='OFFSET(Listen!$N$9,0,0,MAX(1,COUNTIF(Listen!$N$9:$N$158,"?*")),1)')
    return {"codes": f"Listen!$B$9:$B${8 + n}", "bereiche": f"Listen!$E$9:$E${8 + len(AREAS)}",
            "gruppen": f"Listen!$G$9:$G${8 + len(q.GRUPPEN)}", "zustand": "Listen!$J$9:$J$10",
            "fm": "Listen!$L$9:$L$10"}


# ============================================================ Planung
def blatt_planung(wb, personen, listen, k_tab):
    ws = wb.create_sheet("Planung")
    blatt_grundlage(ws, f'="Planung "&{JAHR_REF}',
                    "Eine Zeile pro Person. Pro Woche einen Bereich wählen – oder zwei (z.B. BLU/FRE = je halb). "
                    "F Ferien · K Krank · U Unfall. Mehrere Wochen: markieren, wählen, Ctrl+Enter.", "1F3864")
    leg = [(f"{a[0]} {a[2]}", a[3], None) for a in AREAS] + [(f"{s[0]} {s[1]}", s[2], None) for s in STATUS]
    leg.append(("BLU/FRE = je halb", "FFFFFF", font(9, italic=True)))
    legende(ws, 1, P_W0, leg, breite=3)

    gruppen = [("Person", "badge", "fschein"), ("Einteilung", "bereich", "gruppe2"),
               ("Vertrag", "zustand", "kwa"), ("Notizen", "funktion", "bemerkung")]
    for text, a, b in gruppen:
        ca, cb = PCOL[a], PCOL[b]
        for c in range(ca, cb + 1):
            zelle(ws, 3, c, rand=UNTEN_DUNKEL)
        zelle(ws, 3, ca, text.upper(), F_GRUPPE, ausr=LINKS)
        ws.merge_cells(start_row=3, start_column=ca, end_row=3, end_column=cb)
    titel = {"badge": "Badge", "vorname": "Vorname", "nachname": "Nachname", "fm": "F/M",
             "partner": "Partner", "fschein": "FSchein", "bereich": "Hauptbereich",
             "gruppe": "Gruppe Eichhof", "gruppe2": "2. Gruppe", "zustand": "Zustand",
             "eintritt": "Eintritt", "austritt": "Austritt", "kwe": "KW Eintritt",
             "kwa": "KW Austritt", "funktion": "Funktion", "bemerkung": "Bemerkung"}
    for key, col in PCOL.items():
        zelle(ws, 4, col, titel[key], F_KOPF, bg=C_KOPF, rand=UNTEN_MITTEL, ausr=UMBRUCH)
    for key in ("gruppe", "gruppe2", "kwe", "kwa"):
        ws.cell(4, PCOL[key]).font = font(10, True, C_GRAU_D)
    kw_kopf(ws, 4, 3, P_W0, oben="datum")
    ws.row_dimensions[3].height = 16
    ws.row_dimensions[4].height = 32

    for i, p in enumerate(personen):
        r = P1 + i
        werte = dict(badge=p["badge"], vorname=p["vorname"], nachname=p["nachname"], fm=p["fm"],
                     partner=p["partner"], fschein=p["fschein"], bereich=p["bereich"],
                     zustand=p["zustand"],
                     eintritt=p["eintritt"] if isinstance(p["eintritt"], datetime.datetime) else None,
                     austritt=p["austritt"] if isinstance(p["austritt"], datetime.datetime) else None,
                     funktion=p["funktion"] or None, bemerkung=p["bemerkung"] or None)
        for key, v in werte.items():
            if v not in (None, ""):
                ws.cell(r, PCOL[key], v)
        for w, v in enumerate(p["wochen"], start=1):
            if v:
                ws.cell(r, P_W0 + w - 1, v)

    w1, wN = pw(1), pw(NW)
    g1, g2 = CL(PCOL["gruppe"]), CL(PCOL["gruppe2"])
    vn, nn = CL(PCOL["vorname"]), CL(PCOL["nachname"])
    for r in range(P1, P2 + 1):
        for key in PCOL:
            c = ws.cell(r, PCOL[key])
            c.font, c.border, c.alignment = F_TEXT, UNTEN_HAAR, LINKS
        for key in ("fm", "partner", "fschein", "zustand", "kwe", "kwa"):
            ws.cell(r, PCOL[key]).alignment = MITTE
        for key in ("eintritt", "austritt"):
            ws.cell(r, PCOL[key]).number_format = "dd.mm.yyyy"
            ws.cell(r, PCOL[key]).alignment = MITTE
        for key, src in (("kwe", CL(PCOL["eintritt"])), ("kwa", CL(PCOL["austritt"]))):
            ws.cell(r, PCOL[key], f'=IF({src}{r}="","",_xlfn.ISOWEEKNUM({src}{r}))').font = F_TEXT_G
        # Gruppe kommt aus dem Blatt Kräuter Eichhof (dort teilt die Kräuter-Leitung zu)
        for key, kcol in (("gruppe", k_tab["gruppe"]), ("gruppe2", k_tab["gruppe2"])):
            ws.cell(r, PCOL[key], f'=IFERROR(INDEX({kcol},MATCH(${CL(P_NAME)}{r},{k_tab["name"]},0))&"","")').font = F_TEXT_G
        for col in range(P_W0, P_W0 + NW):
            c = ws.cell(r, col)
            c.alignment, c.border, c.font = MITTE_SCHRUMPF, RAHMEN, font(9)
        ws.cell(r, P_NAME, f'=IF(${vn}{r}="","",TRIM(${vn}{r}&" "&${nn}{r}))')
        eic = f'COUNTIF(${w1}{r}:${wN}{r},"*EIC*")'
        ws.cell(r, P_KEY1, f'=IF({eic}=0,"",IFERROR(MATCH(${g1}{r},{listen["gruppen"]},0),9)*1000+ROW())')
        ws.cell(r, P_KEY2, f'=IF(OR(${CL(P_KEY1)}{r}="",${g2}{r}=""),"",IFERROR(MATCH(${g2}{r},{listen["gruppen"]},0)*1000+ROW(),""))')
        ws.cell(r, P_NEU, f'=IF(AND(${CL(P_NAME)}{r}<>"",{eic}>0,${g1}{r}=""),ROW(),"")')

    def dv(formula, rng, style="stop", error=None, prompt=None):
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
    dv(listen["zustand"], f"{col('zustand')}{P1}:{col('zustand')}{P2}")
    wrng = f"{pw(1)}{P1}:{pw(NW)}{P2}"
    dv(listen["codes"], wrng, error="Bitte aus der Liste wählen. Zwei Bereiche in einer Woche: z.B. BLU/FRE.")

    # Farben
    bcol = f"{col('bereich')}{P1}:{col('bereich')}{P2}"
    for code, name, _k, farbe in AREAS:
        ws.conditional_formatting.add(bcol, CellIsRule(operator="equal", formula=[f'"{name}"'], fill=fill(farbe)))
    links = f"A{P1}:{col('bemerkung')}{P2}"
    ws.conditional_formatting.add(links, FormulaRule(formula=[f"AND($B{P1}<>\"\",MOD(ROW(),2)=0)"], fill=fill(C_BANDING)))
    tl = f"{pw(1)}{P1}"
    ws.conditional_formatting.add(wrng, FormulaRule(
        formula=[f'ISNUMBER(SEARCH("/",{tl}))'], fill=fill("FFFFFF"), font=Font(italic=True, bold=True, color=C_TITEL)))
    for code, _n, _k, farbe in AREAS:
        ws.conditional_formatting.add(wrng, FormulaRule(formula=[f'ISNUMBER(SEARCH("{code}",{tl}))'], fill=fill(farbe)))
    for code, _n, farbe in STATUS:
        ws.conditional_formatting.add(wrng, CellIsRule(operator="equal", formula=[f'"{code}"'], fill=fill(farbe),
                                                       font=Font(bold=True)))
    ws.conditional_formatting.add(f"{vn}{P1}:{nn}{P2}", FormulaRule(
        formula=[f'AND(${CL(P_NAME)}{P1}<>"",COUNTIF(${CL(P_NAME)}${P1}:${CL(P_NAME)}${P2},${CL(P_NAME)}{P1})>1)'],
        fill=fill(C_FEHLT), font=Font(color=C_FEHLT_T)))

    breiten = dict(badge=7, vorname=13, nachname=16, fm=5, partner=7.5, fschein=7.5, bereich=16,
                   gruppe=17, gruppe2=12, zustand=7.5, eintritt=10.5, austritt=10.5, kwe=8, kwa=8,
                   funktion=20, bemerkung=30)
    for key, wdt in breiten.items():
        ws.column_dimensions[CL(PCOL[key])].width = wdt
    ws.column_dimensions["Q"].width = 2
    for w in range(1, NW + 1):
        ws.column_dimensions[pw(w)].width = 6
    ws.column_dimensions[CL(P_KEY1 - 1)].width = 2
    for c in (P_KEY1, P_KEY2, P_NAME, P_NEU):
        ws.column_dimensions[CL(c)].hidden = True
    ws.freeze_panes = f"D{P1}"
    ws.auto_filter.ref = f"A4:{CL(P_NEU)}{P2}"
    ws.print_title_rows = "3:4"
    return ws


# ============================================================ Kräuter Eichhof
K_W0 = 7                    # Spalte G = KW 1
K_PLAETZE_ZUTEILUNG = 150


def kw_k(w):
    return CL(K_W0 + w - 1)


def kraeuter_layout(personen):
    """Berechnet die Zeilen des Blatts Kräuter Eichhof (wird auch für die Planung gebraucht)."""
    zaehl = {}
    for p in personen:
        if "EIC" in "".join(p["wochen"]):
            zaehl[p.get("gruppe") or ""] = zaehl.get(p.get("gruppe") or "", 0) + 1
    r = 6
    lay = {"neu": (r + 1, r + 10)}
    r += 12
    bloecke = []
    gruppen = [(i + 1, n, g) for i, (n, g, *_x) in enumerate(q.GRUPPEN)] + [(9, "Noch ohne Gruppe", "")]
    for gi, name, geschl in gruppen:
        platz = max(5, zaehl.get(name if gi != 9 else "", 0) + 3)
        b = {"gi": gi, "name": name, "geschl": geschl, "kopf": r, "p1": r + 1, "p2": r + platz}
        r += platz + 1
        if gi != 9:
            b.update(bedarf=r, ist=r + 1, diff=r + 2)
            r += 3
        else:
            b.update(ist=r)
            r += 1
        bloecke.append(b)
        r += 1
    lay["bloecke"] = bloecke
    lay["total"] = r
    r += 1
    summen = {}
    for titel in ("Total", "Frauen", "Männer"):
        summen[titel] = {"kopf": r, "bedarf": r + 1, "ist": r + 2, "diff": r + 3}
        r += 5
    lay["summen"] = summen
    lay["zuteilen"] = r + 1
    lay["z1"] = r + 4
    lay["z2"] = r + 3 + K_PLAETZE_ZUTEILUNG
    return lay


def k_tab_refs(lay):
    z1, z2 = lay["z1"], lay["z2"]
    s = "'Kräuter Eichhof'!"
    return {"name": f"{s}$C${z1}:$C${z2}", "gruppe": f"{s}$D${z1}:$D${z2}", "gruppe2": f"{s}$E${z1}:$E${z2}"}


def blatt_kraeuter(wb, personen, bedarf, listen, lay):
    ws = wb.create_sheet("Kräuter Eichhof")
    blatt_grundlage(ws, f'="Kräuter Eichhof – Gruppen "&{JAHR_REF}',
                    "Namen kommen automatisch aus der Planung. Pro Gruppe die Zeile «Bedarf» ausfüllen – "
                    "Ist und Differenz rechnet Excel. Neue Leute ganz unten bei «Gruppen zuteilen» einteilen.", "5B9BD5")
    legende(ws, 1, K_W0, [("x = ganze Woche da", FARBE["EIC"], None), ("0.5 = halbe Woche", "EEF5FB", None),
                          ("F Ferien", FARBE["F"], None), ("K Krank", FARBE["K"], None), ("U Unfall", FARBE["U"], None),
                          ("leer = woanders", "FFFFFF", font(9, color=C_GRAU))], breite=4)
    for c, t in enumerate(["Nr", "f/m", "Name", "Funktion", "Anteil"], start=1):
        zelle(ws, 4, c, t, F_KOPF, bg=C_KOPF, rand=UNTEN_MITTEL, ausr=LINKS if c > 2 else MITTE)
    kw_kopf(ws, 4, 3, K_W0, oben="datum")
    ws.row_dimensions[4].height = 22
    keys = f"Planung!${CL(P_KEY1)}${P1}:${CL(P_KEY2)}${P2}"
    neu_key = f"Planung!${CL(P_NEU)}${P1}:${CL(P_NEU)}${P2}"
    letzte = K_W0 + NW - 1
    z1 = lay["z1"]

    # ---------- Neu bei Eichhof
    n1, n2 = lay["neu"]
    anzahl_neu = f'COUNT({neu_key})'
    band(ws, n1 - 1, 1, 5, f'="Neu bei Eichhof – noch keiner Gruppe zugeteilt ("&{anzahl_neu}&")"',
         dunkel=False, farbe=C_ZUVIEL, linie_bis=letzte)
    for k in range(1, 11):
        r = n1 + k - 1
        rr = f"SMALL({neu_key},{k})"
        zelle(ws, r, 2, f'=IFERROR(LOWER(INDEX(Planung!$D:$D,{rr})),"")', F_KLEIN, ausr=MITTE)
        zelle(ws, r, 3, f'=IFERROR(INDEX(Planung!${CL(P_NAME)}:${CL(P_NAME)},{rr}),IF({k}=1,"Alle Leute sind einer Gruppe zugeteilt.",""))',
              F_TEXT, rand=UNTEN_HAAR)
        zelle(ws, r, 4, f'=IFERROR("ab KW "&MATCH("*EIC*",INDEX(Planung!${pw(1)}:${pw(NW)},{rr},0),0),"")', F_TEXT_G, rand=UNTEN_HAAR)
        zelle(ws, r, 5, f'=IFERROR(IF({rr}>0,HYPERLINK("#\'Kräuter Eichhof\'!C"&({z1}+COUNTA($C${z1}:$C${lay["z2"]})),"zuteilen →"),""),"")',
              F_LINK, rand=UNTEN_HAAR)

    # ---------- Gruppen
    for b in lay["bloecke"]:
        gi, name = b["gi"], b["name"]
        grau = gi == 9
        band(ws, b["kopf"], 1, 5, name, dunkel=False, farbe=C_KOPF if grau else FARBE["EIC"], hoehe=19, linie_bis=letzte)
        if not grau:
            zelle(ws, b["kopf"], 4, "Frauen-Gruppe" if b["geschl"] == "F" else "Männer-Gruppe", font(9, italic=True, color=C_GRAU_D))
        start = f'COUNTIF({keys},"<"&{gi}*1000)'
        for j, r in enumerate(range(b["p1"], b["p2"] + 1), start=1):
            kth = f"SMALL({keys},{start}+{j})"
            ws.cell(r, 6, f'=IFERROR(IF(INT({kth}/1000)={gi},MOD({kth},1000),""),"")')
            zelle(ws, r, 1, f'=IF($F{r}="","",{j})', F_KLEIN, ausr=MITTE, rand=UNTEN_HAAR)
            zelle(ws, r, 2, f'=IF($F{r}="","",LOWER(INDEX(Planung!$D:$D,$F{r})))', F_KLEIN, ausr=MITTE, rand=UNTEN_HAAR)
            zelle(ws, r, 3, f'=IF($F{r}="","",INDEX(Planung!${CL(P_NAME)}:${CL(P_NAME)},$F{r}))', F_TEXT, rand=UNTEN_HAAR)
            if grau:
                zelle(ws, r, 4, f'=IF($F{r}="","",HYPERLINK("#\'Kräuter Eichhof\'!C"&({z1}+COUNTA($C${z1}:$C${lay["z2"]})),"zuteilen →"))',
                      F_LINK, rand=UNTEN_HAAR)
                zelle(ws, r, 5, f'=IF($F{r}="","",1)', F_KLEIN, rand=UNTEN_HAAR, fmt='[=1]"";0%')
            else:
                zelle(ws, r, 4, f'=IF($F{r}="","",INDEX(Planung!$O:$O,$F{r})&"")', F_KLEIN, rand=UNTEN_HAAR)
                zelle(ws, r, 5, f'=IF($F{r}="","",IF(INDEX(Planung!$I:$I,$F{r})="",1,0.5))', F_KLEIN,
                      rand=UNTEN_HAAR, fmt='[=1]"";0%', ausr=MITTE)
            for w in range(1, NW + 1):
                c = f"INDEX(Planung!{pw(w)}$1:{pw(w)}${P2},$F{r})"
                teil = f'$E{r}/(1+LEN({c})-LEN(SUBSTITUTE({c},"/","")))'
                zelle(ws, r, K_W0 + w - 1,
                      f'=IF($F{r}="","",IF(OR({c}={{"F","K","U"}}),UPPER({c}),'
                      f'IF(ISNUMBER(SEARCH("EIC",{c})),IF({teil}=1,"x",{teil}),"")))',
                      font(9), rand=RAHMEN, ausr=MITTE)
        grid = f"{kw_k(1)}{b['p1']}:{kw_k(NW)}{b['p2']}"
        tl = f"{kw_k(1)}{b['p1']}"
        ws.conditional_formatting.add(grid, CellIsRule(operator="equal", formula=['"x"'], fill=fill(FARBE["EIC"])))
        ws.conditional_formatting.add(grid, FormulaRule(formula=[f"ISNUMBER({tl})"], fill=fill("EEF5FB")))
        for code, _n, farbe in STATUS:
            ws.conditional_formatting.add(grid, CellIsRule(operator="equal", formula=[f'"{code}"'], fill=fill(farbe),
                                                           font=Font(bold=True)))
        voll = f'COUNTIFS({keys},">="&{gi}*1000,{keys},"<"&({gi}+1)*1000)'
        platz = b["p2"] - b["p1"] + 1
        warn = f'=IF({voll}>{platz},"Achtung: "&({voll}-{platz})&" Personen haben keinen Platz","")'
        g_ref = f'"{name}"'
        if not grau:
            extra = (f'*(({pfix("gruppe")}={g_ref})+({pfix("gruppe2")}={g_ref}))/(1+({pfix("gruppe2")}<>""))')
            zelle(ws, b["bedarf"], 3, "Bedarf", F_TEXT_B, rand=UNTEN_HAAR)
            zelle(ws, b["bedarf"], 4, warn, font(9, True, C_FEHLT_T))
            zelle(ws, b["ist"], 3, "Ist (eingeplant)", F_TEXT_G, rand=UNTEN_HAAR)
            zelle(ws, b["diff"], 3, "Differenz", F_TEXT_B)
            for w in range(1, NW + 1):
                col = K_W0 + w - 1
                zelle(ws, b["bedarf"], col, bedarf[name][w - 1], F_TEXT_B, rand=RAHMEN_EINGABE, ausr=MITTE, fmt=ZAHL)
                zelle(ws, b["ist"], col, f_ist(pw(w), '"EIC"', extra=extra), F_TEXT_G, bg=C_RECHNET, rand=RAHMEN, ausr=MITTE, fmt=ZAHL)
                zelle(ws, b["diff"], col, f"=ROUND({kw_k(w)}{b['ist']}-{kw_k(w)}{b['bedarf']},1)", F_TEXT,
                      bg=C_RECHNET, rand=RAHMEN, ausr=MITTE, fmt=ZAHL)
            for c in range(1, letzte + 1):
                ws.cell(b["diff"], c).border = Border(left=ws.cell(b["diff"], c).border.left,
                                                      right=ws.cell(b["diff"], c).border.right,
                                                      top=ws.cell(b["diff"], c).border.top, bottom=S_MITTEL)
            diff_farben(ws, f"{kw_k(1)}{b['diff']}:{kw_k(NW)}{b['diff']}")
        else:
            zelle(ws, b["ist"], 3, "Ist (eingeplant)", F_TEXT_G)
            zelle(ws, b["ist"], 4, warn, font(9, True, C_FEHLT_T))
            for w in range(1, NW + 1):
                zelle(ws, b["ist"], K_W0 + w - 1, f_ist(pw(w), '"EIC"', extra=f'*({pfix("gruppe")}="")'),
                      F_TEXT_G, bg=C_RECHNET, rand=RAHMEN, ausr=MITTE, fmt=ZAHL)
        abstand(ws, (b["diff"] + 1) if not grau else (b["ist"] + 1))

    # ---------- Total Eichhof (wie bisher "Ein-Austritt")
    band(ws, lay["total"], 1, 5, "Total Eichhof", linie_bis=letzte)
    gr = {n: b for b in lay["bloecke"] for n in [b["name"]] if b["gi"] != 9}
    for titel, geschl in (("Total", None), ("Frauen", "F"), ("Männer", "M")):
        s = lay["summen"][titel]
        bed_rows = [gr[n]["bedarf"] for n in gr if geschl is None or gr[n]["geschl"] == geschl]
        band(ws, s["kopf"], 1, 5, {"Total": "Alle Gruppen", "Frauen": "Frauen (Steckling, Fax)",
                                   "Männer": "Männer (Verteilen, Luchs)"}[titel], dunkel=False, farbe=C_KOPF, hoehe=18,
             linie_bis=letzte)
        ws.cell(s["kopf"], 1).font = F_TEXT_B
        zelle(ws, s["bedarf"], 3, "Bedarf", F_TEXT_B, rand=UNTEN_HAAR)
        zelle(ws, s["ist"], 3, "Ist (eingeplant)", F_TEXT_G, rand=UNTEN_HAAR)
        zelle(ws, s["diff"], 3, "Differenz", F_TEXT_B)
        for w in range(1, NW + 1):
            col, cl = K_W0 + w - 1, kw_k(w)
            zelle(ws, s["bedarf"], col, "=ROUND(" + "+".join(f"{cl}{x}" for x in bed_rows) + ",1)", F_TEXT_B,
                  bg=C_RECHNET, rand=RAHMEN, ausr=MITTE, fmt=ZAHL)
            zelle(ws, s["ist"], col, f_ist(pw(w), '"EIC"', geschlecht=geschl), F_TEXT_G, bg=C_RECHNET, rand=RAHMEN,
                  ausr=MITTE, fmt=ZAHL)
            zelle(ws, s["diff"], col, f"=ROUND({cl}{s['ist']}-{cl}{s['bedarf']},1)", F_TEXT, bg=C_RECHNET,
                  rand=Border(left=S_HAAR, right=S_HAAR, top=S_HAAR, bottom=S_MITTEL), ausr=MITTE, fmt=ZAHL)
        diff_farben(ws, f"{kw_k(1)}{s['diff']}:{kw_k(NW)}{s['diff']}")
        abstand(ws, s["diff"] + 1)

    # ---------- Gruppen zuteilen
    zr = lay["zuteilen"]
    band(ws, zr, 1, 5, "Gruppen zuteilen", linie_bis=None)
    zelle(ws, zr + 1, 1, "Neue Leute: in der nächsten freien Zeile den Namen wählen (die Liste zeigt nur Leute ohne Gruppe) "
                         "und die Gruppe wählen. Mit 2. Gruppe arbeitet die Person je zur Hälfte in beiden Gruppen.",
          F_UNTERTITEL)
    for c, t in ((2, "Nr"), (3, "Name"), (4, "Gruppe"), (5, "2. Gruppe (je halb)")):
        zelle(ws, zr + 2, c, t, F_KOPF, bg=C_KOPF, rand=UNTEN_MITTEL)
    zugeteilt = [p for p in personen if p.get("gruppe")]
    for i in range(K_PLAETZE_ZUTEILUNG):
        r = z1 + i
        zelle(ws, r, 2, i + 1, F_KLEIN, ausr=MITTE, rand=UNTEN_HAAR)
        for c in (3, 4, 5):
            zelle(ws, r, c, None, F_TEXT, rand=RAHMEN_EINGABE)
        if i < len(zugeteilt):
            p = zugeteilt[i]
            ws.cell(r, 3, " ".join(f"{p['vorname']} {p['nachname']}".split()))
            ws.cell(r, 4, p["gruppe"])
    d = DataValidation(type="list", formula1="=NeueEichhof", allow_blank=True, showErrorMessage=True,
                       errorStyle="stop", errorTitle="Gruppen zuteilen",
                       error="Bitte einen Namen aus der Liste wählen. Die Liste zeigt Leute bei Eichhof, die noch keine Gruppe haben.")
    d.add(f"C{z1}:C{lay['z2']}")
    ws.add_data_validation(d)
    d2 = DataValidation(type="list", formula1=f"={listen['gruppen']}", allow_blank=True, showErrorMessage=True)
    d2.add(f"D{z1}:E{lay['z2']}")
    ws.add_data_validation(d2)
    ws.conditional_formatting.add(f"C{z1}:C{lay['z2']}", FormulaRule(
        formula=[f'AND(C{z1}<>"",COUNTIF($C${z1}:$C${lay["z2"]},C{z1})>1)'], fill=fill(C_FEHLT), font=Font(color=C_FEHLT_T)))

    for col, wdt in zip("ABCDEF", [4.5, 4.5, 28, 22, 18, 3]):
        ws.column_dimensions[col].width = wdt
    ws.column_dimensions["F"].hidden = True
    for w in range(1, NW + 1):
        ws.column_dimensions[kw_k(w)].width = 5.4
    ws.freeze_panes = f"{kw_k(1)}5"
    return ws


# ============================================================ Gemüse Freiland
G_W0 = 3                    # Spalte C = KW 1
G_VERSATZ = 1               # Zeilen des Original-Blatts + 1


def kw_g(w):
    return CL(G_W0 + w - 1)


def blatt_gemuese(wb, zeilen_freiland):
    ws = wb.create_sheet("Gemüse Freiland")
    blatt_grundlage(ws, f'="Arbeitsplanung "&{JAHR_REF}&" Freiland"',
                    "Stunden pro Tätigkeit und Woche eintragen. Soll = Summe Stunden / 52. "
                    "Ist = eingeplante Personen aus der Planung.", "70AD47")
    legende(ws, 1, G_W0 + 2, [("Soll", FARBE["FRE"], font(9, True)), ("es fehlen Leute", C_FEHLT, font(9, True, C_FEHLT_T)),
                              ("mehr Leute als nötig", C_ZUVIEL, font(9, color=C_ZUVIEL_T))], breite=4)
    zelle(ws, 3, 1, "Monat", F_GRUPPE)
    zelle(ws, 4, 1, "Tätigkeit", F_KOPF, bg=C_KOPF, rand=UNTEN_MITTEL)
    zelle(ws, 4, 2, "Notiz", F_KOPF, bg=C_KOPF, rand=UNTEN_MITTEL)
    kw_kopf(ws, 4, 3, G_W0, oben="monat")
    summe_col = G_W0 + NW + 1
    zelle(ws, 4, summe_col, "Summe", F_KOPF, bg=C_KOPF, rand=UNTEN_MITTEL, ausr=MITTE)
    ws.row_dimensions[4].height = 22
    abstand(ws, 5)
    erste, letzte_t = 5 + G_VERSATZ, 44 + G_VERSATZ
    for z in zeilen_freiland:
        r = z["zeile"] + G_VERSATZ
        if not z["label"] and not any(v is not None for v in z["werte"]):
            abstand(ws, r, 8)
            continue
        bg = FARBE["FRE"] if z["gruen"] else None
        zelle(ws, r, 1, z["label"] or None, F_TEXT, bg=bg, rand=UNTEN_HAAR)
        zelle(ws, r, 2, z["notiz"] if z["notiz"] not in (None, "") else None, F_KLEIN, bg=bg, rand=UNTEN_HAAR)
        for w, v in enumerate(z["werte"], start=1):
            zelle(ws, r, G_W0 + w - 1, v, F_TEXT, rand=RAHMEN, ausr=MITTE)
        zelle(ws, r, summe_col, f"=SUM({kw_g(1)}{r}:{kw_g(NW)}{r})", F_TEXT_G, bg=C_RECHNET, rand=RAHMEN,
              ausr=MITTE, fmt="0")
    r_sum, r_soll, r_ist, r_diff, r_jaet = letzte_t + 2, letzte_t + 3, letzte_t + 4, letzte_t + 5, letzte_t + 7
    abstand(ws, letzte_t + 1, 10)
    zelle(ws, r_sum, 1, "Summe Wochenstunden", F_TEXT_G)
    zelle(ws, r_soll, 1, "Soll: Anzahl Personen", F_TEXT_B, bg=FARBE["FRE"])
    zelle(ws, r_soll, 2, f'="Stunden / "&{STD_REF}', F_KLEIN, bg=FARBE["FRE"])
    zelle(ws, r_ist, 1, "Ist: eingeplante Personen", F_TEXT_B)
    zelle(ws, r_ist, 2, "aus der Planung", F_KLEIN)
    zelle(ws, r_diff, 1, "Differenz (Ist − Soll)", F_TEXT_B)
    zelle(ws, r_jaet, 1, "Jätstunden (automatisierbar)", F_TEXT_G)
    jaet_rows = [8, 17, 21, 25, 28]
    for w in range(1, NW + 1):
        col, cl = G_W0 + w - 1, kw_g(w)
        zelle(ws, r_sum, col, f"=SUM({cl}{erste}:{cl}{letzte_t})", F_TEXT_G, bg=C_RECHNET, rand=RAHMEN, ausr=MITTE, fmt="0")
        zelle(ws, r_soll, col, f"=ROUND({cl}{r_sum}/{STD_REF},1)", F_TEXT_B, bg=FARBE["FRE"], rand=RAHMEN, ausr=MITTE, fmt=ZAHL)
        zelle(ws, r_ist, col, f_ist(pw(w), '"FRE"'), F_TEXT_B, bg=C_RECHNET, rand=RAHMEN, ausr=MITTE, fmt=ZAHL)
        zelle(ws, r_diff, col, f"=ROUND({cl}{r_ist}-{cl}{r_soll},1)", F_TEXT, bg=C_RECHNET,
              rand=Border(left=S_HAAR, right=S_HAAR, top=S_HAAR, bottom=S_MITTEL), ausr=MITTE, fmt=ZAHL)
        zelle(ws, r_jaet, col, "=" + "+".join(f"{cl}{x + G_VERSATZ}" for x in jaet_rows), F_TEXT_G, bg=C_RECHNET,
              rand=RAHMEN, ausr=MITTE, fmt="0")
    diff_farben(ws, f"{kw_g(1)}{r_diff}:{kw_g(NW)}{r_diff}")
    for rr in (r_sum, r_soll, r_jaet):
        zelle(ws, rr, summe_col, f"=SUM({kw_g(1)}{rr}:{kw_g(NW)}{rr})", F_TEXT_G, bg=C_RECHNET, rand=RAHMEN, ausr=MITTE, fmt="0")
    abstand(ws, r_diff + 1, 10)

    # Grafik: liegt rechts der fixierten Spalten (sonst verschiebt sie sich beim Scrollen)
    band(ws, r_jaet + 2, G_W0, G_W0 + 10, "Grafik: Ist und Soll pro Woche")
    kat = Reference(ws, min_col=G_W0, max_col=G_W0 + NW - 1, min_row=4)
    ch = kombi_grafik("Freiland – Personen pro Woche",
                      Reference(ws, min_col=G_W0, max_col=G_W0 + NW - 1, min_row=r_ist),
                      Reference(ws, min_col=G_W0, max_col=G_W0 + NW - 1, min_row=r_soll), kat,
                      farbe_ist="A9D08E", farbe_soll="375623", name_ist="Ist (eingeplant)", name_soll="Soll")
    ws.add_chart(ch, f"{kw_g(1)}{r_jaet + 4}")

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 14
    for w in range(1, NW + 1):
        ws.column_dimensions[kw_g(w)].width = 5.4
    ws.column_dimensions[CL(G_W0 + NW)].width = 2
    ws.column_dimensions[CL(summe_col)].width = 9
    ws.freeze_panes = f"{kw_g(1)}5"
    return ws, {"soll": r_soll}


# ============================================================ Übersicht
U_W0 = 5  # Spalte E = KW 1


def kw_u(w):
    return CL(U_W0 + w - 1)


def runde(x, schritt=0.5):
    return round(x / schritt) * schritt


def aufteilen(total, f_ist, m_ist):
    """Startwert: SOLL nach dem Frauen-Anteil der Planung aufteilen (pro Woche)."""
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
    blatt_grundlage(ws, None, None, "1F3864")
    zelle(ws, 1, 2, f'="Übersicht "&{JAHR_REF}', F_TITEL)
    zelle(ws, 2, 2, "IST = eingeplant gemäss Planung · SOLL = Bedarf · Knöpfe 1 / 2 / 3 oben links: Frauen und Männer "
                    "ein- und ausblenden · Grafiken: Blatt «Grafik»", F_UNTERTITEL)
    ws.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=True)
    legende(ws, 1, U_W0 + 18, [("es fehlen Leute", C_FEHLT, font(9, True, C_FEHLT_T)),
                               ("mehr Leute als nötig", C_ZUVIEL, font(9, color=C_ZUVIEL_T)),
                               ("Eingabe", "FFFFFF", font(9)), ("rechnet Excel", C_RECHNET, font(9, color=C_GRAU_D))],
            breite=4)
    zelle(ws, 4, 2, "Bereich", F_KOPF, bg=C_KOPF, rand=UNTEN_MITTEL)
    zelle(ws, 4, 3, "Beschreibung", F_KOPF, bg=C_KOPF, rand=UNTEN_MITTEL)
    kw_kopf(ws, 4, 3, U_W0, oben="monat", fmt='"KW "0')
    ws.row_dimensions[4].height = 22
    abstand(ws, 5)
    letzte = U_W0 + NW - 1

    def zeile(r, label, text, werte=None, formel=None, eingabe=False, total=False, farbe=None, diff=False,
              unten=False, einzug=False):
        zelle(ws, r, 2, ("   " if einzug else "") + label, F_TEXT_B if total else (F_TEXT_G if einzug else F_TEXT),
              bg=farbe, rand=UNTEN_HAAR)
        zelle(ws, r, 3, text, F_KLEIN if einzug or not total else F_TEXT_G, bg=farbe, rand=UNTEN_HAAR)
        for w in range(1, NW + 1):
            cell = ws.cell(r, U_W0 + w - 1)
            if formel:
                f = formel(w)
                cell.value = f"=ROUND({f[1:]},1)" if not f.startswith("=COUNT") else f
                cell.fill = fill(C_RECHNET)
                cell.border = RAHMEN
            else:
                cell.value = werte[w - 1]
                cell.border = RAHMEN_EINGABE
            cell.number_format, cell.alignment = ZAHL, MITTE
            cell.font = F_TEXT_B if total else (F_TEXT_G if einzug else F_TEXT)
        if diff:
            diff_farben(ws, f"{kw_u(1)}{r}:{kw_u(NW)}{r}")
        if unten:
            for col in range(2, letzte + 1):
                c = ws.cell(r, col)
                c.border = Border(left=c.border.left, right=c.border.right, top=c.border.top, bottom=S_MITTEL)

    def gruppiere(r1, r2, level=1):
        for rr in range(r1, r2 + 1):
            ws.row_dimensions[rr].outline_level = level

    R = {"vertraege": 7, "ist": 8, "ist_f": 9, "ist_m": 10, "soll": 11, "soll_f": 12, "soll_m": 13,
         "diff": 14, "diff_f": 15, "diff_m": 16, "abw": 17, "reserve": 18, "diff_res": 19, "vj": 20}
    band(ws, 6, 2, 3, "Gesamtbetrieb", linie_bis=letzte)

    bereiche = [
        ("BLU", "Blumen", "eingabe_fm", "Summe Erststempelung Vorjahr – Startwert, bitte anpassen"),
        ("FRE", "Freiland", "freiland", "aus Blatt Gemüse Freiland (Stunden / 52)"),
        ("LOG", "Logistik", "logistik", "fixe Posten gemäss Logistik-Leitung"),
        ("FEL", "Kräuter Feld", "eingabe_fm", "Minimum Teamgrösse gemäss Teamleiter Feld"),
        ("TOM", "Tomaten", "eingabe_fm", "Planzahlen Tomaten"),
        ("EIC", "Kräuter Eichhof", "eichhof", "aus Blatt Kräuter Eichhof (Bedarf je Gruppe)"),
        ("JUC", "Kräuter Juch", "eingabe_fm", "Erststempelung Vorjahr / Plan Juch – Startwert, bitte anpassen"),
    ]
    r = 29
    blocks = {}
    for code, name, art, quelle in bereiche:
        farbe = FARBE[code]
        f_ist26, m_ist26 = q.ist_nach_geschlecht(personen, code)
        ws.cell(r, 1, code)
        band(ws, r, 2, 3, name, dunkel=False, farbe=farbe, linie_bis=letzte)
        code_ref = f"$A${r}"
        b = {"kopf": r}
        r += 1
        zeile(r, "IST", "eingeplant gemäss Planung", formel=lambda w: f_ist(pw(w), code_ref), total=True, farbe=farbe)
        zeile(r + 1, "Frauen", "", formel=lambda w: f_ist(pw(w), code_ref, "F"), einzug=True, farbe=farbe)
        zeile(r + 2, "Männer", "", formel=lambda w: f_ist(pw(w), code_ref, "M"), einzug=True, farbe=farbe)
        b["ist"], b["ist_f"], b["ist_m"] = r, r + 1, r + 2
        gruppiere(r + 1, r + 2)
        r += 3
        b["soll"], b["soll_f"], b["soll_m"] = r, r + 1, r + 2
        if art == "eingabe_fm":
            if code == "FEL":
                f_start, m_start = [0] * NW, list(soll["FEL"])
            else:
                f_start, m_start = aufteilen(soll[code], f_ist26, m_ist26)
            zeile(r, "SOLL", quelle, formel=lambda w: f"={kw_u(w)}{r + 1}+{kw_u(w)}{r + 2}", total=True, farbe=farbe)
            zeile(r + 1, "Frauen", "Eingabe", werte=f_start, eingabe=True, einzug=True, farbe=farbe)
            zeile(r + 2, "Männer", "Eingabe", werte=m_start, eingabe=True, einzug=True, farbe=farbe)
            gruppiere(r + 1, r + 2)
            r += 3
        elif art == "freiland":
            zeile(r, "SOLL", quelle, formel=lambda w: f"='Gemüse Freiland'!{kw_g(w)}{gem['soll']}", total=True, farbe=farbe)
            zeile(r + 1, "Frauen", "Eingabe: fix angestellte Frauen", werte=[round(x, 1) for x in f_ist26],
                  eingabe=True, einzug=True, farbe=farbe)
            zeile(r + 2, "Männer", "SOLL minus Frauen", formel=lambda w: f"=MAX(0,{kw_u(w)}{r}-{kw_u(w)}{r + 1})",
                  einzug=True, farbe=farbe)
            gruppiere(r + 1, r + 2)
            r += 3
        elif art == "logistik":
            details = list(soll["LOG"].items())
            n = len(details)
            zeile(r, "SOLL", quelle, formel=lambda w: f"={kw_u(w)}{r + 1}+{kw_u(w)}{r + 2}", total=True, farbe=farbe)
            zeile(r + 1, "Frauen", "Eingabe", werte=[0] * NW, eingabe=True, einzug=True, farbe=farbe)
            zeile(r + 2, "Männer", "Summe der Posten", formel=lambda w: f"=SUM({kw_u(w)}{r + 3}:{kw_u(w)}{r + 2 + n})",
                  einzug=True, farbe=farbe)
            for i, (lab, werte) in enumerate(details):
                zeile(r + 3 + i, "     " + lab, "Eingabe", werte=werte, eingabe=True, einzug=True, farbe=farbe)
            gruppiere(r + 1, r + 2)
            gruppiere(r + 3, r + 2 + n, level=2)
            r += 3 + n
        elif art == "eichhof":
            k = "'Kräuter Eichhof'!"
            zeile(r, "SOLL", quelle, formel=lambda w: f"={k}{kw_k(w)}{kr_summen['Total']['bedarf']}", total=True, farbe=farbe)
            zeile(r + 1, "Frauen", "Gruppen Steckling, Fax", formel=lambda w: f"={k}{kw_k(w)}{kr_summen['Frauen']['bedarf']}",
                  einzug=True, farbe=farbe)
            zeile(r + 2, "Männer", "Gruppen Verteilen, Luchs", formel=lambda w: f"={k}{kw_k(w)}{kr_summen['Männer']['bedarf']}",
                  einzug=True, farbe=farbe)
            gruppiere(r + 1, r + 2)
            r += 3
        b["diff"], b["diff_f"], b["diff_m"] = r, r + 1, r + 2
        zeile(r, "IST - SOLL", "minus = es fehlen Leute", formel=lambda w: f"={kw_u(w)}{b['ist']}-{kw_u(w)}{b['soll']}",
              total=True, farbe=farbe, diff=True)
        zeile(r + 1, "Frauen", "", formel=lambda w: f"={kw_u(w)}{b['ist_f']}-{kw_u(w)}{b['soll_f']}", einzug=True,
              farbe=farbe, diff=True)
        zeile(r + 2, "Männer", "", formel=lambda w: f"={kw_u(w)}{b['ist_m']}-{kw_u(w)}{b['soll_m']}", einzug=True,
              farbe=farbe, diff=True, unten=True)
        gruppiere(r + 1, r + 2)
        abstand(ws, r + 3, 12)
        r += 4
        blocks[code] = b

    def summe(key):
        return lambda w: "=" + "+".join(f"{kw_u(w)}{blocks[c][key]}" for c in blocks)

    wr = lambda w: prng(pw(w))  # noqa: E731
    zeile(R["vertraege"], "Verträge", "Personen mit Eintrag in dieser Woche (inkl. Abwesende)",
          formel=lambda w: f"=COUNTA({wr(w)})", total=True)
    zeile(R["ist"], "IST", "eingeplant (ohne Ferien / Krank / Unfall)", formel=summe("ist"), total=True)
    zeile(R["ist_f"], "Frauen", "", formel=summe("ist_f"), einzug=True)
    zeile(R["ist_m"], "Männer", "", formel=summe("ist_m"), einzug=True)
    zeile(R["soll"], "SOLL", "Bedarf (Summe Bereiche)", formel=summe("soll"), total=True)
    zeile(R["soll_f"], "Frauen", "", formel=summe("soll_f"), einzug=True)
    zeile(R["soll_m"], "Männer", "", formel=summe("soll_m"), einzug=True)
    zeile(R["diff"], "IST - SOLL", "minus = es fehlen Leute", formel=summe("diff"), total=True, diff=True)
    zeile(R["diff_f"], "Frauen", "", formel=summe("diff_f"), einzug=True, diff=True)
    zeile(R["diff_m"], "Männer", "", formel=summe("diff_m"), einzug=True, diff=True)
    zeile(R["abw"], "Abwesend", "Ferien, Krank, Unfall gemäss Planung",
          formel=lambda w: f'=COUNTIF({wr(w)},"F")+COUNTIF({wr(w)},"K")+COUNTIF({wr(w)},"U")')
    zeile(R["reserve"], "Reserve Krank", "Durchschnitt Krank 2023–2025 (Eingabe)",
          werte=[round(x, 1) for x in soll["KRANK"]], eingabe=True)
    zeile(R["diff_res"], "IST - SOLL", "inkl. Reserve Krank", formel=lambda w: f"={kw_u(w)}{R['diff']}-{kw_u(w)}{R['reserve']}",
          total=True, diff=True)
    zeile(R["vj"], "VJ", "Verträge Vorjahr gemäss Mobatime (Eingabe)", werte=[round(x, 1) for x in soll["VJ"]],
          eingabe=True, unten=True)
    for a, bb in ((R["ist_f"], R["ist_m"]), (R["soll_f"], R["soll_m"]), (R["diff_f"], R["diff_m"])):
        gruppiere(a, bb)
    abstand(ws, 21, 12)

    # ---------- Wen anstellen?
    band(ws, 22, 2, 3, "Wen anstellen?")
    zelle(ws, 23, 2, "von KW", F_TEXT)
    zelle(ws, 24, 2, "bis KW", F_TEXT)
    for rr, v in ((23, 23), (24, 35)):
        zelle(ws, rr, 3, v, F_TEXT_B, rand=RAHMEN_EINGABE, ausr=LINKS)
    kw_rng = f"${kw_u(1)}$4:${kw_u(NW)}$4"

    def check(r_diff):
        d = f"${kw_u(1)}${r_diff}:${kw_u(NW)}${r_diff}"
        n = f'COUNTIFS({d},"<-0.5",{kw_rng},">="&$C$23,{kw_rng},"<="&$C$24)'
        mx = f"-SUMPRODUCT(MIN({d}+(({kw_rng}<$C$23)+({kw_rng}>$C$24))*1000))"
        return n, mx

    nf, mf = check(R["diff_f"])
    nm, mm = check(R["diff_m"])
    wochen = "($C$24-$C$23+1)"
    zelle(ws, 25, 2, "Frauen", F_TEXT)
    zelle(ws, 25, 3, f'=IF({nf}=0,"genug Frauen eingeplant","in "&{nf}&" von "&{wochen}&" Wochen fehlen Frauen (höchstens "&ROUND({mf},1)&")")',
          F_TEXT, bg=C_RECHNET, rand=RAHMEN)
    zelle(ws, 26, 2, "Männer", F_TEXT)
    zelle(ws, 26, 3, f'=IF({nm}=0,"genug Männer eingeplant","in "&{nm}&" von "&{wochen}&" Wochen fehlen Männer (höchstens "&ROUND({mm},1)&")")',
          F_TEXT, bg=C_RECHNET, rand=RAHMEN)
    zelle(ws, 27, 2, "Empfehlung", F_TEXT_B)
    zelle(ws, 27, 3, f'=IF(AND({nf}>0,{nm}>0),"Frauen und Männer fehlen – ein Paar passt",'
                     f'IF({nm}>0,"Es fehlen Männer – eher Männer anstellen",'
                     f'IF({nf}>0,"Es fehlen Frauen – eher Frauen anstellen","In diesem Zeitraum fehlt niemand")))',
          F_TEXT_B, bg=C_RECHNET, rand=RAHMEN)
    abstand(ws, 28, 12)

    ws.column_dimensions["A"].hidden = True
    ws.column_dimensions["B"].width = 18
    ws.column_dimensions["C"].width = 52
    ws.column_dimensions["D"].width = 2
    for w in range(1, NW + 1):
        ws.column_dimensions[kw_u(w)].width = 6.2
    ws.freeze_panes = f"{kw_u(1)}5"
    zeilen = {"Gesamtbetrieb": R} | {name: blocks[code] for code, name, *_x in bereiche}
    return ws, zeilen


# ============================================================ Grafik
def blatt_grafik(wb, u_zeilen):
    ws = wb.create_sheet("Grafik")
    blatt_grundlage(ws, f'="Grafik "&{JAHR_REF}',
                    "Bereich wählen. Säulen = eingeplant (IST), Linie = Bedarf (SOLL). "
                    "Die Zahlen kommen aus dem Blatt Übersicht.", "1F3864")
    zelle(ws, 4, 1, "Bereich", F_TEXT_B, ausr=LINKS)
    zelle(ws, 4, 2, "Gesamtbetrieb", F_ABSCHNITT, rand=RAHMEN_EINGABE, ausr=LINKS)
    ws.row_dimensions[4].height = 22
    namen = list(u_zeilen.keys())
    d = DataValidation(type="list", formula1='"' + ",".join(namen) + '"', allow_blank=False, showErrorMessage=True)
    d.add("B4")
    ws.add_data_validation(d)

    # Daten für die Grafiken (unten auf dem Blatt)
    d0 = 64
    band(ws, d0 - 2, 1, 3, "Daten für die Grafiken (rechnet Excel)", dunkel=False, farbe=C_KOPF)
    tab0 = d0 + 9
    zelle(ws, tab0 - 1, 1, "Zeilen im Blatt Übersicht", F_GRUPPE)
    spalten = ["ist", "soll", "ist_f", "soll_f", "ist_m", "soll_m"]
    for i, (name, z) in enumerate(u_zeilen.items()):
        zelle(ws, tab0 + i, 1, name, F_KLEIN)
        for j, key in enumerate(spalten):
            zelle(ws, tab0 + i, 2 + j, z[key], F_KLEIN)
    tab = f"$A${tab0}:$G${tab0 + len(u_zeilen) - 1}"
    zeilen = [("KW", None), ("IST Frauen", 4), ("SOLL Frauen", 5), ("IST Männer", 6), ("SOLL Männer", 7),
              ("IST Total", 2), ("SOLL Total", 3)]
    for i, (lab, k) in enumerate(zeilen):
        r = d0 + i
        zelle(ws, r, 1, lab, F_KLEIN)
        for w in range(1, NW + 1):
            c = 3 + w - 1
            if k is None:
                zelle(ws, r, c, w, F_KLEIN, ausr=MITTE)
            else:
                ucol = kw_u(w)
                zelle(ws, r, c, f"=INDEX('Übersicht'!{ucol}$1:{ucol}$400,VLOOKUP($B$4,{tab},{k},FALSE))", F_KLEIN,
                      ausr=MITTE, fmt=ZAHL)
    kat = Reference(ws, min_col=3, max_col=2 + NW, min_row=d0)

    def ref(i):
        return Reference(ws, min_col=3, max_col=2 + NW, min_row=d0 + i)

    for nr, (t, ist_i, soll_i, farbe) in enumerate((("Frauen", 1, 2, "C9B8E0"), ("Männer", 3, 4, C_IST_BALKEN),
                                                    ("Total", 5, 6, "BFBFBF"))):
        ch = kombi_grafik(f"{t} – eingeplant und Bedarf pro Woche", ref(ist_i), ref(soll_i), kat, farbe_ist=farbe,
                          breite=30, hoehe=8.2)
        ws.add_chart(ch, f"A{6 + nr * 19}")
    ws.column_dimensions["A"].width = 16
    ws.column_dimensions["B"].width = 22
    for w in range(1, NW + 1):
        ws.column_dimensions[CL(2 + w)].width = 4.2
    return ws


# ============================================================ Anleitung
ANLEITUNG = [
    ("Personalplanung – Anleitung", "titel"),
    ("Ein gemeinsames File für die Personalplanung. Jede Person hat ihr eigenes Blatt.", "unter"),
    ("Grundregeln", "abschnitt"),
    ("Jeder schreibt nur in sein Blatt: Personalchefin = Planung und Übersicht · Kräuter-Leitung = Kräuter Eichhof · "
     "Gemüse-Leitung = Gemüse Freiland.", None),
    ("Weisse Felder mit Rahmen werden von Hand ausgefüllt. Grau hinterlegte Felder rechnet Excel – dort nichts eintippen.", None),
    ("Nichts nach rechts ziehen. Alle Wochen des Jahres sind schon vorbereitet.", None),
    ("Das File liegt auf dem Netzlaufwerk: nach dem Eintragen speichern und schliessen, damit die anderen schreiben können.", None),
    ("Planung (Personalchefin)", "abschnitt"),
    ("Eine Zeile pro Person mit Hauptbereich, Vertrag und Notizen. Neue Person: in der nächsten leeren Zeile eintragen.", None),
    ("Pro Woche aus der Liste wählen: EIC Eichhof · JUC Juch · FEL Feld · FRE Freiland · TOM Tomaten · BLU Blumen · "
     "LOG Logistik · F Ferien · K Krank · U Unfall (immer ganze Wochen).", None),
    ("Zwei Bereiche in derselben Woche: das Paar aus der Liste wählen, z.B. BLU/FRE. Die Person zählt dann bei "
     "beiden Bereichen je halb.", None),
    ("Mehrere Wochen auf einmal: Wochen markieren, im ersten Feld wählen oder tippen, dann Ctrl+Enter.", None),
    ("Die Spalten Gruppe Eichhof und 2. Gruppe füllt Excel aus dem Blatt Kräuter Eichhof.", None),
    ("Übersicht (Personalchefin)", "abschnitt"),
    ("Pro Bereich: IST (eingeplant), SOLL (Bedarf) und IST - SOLL, jeweils für Frauen und Männer. "
     "Rot = es fehlen Leute, gelb = mehr Leute eingeplant als nötig.", None),
    ("Mit den Knöpfen 1 / 2 / 3 oben links die Zeilen Frauen / Männer ein- und ausblenden.", None),
    ("SOLL für Blumen, Tomaten, Kräuter Feld, Juch und Logistik wird hier eingetragen. Eichhof und Freiland kommen "
     "aus ihren Blättern.", None),
    ("Wen anstellen?: Zeitraum eingeben – Excel zeigt, ob Frauen, Männer oder beide fehlen.", None),
    ("Grafik", "abschnitt"),
    ("Oben den Bereich wählen. Drei Grafiken: Frauen, Männer und Total – Säulen = eingeplant, Linie = Bedarf.", None),
    ("Kräuter Eichhof (Kräuter-Leitung)", "abschnitt"),
    ("Oben stehen neue Leute, die noch keiner Gruppe zugeteilt sind. Ganz unten bei «Gruppen zuteilen» den Namen "
     "aus der Liste wählen und die Gruppe wählen – fertig.", None),
    ("Die Namen erscheinen danach automatisch in ihrer Gruppe. x = ganze Woche da, 0.5 = halbe Woche, "
     "F / K / U = abwesend, leer = in dieser Woche in einem anderen Bereich.", None),
    ("Pro Gruppe die Zeile Bedarf ausfüllen. Ist und Differenz rechnet Excel.", None),
    ("Gemüse Freiland (Gemüse-Leitung)", "abschnitt"),
    ("Wie bisher: Stunden pro Tätigkeit und Woche. Soll = Summe Stunden / 52. Darunter Ist (eingeplant) und Differenz.", None),
    ("Neues Jahr", "abschnitt"),
    ("File kopieren, im Blatt Listen das Jahr ändern, in der Planung die Wochen leeren. Der Bedarf des alten Jahres "
     "ist der Startwert für das neue.", None),
]


def blatt_anleitung(wb):
    ws = wb.create_sheet("Anleitung")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "A6A6A6"
    r = 1
    for t, art in ANLEITUNG:
        if art == "titel":
            zelle(ws, r, 2, t, F_TITEL)
            ws.row_dimensions[r].height = 28
        elif art == "unter":
            zelle(ws, r, 2, t, F_UNTERTITEL)
            r += 1
        elif art == "abschnitt":
            r += 1
            zelle(ws, r, 2, t, F_ABSCHNITT, rand=UNTEN_DUNKEL)
            ws.row_dimensions[r].height = 22
        else:
            zelle(ws, r, 2, "•  " + t, font(10.5), ausr=Alignment(wrap_text=True, vertical="top", indent=1))
            ws.row_dimensions[r].height = 30 if len(t) > 115 else 17
        r += 1
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 120
    return ws


# ============================================================ main
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
    lay = kraeuter_layout(personen)
    blatt_planung(wb, personen, listen, k_tab_refs(lay))
    blatt_kraeuter(wb, personen, bedarf, listen, lay)
    _gm, gem = blatt_gemuese(wb, freiland)
    _ub, u_zeilen = blatt_uebersicht(wb, personen, soll, lay["summen"], gem)
    blatt_grafik(wb, u_zeilen)
    blatt_anleitung(wb)
    wb._sheets = [wb[n] for n in ("Planung", "Übersicht", "Grafik", "Kräuter Eichhof", "Gemüse Freiland",
                                  "Anleitung", "Listen")]
    wb.active = 0
    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
    if a.vorschau:
        bereiche = {"Planung": "A1:AJ34", "Übersicht": "B1:AF60", "Grafik": "A1:AE62",
                    "Kräuter Eichhof": f"A{lay['total'] - 3}:AE{lay['z1'] + 12}", "Gemüse Freiland": "A1:AH90", "Anleitung": "A1:B40",
                    "Listen": "A1:N45"}
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
