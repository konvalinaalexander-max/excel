#!/usr/bin/env python3
"""Erzeugt die Daten-Datei "personalplanung-daten.json" mit den echten Daten aus den Excel-Dateien.

Die Datei enthält Personendaten. Sie gehört auf das Netzlaufwerk und NICHT ins Repository.

Aufruf:
  python3 build_daten.py --gesamt G.xlsx --kraeuter K.xlsx --freiland F.xlsx --out personalplanung-daten.json
"""
import argparse
import datetime
import json
import re

import openpyxl

import quellen as q

GRUPPEN_CODE = {"Steckling": "STK", "Fax 10.5er": "F10", "Fax 13er": "F13",
                "Verteilen / Topfmaschine": "VER", "Luchs": "LUC"}
HAUPT_CODE = {"Kräuter Eichhof": "EIC", "Kräuter Juch": "JUC", "Kräuter Feld": "FEL",
              "Gemüse Freiland": "FRE", "Tomaten": "TOM", "Blumen": "BLU", "Logistik": "LOG"}
JAHRE = [{"y": 2026, "n": 53, "off": 0}, {"y": 2027, "n": 52, "off": 53}]
T = 105
IMPORT_INFO = [
    "Bedarf nur als Gesamtzahl: In den Originaldateien steht der Bedarf aller Bereiche nur als Anzahl Personen pro Woche, "
    "nicht nach Frauen und Männern. «Bedarf Frauen/Männer» bleibt deshalb leer, bis ihr ihn unter Übersicht → Bedarf eingeben erfasst.",
    "Kräuter Eichhof: Bedarf pro Gruppe aus der Kräuter-Datei (2026 KW 1–52, 2027 KW 1–8). Eine KW 53 gibt es dort nicht. "
    "Die Gesamtplanung hat für Eichhof «Frauen/Männer Plan» nur in KW 1–8 und 49–53; diese Zahlen passen nicht zum Gruppenbedarf und sind nicht übernommen.",
    "Gemüse Freiland: Bedarf = Summe Wochenstunden ÷ 52 wie in der Datei Arbeitsplanung Freiland 2026 (KW 1–52). "
    "KW 53 und 2027 sind dort nicht geplant. Jätstunden wie in der Original-Formel (Salat, Fenchel, Kürbis, Zwiebel, Karotten).",
    "Kräuter Juch ab KW 27: Zahlen aus dem Blatt «korrg. Zahlen» (Summe Plan, Stand 13.08.2026). In der Übersicht standen dort noch "
    "Vorjahreswerte und ab KW 33 der Wert 50.",
    "Logistik: Zeile «SOLL» der Übersicht. Blumen, Tomaten, Kräuter Feld: Zeile «SOLL» der Übersicht.",
    "2027: aus der Gesamtplanung nur KW 1–12 (dort bewusst verlängert). Ab KW 13 sind die Zahlen im Original durch Weiterziehen "
    "entstanden (sie steigen jede Woche um 1) und bleiben leer. Freiland und Feld 2027 sind schon ab KW 1 so entstanden und bleiben leer.",
    "Leere Zellen in einer Bedarfszeile zählen als 0 – so rechnet auch das Original seine Summen. Wochen, die im Original gar nicht "
    "geplant sind, bleiben leer und erscheinen als «–» (nicht erfasst).",
    "Eingeplant wird neu aus dem Wochenplan gerechnet. Im Original zählt die Formel bei Blumen in KW 52 und 53 die Wochen 2 und 3.",
    "Reserve Krank: «Krank Durchschnitt 2023–2025» aus der Übersicht, nur für 2026.",
]


def iso_t(datum):
    """Kalenderwoche eines Datums als Index t über beide Jahre (oder None)."""
    y, w, _ = datum.isocalendar()
    for j in JAHRE:
        if j["y"] == y and 1 <= w <= j["n"]:
            return j["off"] + w - 1
    return None


def zusatz_gesamt(wb):
    ws = wb["Planung"]
    hdr = [c.value for c in ws[2]]
    ix = {h: i for i, h in enumerate(hdr)}
    out = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        if not row[ix["Vorname"]]:
            continue
        w27 = []
        for k in range(1, 53):
            v = row[ix[f"27_{k:02d}"]]
            v = "" if v is None else str(v).strip().upper()
            w27.append(q.CODE_MAP.get(v, v))

        def dt(x):
            return x.date().isoformat() if isinstance(x, datetime.datetime) else None

        out.append({"w27": w27, "ein": dt(row[ix["Eintritt"]]), "aus": dt(row[ix["Austritt"]]),
                    "fs": str(row[ix["FSchein"]] or ""), "zu": str(row[ix["Zustand"]] or ""),
                    "badge": row[ix["Badge"]]})
    return out


FUNKTION_GRUPPE = [("luchs", "LUC"), ("13er", "F13"), ("10.5", "F10"), ("10,5", "F10"), ("steckling", "STK"),
                   ("topf", "VER"), ("verteil", "VER")]   # nur eindeutige Stichworte


def gruppen_streng(kraeuter, personen):
    """Gruppe Eichhof nur aus eindeutigen Angaben: Funktion der Person oder ein genau passender, eindeutiger
    Name in der Kräuter-Datei. Die Quelle wird gespeichert (gq), Namenszuordnungen gelten als «zu prüfen»."""
    ws = openpyxl.load_workbook(kraeuter, data_only=True)["Gruppen Kostenstelle 2026"]
    kandidaten = [p for p in personen if p["hb"] == "EIC" or any("EIC" in (c or "") for c in p["w"])]
    treffer = {}
    for gname, _g, _rb, r_pers in q.GRUPPEN:
        code = GRUPPEN_CODE[gname]
        for r in r_pers:
            text = str(ws.cell(r, 5).value or "").strip()
            if not text:
                continue
            for teil in re.split(r"/", text):
                tok = [t for t in q.norm(teil.replace("(", " ").replace(")", " "))
                       if t not in ("mann", "fahrer", "ab", "kw") and len(t) > 1]
                if not tok:
                    continue
                passt = [p for p in kandidaten if set(tok) <= set(q.norm(p["vn"]) + q.norm(p["nn"]))]
                if len(passt) == 1:
                    treffer.setdefault(passt[0]["id"], set()).add((code, teil.strip()))
    n_f = n_n = 0
    for p in personen:
        p["g"], p["gq"] = "", ""
        f = (p["fn"] or "").lower()
        fg = {g for kw, g in FUNKTION_GRUPPE if kw in f}
        if len(fg) == 1:
            p["g"], p["gq"] = fg.pop(), f"Funktion: «{p['fn']}»"
            n_f += 1
            continue
        t = treffer.get(p["id"], set())
        gruppen = {c for c, _ in t}
        if len(gruppen) == 1:
            code = gruppen.pop()
            namen = sorted({n for c, n in t})
            p["g"], p["gq"] = code, "Kräuter-Datei: «" + "», «".join(namen) + "» bei " + [g for g, c in GRUPPEN_CODE.items() if c == code][0]
            n_n += 1
    return n_f, n_n


def finde_person(personen, text):
    toks = set(re.findall(r"[a-zà-ÿőű]+", str(text or "").lower()))
    if not toks:
        return None
    best = None
    for p in personen:
        pt = set(re.findall(r"[a-zà-ÿőű]+", f"{p['vn']} {p['nn']}".lower()))
        if pt and pt <= toks | pt and len(pt & toks) >= 2:
            if best is None or len(pt & toks) > len(set(re.findall(r"[a-zà-ÿőű]+", f"{best['vn']} {best['nn']}".lower())) & toks):
                best = p
    return best["id"] if best else None


def mutationen(wb, personen):
    ws = wb["Mutationen"]
    out = []
    for i, r in enumerate(ws.iter_rows(min_row=3, values_only=True)):
        if not any(r):
            continue
        kw = r[0] if isinstance(r[0], int) else None
        name = " ".join(str(x).strip() for x in (r[2], r[3]) if x)
        bem = " · ".join(str(x).strip() for x in (r[5], r[6]) if x)
        buero = r[8].date().isoformat() if isinstance(r[8], datetime.datetime) else (str(r[8]) if r[8] else "")
        kst = str(r[4] or "").strip()
        area = kst.upper() if kst.upper() in HAUPT_CODE.values() else HAUPT_CODE.get(kst, "")
        out.append({"id": i + 1, "t": (kw - 1) if kw and 1 <= kw <= 53 else None, "typ": str(r[1] or "").strip(),
                    "name": name, "pid": finde_person(personen, name), "area": area, "bem": bem,
                    "von": str(r[7] or "").strip(), "buero": buero})
    return out


def notizen(wb):
    ws = wb["Notizen"]
    out = {}
    spalten = {2: "FRE", 3: "EIC", 4: "BLU"}
    for r in ws.iter_rows(min_row=4, values_only=True):
        datum = r[1]
        if not isinstance(datum, datetime.datetime):
            continue
        t = iso_t(datum.date())
        if t is None:
            continue
        for c, area in spalten.items():
            v = r[c]
            if v not in (None, "") and str(v).strip():
                eintrag = f"{datum:%d.%m.}: {str(v).strip()}"
                alt = out.setdefault(str(t), {}).get(area, "")
                out[str(t)][area] = (alt + "\n" if alt else "") + eintrag
    return out


def traktanden(wb):
    ws = wb["Traktanden"]
    out = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if r[1]:
            out.append({"nr": str(r[0] or "").strip(), "titel": str(r[1]).strip(), "mod": str(r[2] or "").strip(),
                        "min": str(r[3] or "").strip()})
    return out


def aufgaben(pfad):
    ws = openpyxl.load_workbook(pfad, data_only=True)["Aufgaben Liste"]
    rows = []
    for r in range(2, ws.max_row + 1):
        v = [ws.cell(r, c).value for c in range(1, 15)]
        if not any(v):
            continue
        rows.append({"bereich": str(v[0] or ""), "haupt": str(v[1] or ""), "aufgabe": str(v[2] or "").strip(),
                     "wann": str(v[3] or "").strip(), "sommer": str(v[4] or "").strip(),
                     "winter": " ".join(str(x).strip() for x in v[5:7] if x),
                     "tage": [str(x or "").strip() for x in v[7:12]],
                     "stv1": str(v[12] or "").strip(), "stv2": str(v[13] or "").strip()})
    return rows


def daten(gesamt, kraeuter, freiland):
    personen_q, soll = q.lese_gesamt(gesamt)
    frei = q.lese_freiland(freiland)
    wb = openpyxl.load_workbook(gesamt, data_only=True)
    zusatz = zusatz_gesamt(wb)

    partner_map = {}
    personen = []
    for i, p in enumerate(personen_q):
        z = zusatz[i]
        pid = p["partner"]
        if pid is not None:
            partner_map.setdefault(pid, len(partner_map) + 1)
        personen.append({
            "id": i + 1, "badge": "" if z["badge"] is None else str(z["badge"]),
            "vn": " ".join(str(p["vorname"] or "").split()), "nn": " ".join(str(p["nachname"] or "").split()),
            "fm": p["fm"] or "",
            "pa": partner_map.get(pid) if pid is not None else None,
            "hb": HAUPT_CODE.get(p["bereich"], ""), "fs": z["fs"], "zu": z["zu"],
            "ein": z["ein"], "aus": z["aus"], "fn": p["funktion"], "bem": p["bemerkung"],
            "w": list(p["wochen"]) + z["w27"],
        })

    ub = wb["Übersicht"]
    leer27 = [None] * 52

    n_f, n_n = gruppen_streng(kraeuter, personen)
    print(f"Gruppen Eichhof: {n_f} aus der Funktion, {n_n} über den Namen in der Kräuter-Datei (zu prüfen)")

    def zahl(x):
        return round(x, 2) if isinstance(x, (int, float)) else 0

    def reihe(row, kw27=12, werte26=None):
        """2026 KW 1–53 aus der Übersicht (leere Zelle = 0), 2027 nur KW 1–kw27, danach leer (None)."""
        v26 = [zahl(x) for x in werte26] if werte26 is not None else [zahl(ub.cell(row, 5 + k).value) for k in range(53)]
        v27 = [zahl(ub.cell(row, 5 + 54 + k).value) for k in range(kw27)] + [None] * (52 - kw27)
        return v26 + v27

    assert ub.cell(2, 5 + 54).value == "27_01"
    nichts = lambda: [None] * T
    bed = {
        "BLU": {"T": reihe(22), "F": nichts(), "M": nichts()},
        "TOM": {"T": reihe(53), "F": nichts(), "M": nichts()},
        "JUC": {"T": reihe(68, werte26=soll["JUC"]), "F": nichts(), "M": nichts()},
        "LOG": {"T": reihe(38), "F": nichts(), "M": nichts()},
        "FEL": {"T": reihe(47, kw27=0), "F": nichts(), "M": nichts()},
        "EIC": {"F": nichts(), "M": nichts()},
        "FRE": {"F": nichts(), "M": nichts()},
    }

    # Kräuter: Bedarf pro Gruppe, 2026 KW 1–52 und die 2027-Wochen, die in der Datei stehen
    kws = openpyxl.load_workbook(kraeuter, data_only=True)["Gruppen Kostenstelle 2026"]
    kopf = {c: kws.cell(1, c).value for c in range(1, kws.max_column + 1)}
    c26 = 18
    assert kopf[c26] == 1 and kopf[c26 + 51] == 52
    c27 = next(c for c in range(c26 + 52, kws.max_column + 1) if kopf.get(c) == 1)
    n27 = 0
    while kopf.get(c27 + n27) == n27 + 1:
        n27 += 1
    bedarf_eic = {}
    for gname, _g, r_bed, _r in q.GRUPPEN:
        v26 = [zahl(kws.cell(r_bed, c26 + k).value) for k in range(52)] + [None]
        v27 = [zahl(kws.cell(r_bed, c27 + k).value) for k in range(n27)] + [None] * (52 - n27)
        bedarf_eic[GRUPPEN_CODE[gname]] = v26 + v27

    # Freiland: Stunden KW 1–52; Jätstunden-Zeilen aus der Original-Formel
    fws = openpyxl.load_workbook(freiland)["Freiland 2024"]
    jaet_formel = str(fws.cell(49, 3).value or "")
    jaet_zeilen = {int(x) for x in re.findall(r"C(\d+)", jaet_formel)}
    assert jaet_zeilen, jaet_formel
    stunden = []
    for z in frei:
        if not z["label"] and not any(v is not None for v in z["werte"]):
            stunden.append({"abstand": True})
            continue
        v26 = [v if v is not None else 0 for v in z["werte"][:52]] + [None]
        stunden.append({"t": z["label"], "n": str(z["notiz"] or ""), "gr": z["gruen"], "jaet": z["zeile"] in jaet_zeilen,
                        "v": v26 + leer27})
    jetzt = datetime.datetime.now().isoformat(timespec="seconds")
    data = {
        "schema": 2, "jahre": JAHRE, "T": T, "personen": personen, "gw": {},
        "bedarfEIC": bedarf_eic, "bed": bed, "stunden": stunden, "std": 52,
        "vj": [round(x, 1) for x in soll["VJ"]] + leer27, "krank": [round(x, 2) for x in soll["KRANK"]] + leer27,
        "notes": {}, "wnotes": notizen(wb), "sitzung": {}, "mut": mutationen(wb, personen), "bew": [],
        "trakt": traktanden(wb), "anfragen": [], "aufgaben": aufgaben(kraeuter),
        "log": [{"t": jetzt, "r": "System", "u": "", "x": "Daten aus den Excel-Dateien übernommen (Planung 2026 und 2027)"}],
        "nextId": len(personen) + 1, "nextReq": 1, "nextMut": 10000, "nextBew": 1,
        "importInfo": IMPORT_INFO,
    }
    return {"format": "personalplanung", "version": 1, "gespeichert": jetzt, "von": "Import", "data": data}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--gesamt", required=True)
    ap.add_argument("--kraeuter", required=True)
    ap.add_argument("--freiland", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    obj = daten(a.gesamt, a.kraeuter, a.freiland)
    with open(a.out, "w", encoding="utf8") as f:
        json.dump(obj, f, ensure_ascii=False, separators=(",", ":"))
    d = obj["data"]
    print(f"gespeichert: {a.out} · {len(d['personen'])} Personen · {len(d['mut'])} Mutationen · "
          f"{len(d['wnotes'])} Wochen mit Notizen · {len(d['aufgaben'])} Aufgaben")


if __name__ == "__main__":
    main()
