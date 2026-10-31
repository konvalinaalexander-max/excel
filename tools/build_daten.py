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
from build_personalplanung import aufteilen

GRUPPEN_CODE = {"Steckling": "STK", "Fax 10.5er": "F10", "Fax 13er": "F13",
                "Verteilen / Topfmaschine": "VER", "Luchs": "LUC"}
HAUPT_CODE = {"Kräuter Eichhof": "EIC", "Kräuter Juch": "JUC", "Kräuter Feld": "FEL",
              "Gemüse Freiland": "FRE", "Tomaten": "TOM", "Blumen": "BLU", "Logistik": "LOG"}
JAHRE = [{"y": 2026, "n": 53, "off": 0}, {"y": 2027, "n": 52, "off": 53}]
T = 105


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
    bedarf, kr_namen = q.lese_kraeuter(kraeuter)
    q.gruppen_zuordnung(personen_q, kr_namen)
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
            "vn": p["vorname"], "nn": p["nachname"], "fm": p["fm"] or "",
            "pa": partner_map.get(pid) if pid is not None else None,
            "hb": HAUPT_CODE.get(p["bereich"], ""), "fs": z["fs"], "zu": z["zu"],
            "ein": z["ein"], "aus": z["aus"], "fn": p["funktion"], "bem": p["bemerkung"],
            "g": GRUPPEN_CODE.get(p.get("gruppe") or "", ""),
            "w": list(p["wochen"]) + z["w27"],
        })

    def zwei(werte):  # 2027 startet mit den Werten von 2026 (wie im bisherigen File)
        werte = [round(x, 1) if isinstance(x, (int, float)) else 0 for x in werte]
        return werte + werte[:52]

    soll_b = {}
    for code in ("BLU", "TOM", "JUC"):
        f_ist, m_ist = q.ist_nach_geschlecht(personen_q, code)
        f, m = aufteilen(soll[code], f_ist, m_ist)
        soll_b[code] = {"F": zwei(f), "M": zwei(m)}
    soll_b["FEL"] = {"F": [0] * T, "M": zwei(soll["FEL"])}
    log = [sum(v[i] for v in soll["LOG"].values()) for i in range(q.NW)]
    soll_b["LOG"] = {"F": [0] * T, "M": zwei(log)}
    f_ist, _m = q.ist_nach_geschlecht(personen_q, "FRE")
    soll_b["FRE"] = {"F": zwei(f_ist)}

    stunden = []
    for z in frei:
        if not z["label"] and not any(v is not None for v in z["werte"]):
            stunden.append({"abstand": True})
            continue
        stunden.append({"t": z["label"], "n": str(z["notiz"] or ""), "gr": z["gruen"],
                        "v": zwei([v if v is not None else 0 for v in z["werte"]])})
    vertraege26 = [sum(1 for p in personen if p["w"][k]) for k in range(52)]
    jetzt = datetime.datetime.now().isoformat(timespec="seconds")
    data = {
        "schema": 2, "jahre": JAHRE, "T": T, "personen": personen, "gw": {},
        "bedarfEIC": {GRUPPEN_CODE[n]: zwei(w) for n, w in bedarf.items()},
        "soll": soll_b, "stunden": stunden, "std": 52,
        "vj": [round(x, 1) for x in soll["VJ"]] + vertraege26, "krank": zwei(soll["KRANK"]),
        "notes": {}, "wnotes": notizen(wb), "sitzung": {}, "mut": mutationen(wb, personen), "bew": [],
        "trakt": traktanden(wb), "anfragen": [], "aufgaben": aufgaben(kraeuter),
        "log": [{"t": jetzt, "r": "System", "u": "", "x": "Daten aus den Excel-Dateien übernommen (Planung 2026 und 2027)"}],
        "nextId": len(personen) + 1, "nextReq": 1, "nextMut": 10000, "nextBew": 1,
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
