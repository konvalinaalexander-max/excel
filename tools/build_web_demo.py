#!/usr/bin/env python3
"""Erzeugt den Web-Prototyp "Personalplanung" (eine HTML-Datei) aus den bestehenden Excel-Dateien.

Die echte Planungsstruktur 2026 wird übernommen (Bereiche pro Woche, Frauen/Männer, Paare,
Gruppen, Bedarf), die Namen werden durch erfundene Namen ersetzt. Badge, Funktion und
Bemerkungen werden nicht übernommen.

Aufruf:
  python3 build_web_demo.py --gesamt G.xlsx --kraeuter K.xlsx --freiland F.xlsx \
      --vorlage web/personalplanung.html --out personalplanung.html
"""
import argparse
import json
import random

import quellen as q
from build_personalplanung import aufteilen

GRUPPEN_CODE = {"Steckling": "STK", "Fax 10.5er": "F10", "Fax 13er": "F13",
                "Verteilen / Topfmaschine": "VER", "Luchs": "LUC"}
HAUPT_CODE = {"Kräuter Eichhof": "EIC", "Kräuter Juch": "JUC", "Kräuter Feld": "FEL",
              "Gemüse Freiland": "FRE", "Tomaten": "TOM", "Blumen": "BLU", "Logistik": "LOG"}

VORNAMEN_F = ["Ildikó", "Erzsébet", "Katalin", "Zsófia", "Réka", "Noémi", "Enikő", "Tímea", "Boglárka",
              "Orsolya", "Hajnalka", "Gabriella", "Krisztina", "Mónika", "Edit", "Piroska", "Kinga", "Nóra",
              "Ioana", "Mihaela", "Daniela", "Raluca", "Agnieszka", "Magdalena", "Joanna", "Ewelina",
              "Katarzyna", "Dorota", "Beatriz", "Inês", "Mariana", "Rita", "Sofia", "Teresa", "Luísa",
              "Andrea", "Melinda", "Emese", "Csilla", "Brigitta", "Henrietta", "Viktória", "Szilvia",
              "Judit", "Anikó", "Erika", "Rozália", "Margit", "Ilona", "Klára"]
VORNAMEN_M = ["Levente", "Botond", "Zsolt", "Gergely", "Ákos", "Bence", "Csaba", "Attila", "Lóránd",
              "Szabolcs", "Tamás", "Imre", "Ferenc", "Gábor", "Huba", "Endre", "Kristóf", "Barnabás",
              "Andrei", "Mihai", "Florin", "Radu", "Tomasz", "Piotr", "Marek", "Paweł", "Jakub",
              "Rui", "Tiago", "Nuno", "Miguel", "Joaquim", "Hugo", "Dávid", "Márton", "Örs", "Vilmos",
              "Zoltán", "Sándor", "István", "Árpád", "Géza", "Lajos", "Norbert", "Ottó", "Róbert"]
NACHNAMEN = ["Kovács", "Szabó", "Tóth", "Farkas", "Balogh", "Molnár", "Lakatos", "Takács", "Juhász",
             "Mészáros", "Simon", "Fekete", "Szilágyi", "Török", "Fehér", "Pál", "Vass", "Boros",
             "Hegedűs", "Bíró", "Kiss", "Antal", "Orbán", "Sipos", "Kelemen", "Dobos", "Gál", "Varró",
             "Popescu", "Ionescu", "Munteanu", "Rusu", "Nowak", "Wójcik", "Kamińska", "Lewandowska",
             "Zając", "Silva", "Ferreira", "Pereira", "Costa", "Rodrigues", "Almeida", "Sousa",
             "Csiki", "Benedek", "Kozma", "Ábrahám", "Bartha", "Demeter", "Fazakas", "Gyarmati",
             "Incze", "Jakab", "Kacsó", "Lukács", "Márkus", "Nagy-Bodó", "Pásztor", "Rácz", "Sebestyén"]


def demo_daten(gesamt, kraeuter, freiland):
    rnd = random.Random(2026)
    personen, soll = q.lese_gesamt(gesamt)
    bedarf, kr_namen = q.lese_kraeuter(kraeuter)
    q.gruppen_zuordnung(personen, kr_namen)
    frei = q.lese_freiland(freiland)

    # Partner neu nummerieren, gleiche Nachnamen für Paare
    partner_map, partner_nach = {}, {}
    vergeben = set()
    out = []
    for i, p in enumerate(personen):
        fm = p["fm"] or rnd.choice("FM")
        while True:
            vn = rnd.choice(VORNAMEN_F if fm == "F" else VORNAMEN_M)
            pid = p["partner"]
            if pid is not None and pid in partner_nach:
                nn = partner_nach[pid]
            else:
                nn = rnd.choice(NACHNAMEN)
            if (vn, nn) not in vergeben:
                break
        vergeben.add((vn, nn))
        if pid is not None:
            partner_nach.setdefault(pid, nn)
            partner_map.setdefault(pid, len(partner_map) + 1)
        wochen = list(p["wochen"])
        belegt = [w for w, c in enumerate(wochen) if c]
        out.append({
            "id": i + 1, "vn": vn, "nn": nn, "fm": fm,
            "pa": partner_map.get(pid) if pid is not None else None,
            "hb": HAUPT_CODE.get(p["bereich"], ""),
            "von": (belegt[0] + 1) if belegt else None, "bis": (belegt[-1] + 1) if belegt else None,
            "g": GRUPPEN_CODE.get(p.get("gruppe") or "", ""),
            "w": wochen,
        })
    # ein paar Krankheitswochen als Beispiel
    kandidaten = [p for p in out if sum(1 for c in p["w"][38:46] if c and c not in ("F",)) >= 6]
    for p in rnd.sample(kandidaten, min(8, len(kandidaten))):
        w = rnd.randint(40, 45)
        if p["w"][w - 1] and p["w"][w - 1] != "F":
            p["w"][w - 1] = "K"

    # Bedarf
    soll_bereiche = {}
    for code in ("BLU", "TOM", "JUC"):
        f_ist, m_ist = q.ist_nach_geschlecht(personen, code)
        f, m = aufteilen(soll[code], f_ist, m_ist)
        soll_bereiche[code] = {"F": f, "M": m}
    soll_bereiche["FEL"] = {"F": [0] * q.NW, "M": [round(x, 1) for x in soll["FEL"]]}
    log = [round(sum(v[i] for v in soll["LOG"].values()), 1) for i in range(q.NW)]
    soll_bereiche["LOG"] = {"F": [0] * q.NW, "M": log}
    f_ist, _m = q.ist_nach_geschlecht(personen, "FRE")
    soll_bereiche["FRE"] = {"F": [round(x, 1) for x in f_ist]}

    stunden = []
    for z in frei:
        if not z["label"] and not any(v is not None for v in z["werte"]):
            stunden.append({"t": "", "abstand": True})
            continue
        stunden.append({"t": z["label"], "n": z["notiz"] or "", "gr": z["gruen"],
                        "v": [v if v is not None else 0 for v in z["werte"]]})
    return {
        "jahr": 2026, "nw": q.NW, "personen": out,
        "bedarfEIC": {GRUPPEN_CODE[n]: [round(x, 1) for x in werte] for n, werte in bedarf.items()},
        "soll": soll_bereiche, "stunden": stunden, "stdProPerson": 52,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--gesamt", required=True)
    ap.add_argument("--kraeuter", required=True)
    ap.add_argument("--freiland", required=True)
    ap.add_argument("--vorlage", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    daten = demo_daten(a.gesamt, a.kraeuter, a.freiland)
    html = open(a.vorlage, encoding="utf8").read()
    js = json.dumps(daten, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    assert "__DEMO_DATEN__" in html
    open(a.out, "w", encoding="utf8").write(html.replace("__DEMO_DATEN__", js))
    print("gespeichert:", a.out, f"({len(daten['personen'])} Personen)")


if __name__ == "__main__":
    main()
