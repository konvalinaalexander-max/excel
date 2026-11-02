#!/usr/bin/env python3
"""Baut die fertige Personalplanung als eine einzige HTML-Datei: Programm + Daten.

Die Vorlage (web/personalplanung_app.html) enthält einen leeren Datenblock
<script type="application/json" id="pp-daten">null</script>. Dieses Skript setzt die
Daten (personalplanung-daten.json oder eine ältere personalplanung.html) dort ein.

Die fertige Datei enthält Personendaten und gehört nicht ins Repository.

Aufruf:
  python3 build_html.py --vorlage web/personalplanung_app.html --daten personalplanung-daten.json --out personalplanung.html
"""
import argparse
import json
import re

MARKER = '<script type="application/json" id="pp-daten">'


def daten_lesen(pfad):
    text = open(pfad, encoding="utf8").read()
    if text.lstrip().startswith("{"):
        return json.loads(text)
    i = text.index(MARKER) + len(MARKER)
    return json.loads(text[i:text.index("</script>", i)])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--vorlage", required=True)
    ap.add_argument("--daten", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    html = open(a.vorlage, encoding="utf8").read()
    assert html.count(MARKER) == 1, "Datenblock fehlt oder ist doppelt"
    i = html.index(MARKER) + len(MARKER)
    j = html.index("</script>", i)
    assert html.index(MARKER) < html.index("<script>\n'use strict'"), "Datenblock muss vor dem Programm stehen"
    obj = daten_lesen(a.daten)
    assert obj.get("format") == "personalplanung" and isinstance(obj.get("data", {}).get("personen"), list)
    obj["data"].setdefault("log", [])
    payload = json.dumps(obj, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")
    assert not re.search(r"</script|<!--", payload, re.I)
    open(a.out, "w", encoding="utf8").write(html[:i] + payload + html[j:])
    print(f"gespeichert: {a.out} ({len(obj['data']['personen'])} Personen, Version {obj.get('version')})")


if __name__ == "__main__":
    main()
