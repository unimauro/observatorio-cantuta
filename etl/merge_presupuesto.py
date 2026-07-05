#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Fusiona la serie histórica 2012-2024 (ca_cantuta_raw.json) en
data/presupuesto-cantuta.json SIN tocar 2025/2026 ni detalle_ultimo_anio.
"""
import json, os, datetime

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
JSON = os.path.join(ROOT, "data", "presupuesto-cantuta.json")
RAW = os.path.join(HERE, "ca_cantuta_raw.json")

CUR_YEAR = datetime.date.today().year


def main():
    with open(JSON, encoding="utf-8") as f:
        doc = json.load(f)
    with open(RAW, encoding="utf-8") as f:
        raw = json.load(f)

    serie = doc.get("serie", [])
    by_year = {int(r["year"]): r for r in serie}

    added, updated = [], []
    for ystr, rec in raw.items():
        y = int(ystr)
        if y >= 2025:  # NO tocar 2025/2026 (los pone el otro proceso)
            continue
        pim = rec["pim"]
        dev = rec["dev"]
        ejec = round(100 * dev / pim, 1) if pim else 0.0
        entry = {
            "year": y,
            "pia": rec["pia"],
            "pim": pim,
            "cert": rec["cert"],
            "dev": dev,
            "gir": rec["gir"],
            "ejec_pct": ejec,
        }
        if y == CUR_YEAR and ejec < 70:
            entry["parcial"] = True
        if y in by_year:
            updated.append(y)
        else:
            added.append(y)
        by_year[y] = entry

    doc["serie"] = [by_year[y] for y in sorted(by_year)]

    meta = doc.setdefault("_meta", {})
    meta["nota"] = ("Serie histórica 2012-2024 obtenida de Consulta Amigable MEF "
                    "(pliego 528: U.N. de Educacion Enrique Guzman y Valle - La Cantuta). "
                    "2025-2026 de fuente MEF vigente. Cifras en soles. "
                    "ejec_pct = 100*Devengado/PIM.")

    with open(JSON, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False, separators=(",", ":"))

    print("Agregados:", sorted(added))
    print("Actualizados:", sorted(updated))
    print("Serie final:")
    for r in doc["serie"]:
        p = " PARCIAL" if r.get("parcial") else ""
        print(f"  {r['year']}: PIM {r['pim']/1e6:.1f}M dev {r['dev']/1e6:.1f}M ejec {r['ejec_pct']}%{p}")


if __name__ == "__main__":
    main()
