#!/usr/bin/env python3
"""Estrae lo storico AHT dal foglio 'AHT History' di un workbook, verso il CSV.

Serve una volta sola: le settimane gia' presenti nel template (nel WIP del
2026-08-18 erano le W22..W32, 546 righe, costruite a mano dal file
'WOW AHT Trend by CT') diventano il punto di partenza dello storico
persistente. Da lì in poi ci pensa il build.

Legge l'XML dentro il .xlsm senza aprire Excel e senza openpyxl, come
tools/audit_workbook.py — e per lo stesso motivo: qui Excel non c'e', e sul
file da 60 MB openpyxl va in out-of-memory.

  python tools/seed_aht_history.py TEMPLATE.xlsm data/aht_history.csv --anno 2026
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from audit_workbook import Workbook, _unescape  # noqa: E402

from fasterreports.core.aht_history import RigaStorico, scrivi  # noqa: E402

FOGLIO = "AHT History"


def leggi_celle(wb: Workbook, foglio: str) -> dict[int, dict[str, str]]:
    raw = wb.xml(foglio)
    out: dict[int, dict[str, str]] = defaultdict(dict)
    for m in re.finditer(r'<c r="([A-Z]+)(\d+)"([^>]*)(?:/>|>(.*?)</c>)', raw, re.S):
        col, row, attrs, body = m.group(1), int(m.group(2)), m.group(3), m.group(4) or ""
        t = re.search(r't="([^"]+)"', attrs)
        v = re.search(r"<v>(.*?)</v>", body, re.S)
        if t and t.group(1) == "s" and v:
            val = _unescape(wb.shared[int(v.group(1))])
        elif v:
            val = v.group(1)
        else:
            continue
        out[row][col] = val
    return out


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook", type=Path)
    ap.add_argument("out", type=Path)
    ap.add_argument(
        "--anno", type=int, required=True,
        help="anno ISO delle settimane presenti nel foglio (il foglio porta solo "
             "il numero di settimana, non l'anno: e' esattamente il problema che "
             "la colonna iso_year risolve da qui in avanti)",
    )
    args = ap.parse_args(argv)

    wb = Workbook(args.workbook)
    if FOGLIO not in wb.sheets:
        print(f"Il workbook non contiene il foglio {FOGLIO!r}.", file=sys.stderr)
        print(f"Fogli presenti: {', '.join(wb.sheets)}", file=sys.stderr)
        return 2

    celle = leggi_celle(wb, FOGLIO)
    righe = []
    saltate = []
    for rownum in sorted(celle):
        if rownum == 1:  # intestazione
            continue
        r = celle[rownum]
        try:
            righe.append(RigaStorico(
                iso_year=args.anno,
                week=int(float(r["A"])),
                channel=r["B"].strip(),
                case_type=r["C"].strip(),
                volume=int(float(r["D"])),
                aht=float(r.get("E") or 0.0),
            ))
        except (KeyError, TypeError, ValueError) as exc:
            saltate.append(f"riga {rownum}: {exc} ({r})")

    if saltate:
        print(f"ATTENZIONE: {len(saltate)} righe non leggibili, NON scritte:", file=sys.stderr)
        for s in saltate[:10]:
            print(f"  {s}", file=sys.stderr)
        return 1

    scrivi(args.out, righe)
    sett = sorted({r.week for r in righe})
    print(f"{len(righe)} righe scritte in {args.out}")
    print(f"settimane: {args.anno}-W{sett[0]:02d} .. {args.anno}-W{sett[-1]:02d} ({len(sett)} settimane)")
    canali = sorted({r.channel for r in righe})
    print(f"canali: {', '.join(canali)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
