#!/usr/bin/env python3
"""Enumera tutte le forme delle celle-turno di un roster, e prova a interpretarle.

È il controllo da rifare quando arriva un roster nuovo: se il WFM inizia a
scrivere una forma che il parser non conosce, meglio saperlo qui che scoprirlo
da ore previste sbagliate.

  python tools/audit_shift_forms.py "samples/omni-report/sorgenti/Turni_W30.xlsx"
  python tools/audit_shift_forms.py FILE --backoffice --sheet "Only Cases Shifts"

Lo "schema" riduce una cella alla sua forma: cifre -> D, lettere -> A,
sequenze di spazi -> ~. Così `0900_1331_1431_1800` e `0830_1301_1401_1730`
sono la stessa forma, e le 50000 celle si riducono a poche decine di casi da
guardare a occhio.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import Counter
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from fasterreports.core.shifts import ShiftParseError, parse_shift, parse_slot  # noqa: E402
from fasterreports.core.xlsxsource import col_to_index, read_sheet  # noqa: E402

DATE_TEXT = re.compile(r"^\d{1,2}/\d{1,2}/\d{4}$")


def schema(value: str) -> str:
    s = re.sub(r"\d", "D", str(value))
    s = re.sub(r"[^\WD_]", "A", s, flags=re.UNICODE)
    s = re.sub(r" +", "~", s)
    return s


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("workbook", type=Path)
    ap.add_argument("--sheet", default=None)
    ap.add_argument("--header-row", type=int, default=None,
                    help="riga delle intestazioni (default: 2 roster, 1 backoffice)")
    ap.add_argument("--backoffice", action="store_true",
                    help="usa il parser degli slot invece di quello dei turni")
    args = ap.parse_args(argv)

    sheet = read_sheet(args.workbook, args.sheet)
    header_row = args.header_row or (1 if args.backoffice else 2)
    print(f"Foglio {sheet.name!r}  dim={sheet.dimension}  "
          f"formule={sheet.n_formulas}  righe={len(sheet.rows)}")
    if sheet.n_formulas and not sheet.calc_chain_present:
        print("  ! il file ha formule ma nessuna calcChain: i valori in cache "
              "potrebbero essere vecchi")

    hdr = sheet.row(header_row)
    if args.backoffice:
        datecols = [c for c, v in hdr.items()
                    if re.match(r"^\d+(\.0+)?$", str(v).strip()) and float(v) > 40000]
    else:
        datecols = [c for c, v in hdr.items() if DATE_TEXT.match(str(v).strip())]
    datecols.sort(key=col_to_index)
    print(f"  colonne-data: {len(datecols)}"
          + (f"  ({datecols[0]}..{datecols[-1]})" if datecols else ""))

    forms: Counter[str] = Counter()
    example: dict[str, tuple[str, str]] = {}
    ok: Counter[str] = Counter()
    markers: Counter[str] = Counter()
    errors: list[tuple[str, str, str]] = []

    parse = parse_slot if args.backoffice else parse_shift
    for rownum in sorted(sheet.rows):
        if rownum <= header_row:
            continue
        row = sheet.row(rownum)
        for col in datecols:
            v = row.get(col)
            if v is None:
                continue
            sc = schema(v)
            forms[sc] += 1
            example.setdefault(sc, (f"{col}{rownum}", repr(v)))
            try:
                parsed = parse(v, where=f"{col}{rownum}")
            except ShiftParseError as exc:
                errors.append((f"{col}{rownum}", repr(v), exc.reason))
                continue
            ok[parsed.kind] += 1
            for m in getattr(parsed, "markers", ()):
                markers[m] += 1

    print(f"\n=== FORME DISTINTE ({len(forms)}) ===")
    for sc, n in forms.most_common():
        cell, ex = example[sc]
        print(f"  {n:>6}  {sc:<28} es. {cell}: {ex}")

    print(f"\n=== INTERPRETAZIONE ===")
    for kind, n in ok.most_common():
        print(f"  {n:>6}  {kind}")
    if markers:
        print(f"\n=== MARCATORI CONSERVATI (significato non documentato) ===")
        for m, n in markers.most_common():
            print(f"  {n:>6}  {m}")

    if errors:
        print(f"\n=== NON INTERPRETABILI ({len(errors)}) ===")
        seen: Counter[str] = Counter()
        for cell, raw, reason in errors:
            if seen[raw] == 0:
                print(f"  {cell}: {raw}  -> {reason}")
            seen[raw] += 1
        print(f"  ({len(seen)} valori distinti)")
        return 1

    print("\nTutte le celle interpretate.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
