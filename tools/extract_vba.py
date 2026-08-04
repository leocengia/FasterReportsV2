#!/usr/bin/env python3
"""Estrae il codice VBA di un workbook, per leggerlo senza aprire Excel.

Serve perche' il VBA fa parte del contratto: il modulo `CreaMalpractice` legge
colonne per lettera fissa (`AT_DATASET!B/F/K/P/Q`, `SF_DATABASE!BB/DY/AE`), e
quelle lettere devono comparire in config/columns.yml. Se qualcuno modifica la
macro, va riletto.

  pip install -e '.[audit]'
  python tools/extract_vba.py "samples/Omni Report W30.xlsm"
  python tools/extract_vba.py FILE --grep Cells

Nota: la decompressione dei moduli VBA segue MS-OVBA; scriversela a mano e' un
buon modo per ottenere byte plausibili e sbagliati. Qui si usa oletools.
"""

from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook", type=Path)
    ap.add_argument("--grep", help="mostra solo le righe che contengono questo testo")
    ap.add_argument(
        "--columns",
        action="store_true",
        help="riassume quali Cells(r, \"XX\") il codice legge, per foglio",
    )
    args = ap.parse_args(argv)

    try:
        from oletools.olevba import VBA_Parser
    except ImportError:
        print(
            "oletools non installato:  pip install -e '.[audit]'",
            file=sys.stderr,
        )
        return 2

    if not args.workbook.is_file():
        print(f"File non trovato: {args.workbook}", file=sys.stderr)
        return 2

    parser = VBA_Parser(str(args.workbook))
    if not parser.detect_vba_macros():
        print("Nessuna macro nel file.")
        return 0

    columns: dict[str, set[str]] = {}
    for _fname, _stream, vba_name, code in parser.extract_macros():
        # I moduli foglio contengono solo l'intestazione generata da Excel.
        if len(code.strip().splitlines()) <= 6 and "Attribute VB_Name" in code:
            continue
        if args.columns:
            for m in re.finditer(r'Cells\(\s*\w+\s*,\s*"([A-Z]{1,3})"\s*\)', code):
                columns.setdefault(vba_name, set()).add(m.group(1))
            continue
        print(f"{'#' * 20} {vba_name} {'#' * 20}")
        if args.grep:
            for i, line in enumerate(code.splitlines(), 1):
                if args.grep.lower() in line.lower():
                    print(f"{i:>5}: {line}")
        else:
            print(code)

    if args.columns:
        for mod, cols in sorted(columns.items()):
            ordered = sorted(cols, key=lambda c: (len(c), c))
            print(f"{mod}: {', '.join(ordered)}")
        print(
            "\nNota: il modulo non dice a quale foglio appartiene ciascuna lettera.\n"
            "Va letto il codice per associarle (vedi docs/audit-workbook-W30.md §4)."
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
