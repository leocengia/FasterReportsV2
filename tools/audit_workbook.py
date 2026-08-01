#!/usr/bin/env python3
"""Ispezione di un workbook Excel senza aprire Excel — e senza openpyxl.

Legge direttamente l'XML dentro il file .xlsm/.xlsx. Non carica il foglio in
memoria come oggetto, quindi funziona anche sul WOW da 60 MB dove openpyxl in
scrittura va in OOM (contesto WOW §5).

Tre usi:

  --headers        intestazioni di riga 1 dei fogli indicati (o di tutti)
  --usage          quali colonne dei dataset sono consumate da quali formule
                   -> e' il test di auto-verifica del contratto (piano §11)
  --tables         i ListObject e se il loro intervallo copre i dati

  python tools/audit_workbook.py "samples/Omni Report W30.xlsm" --all
  python tools/audit_workbook.py FILE --usage --datasets AT_DATASET SF_DATABASE
  python tools/audit_workbook.py FILE --headers --json > tests/fixtures/headers.json

Nota sul matching dei nomi foglio nelle formule: `AT_DATASET` e' una
sottostringa di `PSAT_DATASET`. Senza un confine a sinistra si attribuiscono ad
AT tutte le colonne di PSAT — errore facile e silenzioso.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from collections import defaultdict
from pathlib import Path

DEFAULT_DATASETS = ["AT_DATASET", "ATwi_DATASET", "SF_DATABASE", "PSAT_DATASET"]


def col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


class Workbook:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.z = zipfile.ZipFile(path)
        wb = self.z.read("xl/workbook.xml").decode("utf8", "replace")
        rels = self.z.read("xl/_rels/workbook.xml.rels").decode("utf8", "replace")
        relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
        self.sheets: dict[str, str] = {}
        self.hidden: set[str] = set()
        for m in re.finditer(r"<sheet ([^>]*)/>", wb):
            attrs = m.group(1)
            name = re.search(r'name="([^"]*)"', attrs)
            rid = re.search(r'r:id="(rId\d+)"', attrs)
            if not (name and rid):
                continue
            target = relmap.get(rid.group(1), "").lstrip("/")
            if target and not target.startswith("xl/"):
                target = "xl/" + target
            self.sheets[name.group(1)] = target
            if 'state="hidden"' in attrs or 'state="veryHidden"' in attrs:
                self.hidden.add(name.group(1))
        self._shared: list[str] | None = None

    @property
    def shared(self) -> list[str]:
        if self._shared is None:
            self._shared = []
            if "xl/sharedStrings.xml" in self.z.namelist():
                data = self.z.read("xl/sharedStrings.xml").decode("utf8", "replace")
                for si in re.finditer(r"<si>(.*?)</si>", data, re.S):
                    self._shared.append(
                        "".join(re.findall(r"<t[^>]*>(.*?)</t>", si.group(1), re.S))
                    )
        return self._shared

    def xml(self, sheet: str) -> str:
        return self.z.read(self.sheets[sheet]).decode("utf8", "replace")

    def dimension(self, sheet: str) -> str:
        m = re.search(r'<dimension ref="([^"]+)"', self.xml(sheet))
        return m.group(1) if m else "?"

    def header_row(self, sheet: str, row: int = 1) -> dict[str, str]:
        raw = self.xml(sheet)
        m = re.search(rf'<row r="{row}"[^>]*>(.*?)</row>', raw, re.S)
        if not m:
            return {}
        out: dict[str, str] = {}
        for ref, attrs, body in re.findall(
            r'<c r="([A-Z]+\d+)"([^>]*)>(.*?)</c>', m.group(1), re.S
        ):
            t = re.search(r't="([^"]+)"', attrs)
            v = re.search(r"<v>(.*?)</v>", body, re.S)
            inline = re.search(r"<is>.*?<t[^>]*>(.*?)</t>", body, re.S)
            if t and t.group(1) == "s" and v:
                val = self.shared[int(v.group(1))]
            elif inline:
                val = inline.group(1)
            elif v:
                val = v.group(1)
            else:
                continue
            out[re.match(r"[A-Z]+", ref).group(0)] = _unescape(val)
        return out

    def formula_usage(self, datasets: list[str]) -> dict[str, dict[str, set[str]]]:
        usage: dict[str, dict[str, set[str]]] = {d: defaultdict(set) for d in datasets}
        # Due insidie, entrambe scoperte col sangue su questo workbook:
        #
        # 1. I nomi di foglio con spazi, nelle formule, sono fra apici:
        #    `'Slot Only Cases'!$A$2`. Un pattern che pretende `Cases!` non li
        #    vede, e il foglio risulta "letto da nessuna formula" — falso.
        # 2. Un nome di foglio puo' essere sottostringa di un altro:
        #    `AT_DATASET` sta dentro `PSAT_DATASET`, `Turni` dentro
        #    `Helper Turni`. Serve un confine a sinistra che escluda anche uno
        #    spazio, altrimenti si attribuiscono a uno le colonne dell'altro.
        patterns = {
            d: re.compile(
                r"(?<![A-Za-z0-9_ ])" + re.escape(d) + r"'?!\$?([A-Z]{1,3})\$?\d*"
            )
            for d in datasets
        }
        for name in self.sheets:
            raw = self.xml(name)
            for f in re.findall(r"<f[^>]*>(.*?)</f>", raw, re.S):
                for d, pat in patterns.items():
                    for m in pat.finditer(f):
                        usage[d][m.group(1)].add(name)
        return usage

    def tables(self) -> list[dict]:
        out = []
        owner: dict[str, str] = {}
        for name, target in self.sheets.items():
            rel = target.replace("worksheets/", "worksheets/_rels/") + ".rels"
            if rel not in self.z.namelist():
                continue
            x = self.z.read(rel).decode("utf8", "replace")
            for t in re.findall(r'Target="([^"]*tables/[^"]+)"', x):
                owner[Path(t).name] = name
        for n in sorted(p for p in self.z.namelist() if "/tables/" in p):
            x = self.z.read(n).decode("utf8", "replace")
            m = re.search(r'<table[^>]*name="([^"]+)"[^>]*ref="([^"]+)"', x)
            if not m:
                continue
            sheet = owner.get(Path(n).name, "?")
            out.append(
                {
                    "name": m.group(1),
                    "ref": m.group(2),
                    "sheet": sheet,
                    "columns": len(re.findall(r"<tableColumn[^>]*name=", x)),
                    "sheet_dimension": self.dimension(sheet) if sheet in self.sheets else "?",
                }
            )
        return out


def _unescape(s: str) -> str:
    return (
        s.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&apos;", "'")
        .replace("&amp;", "&")
    )


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook", type=Path)
    ap.add_argument("--headers", action="store_true")
    ap.add_argument("--usage", action="store_true")
    ap.add_argument("--tables", action="store_true")
    ap.add_argument("--all", action="store_true", help="tutti i controlli")
    ap.add_argument("--sheets", nargs="*", help="limita ai fogli indicati")
    ap.add_argument("--datasets", nargs="*", default=DEFAULT_DATASETS)
    ap.add_argument("--json", action="store_true", help="output JSON (per le fixture)")
    args = ap.parse_args(argv)

    if not args.workbook.is_file():
        print(f"File non trovato: {args.workbook}", file=sys.stderr)
        return 2
    if not (args.headers or args.usage or args.tables or args.all):
        args.all = True

    wb = Workbook(args.workbook)
    payload: dict = {"workbook": args.workbook.name}

    if args.headers or args.all:
        targets = args.sheets or args.datasets
        payload["headers"] = {}
        for s in targets:
            if s not in wb.sheets:
                print(f"[skip] foglio assente: {s}", file=sys.stderr)
                continue
            hdr = wb.header_row(s)
            payload["headers"][s] = {
                "dimension": wb.dimension(s),
                "hidden": s in wb.hidden,
                "columns": hdr,
            }

    if args.usage or args.all:
        usage = wb.formula_usage(args.datasets)
        payload["usage"] = {
            d: {c: sorted(v) for c, v in sorted(cols.items(), key=lambda kv: col_index(kv[0]))}
            for d, cols in usage.items()
        }

    if args.tables or args.all:
        payload["tables"] = wb.tables()

    if args.json:
        print(json.dumps(payload, indent=2, ensure_ascii=False))
        return 0

    print(f"Workbook: {args.workbook}  ({len(wb.sheets)} fogli, {len(wb.hidden)} nascosti)")

    for sheet, info in payload.get("headers", {}).items():
        print(f"\n=== INTESTAZIONI {sheet}  (dim {info['dimension']}) ===")
        for col, val in sorted(info["columns"].items(), key=lambda kv: col_index(kv[0])):
            print(f"  {col:>3} | {val}")

    for ds, cols in payload.get("usage", {}).items():
        print(f"\n=== COLONNE DI {ds} CONSUMATE DA FORMULE ===")
        if not cols:
            print("  (nessuna)")
        for col, sheets in cols.items():
            print(f"  {col:>3} <- {', '.join(sheets)}")

    if "tables" in payload:
        print("\n=== TABELLE (ListObject) ===")
        for t in payload["tables"]:
            note = ""
            m = re.search(r"[A-Z]+(\d+)$", t["ref"])
            d = re.search(r"[A-Z]+(\d+)$", t["sheet_dimension"])
            if m and d and int(d.group(1)) > int(m.group(1)):
                note = (
                    f"  <-- ATTENZIONE: i dati del foglio arrivano a riga {d.group(1)}, "
                    f"la tabella si ferma a {m.group(1)}"
                )
            print(f"  {t['name']} su '{t['sheet']}' ref={t['ref']} ({t['columns']} col.){note}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
