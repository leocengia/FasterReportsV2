#!/usr/bin/env python3
"""Estrae i 4 fogli DATASET di un workbook chiuso come CSV.

A cosa serve: rendere il collaudo di una settimana chiusa **riproducibile senza
dipendere dagli export originali**. I fogli DATASET del workbook *sono* i dati
che erano stati incollati, quindi ri-esportandoli si ottengono gli input di
quella settimana e si puo' far girare la pipeline sugli stessi numeri per
confrontare il risultato.

    python tools/export_datasets_csv.py "samples/omni-report/Omni Report W30.xlsm" -o input/

Legge in streaming: `AT_DATASET` e' un XML da 30 MB e non ha senso tenerlo tutto
in memoria.

LIMITE DA TENERE PRESENTE. Questi CSV hanno i dati giusti ma **non il formato
degli export veri**: qui le date si scrivono in ISO e i decimali col punto,
mentre l'export vero potrebbe usare `gg/mm/aaaa` e la virgola. Servono a
verificare che il *motore* produca gli stessi numeri, non che il parser regga il
formato reale — per quello serve un export vero.
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
import zipfile
from datetime import timedelta
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from fasterreports.core.coerce import EXCEL_EPOCH  # noqa: E402
from fasterreports.core.contract import col_to_index, index_to_col, load_contract  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]

_ROW = re.compile(rb'<row [^>]*r="(\d+)"[^>]*>(.*?)</row>', re.S)
_CELL = re.compile(rb'<c r="([A-Z]+)\d+"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
_V = re.compile(rb"<v>(.*?)</v>", re.S)
_IS = re.compile(rb"<is>.*?<t[^>]*>(.*?)</t>", re.S)
_T = re.compile(rb't="([^"]+)"')


def _unescape(s: str) -> str:
    return (
        s.replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
        .replace("&apos;", "'").replace("&amp;", "&")
    )


def load_shared(z: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    data = z.read("xl/sharedStrings.xml").decode("utf8", "replace")
    return [
        _unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", m.group(1), re.S)))
        for m in re.finditer(r"<si>(.*?)</si>", data, re.S)
    ]


def sheet_path(z: zipfile.ZipFile, name: str) -> str:
    wb = z.read("xl/workbook.xml").decode("utf8", "replace")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8", "replace")
    relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    for nm, rid in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wb):
        if _unescape(nm) == name:
            t = relmap[rid].lstrip("/")
            return t if t.startswith("xl/") else "xl/" + t
    raise SystemExit(f"Foglio {name!r} non trovato.")


def iter_rows(raw: bytes, shared: list[str]):
    """(numero_riga, {colonna: valore}) in streaming."""
    for rm in _ROW.finditer(raw):
        rownum = int(rm.group(1))
        cells: dict[str, str] = {}
        for col, attrs, inner in _CELL.findall(rm.group(2)):
            if inner is None:
                continue
            t = _T.search(attrs)
            kind = t.group(1).decode() if t else ""
            v = _V.search(inner)
            if kind == "s" and v:
                idx = int(v.group(1))
                val = shared[idx] if idx < len(shared) else ""
            elif kind == "inlineStr":
                m = _IS.search(inner)
                val = _unescape(m.group(1).decode("utf8", "replace")) if m else ""
            elif v:
                val = _unescape(v.group(1).decode("utf8", "replace"))
            else:
                continue
            cells[col.decode()] = val
        if cells:
            yield rownum, cells


def to_iso(value: str) -> str:
    """Seriale Excel -> ISO. Se non e' un numero si lascia com'e'."""
    try:
        serial = float(value)
    except ValueError:
        return value
    dt = EXCEL_EPOCH + timedelta(days=serial)
    if dt.hour or dt.minute or dt.second:
        return dt.strftime("%Y-%m-%d %H:%M:%S")
    return dt.strftime("%Y-%m-%d")


def export(workbook: Path, dataset, out_dir: Path, filename: str) -> tuple[Path, int, int]:
    with zipfile.ZipFile(workbook) as z:
        shared = load_shared(z)
        raw = z.read(sheet_path(z, dataset.sheet))

    start = dataset.start_index
    end = col_to_index(dataset.data_end_col) if dataset.data_end_col else dataset.last_input_index
    derived = {f.target_index for f in dataset.derived_fields}
    cols = [index_to_col(i) for i in range(start, end + 1) if i not in derived]

    # Le colonne datetime del contratto vanno riscritte come date leggibili:
    # nel foglio sono seriali, e un CSV con "46223.26" al posto di una data
    # sarebbe un export che nessuno ha mai visto.
    date_cols = {
        f.target_col for f in dataset.input_fields if f.dtype == "datetime"
    }

    # La colonna chiave dice quali righe sono dati veri: il foglio e' spesso
    # dimensionato molto oltre (AT_DATASET arriva a riga 130000 per le formule
    # P/Q, ma i dati sono 26541).
    key = dataset.input_fields[0].target_col

    out_path = out_dir / filename
    written = 0
    header: list[str] = []
    with out_path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.writer(fh)
        for rownum, cells in iter_rows(raw, shared):
            if rownum == dataset.header_row:
                header = [cells.get(c, "") for c in cols]
                writer.writerow(header)
                continue
            if rownum < dataset.header_row or not cells.get(key, "").strip():
                continue
            row = []
            for c in cols:
                v = cells.get(c, "")
                row.append(to_iso(v) if c in date_cols and v else v)
            writer.writerow(row)
            written += 1
    return out_path, len(header), written


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("workbook", type=Path)
    ap.add_argument("-o", "--out", type=Path, default=ROOT / "input")
    ap.add_argument("--config", type=Path, default=ROOT / "config")
    ap.add_argument(
        "--datasets", nargs="*",
        default=["AT_DATASET", "ATwi_DATASET", "SF_DATABASE", "PSAT_DATASET"],
    )
    args = ap.parse_args(argv)

    if not args.workbook.is_file():
        print(f"File non trovato: {args.workbook}", file=sys.stderr)
        return 2
    contract = load_contract(args.config / "columns.yml")
    args.out.mkdir(parents=True, exist_ok=True)

    import yaml

    settings = yaml.safe_load((args.config / "settings.yml").read_text(encoding="utf-8"))
    names = settings.get("input_files") or {}

    print(f"Sorgente: {args.workbook.name}\nDestinazione: {args.out}\n")
    for ds_name in args.datasets:
        dataset = contract.dataset(ds_name)
        filename = names.get(ds_name) or f"{ds_name}.csv"
        path, n_cols, n_rows = export(args.workbook, dataset, args.out, filename)
        size = path.stat().st_size / 1_048_576
        print(f"  {filename:12} {n_rows:>7} righe · {n_cols:>3} colonne · {size:5.1f} MB")

    print(
        "\nNOTA: questi CSV hanno i dati giusti ma non il formato degli export\n"
        "veri (qui date in ISO e decimali col punto). Servono a verificare che il\n"
        "motore produca gli stessi numeri, non che il parser regga il formato\n"
        "reale: per quello serve un export vero."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
