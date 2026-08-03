"""Lettura di un foglio .xlsx, in sola lettura, senza openpyxl.

Perché non openpyxl: su questo progetto openpyxl è ammesso solo per letture
offline, e nemmeno lì serve. Leggere l'XML dentro lo zip costa poche righe, non
aggiunge dipendenze e regge file grossi senza caricarli come oggetti — è così
che `tools/audit_workbook.py` ispeziona il WOW da 60 MB, che in scrittura manda
openpyxl in out-of-memory.

Restituisce i **valori in cache** delle celle con formula, che è esattamente
quello che serve: `Back_Office_Time_Final.xlsx` ha 2873 formule e senza Excel
non si possono rivalutare.

  ATTENZIONE: se quel file venisse salvato senza essere ricalcolato, i valori
  in cache sarebbero vecchi. `sheet_meta()` espone `calc_chain_present` per
  segnalarlo nel preflight.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from .errors import SourceError

_CELL = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*)>(.*?)</c>', re.S)
_CELL_EMPTY = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*)/>')
_ROW = re.compile(r'<row r="(\d+)"[^>]*>(.*?)</row>', re.S)


def col_to_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def index_to_col(index: int) -> str:
    out = ""
    while index:
        index, rem = divmod(index - 1, 26)
        out = chr(65 + rem) + out
    return out


def _unescape(s: str) -> str:
    return (
        s.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&apos;", "'")
        .replace("&amp;", "&")
    )


@dataclass
class Sheet:
    name: str
    rows: dict[int, dict[str, str]] = field(default_factory=dict)
    dimension: str = ""
    n_formulas: int = 0
    calc_chain_present: bool = False

    def row(self, n: int) -> dict[str, str]:
        return self.rows.get(n, {})

    def cell(self, col: str, row: int) -> str | None:
        return self.rows.get(row, {}).get(col)

    @property
    def max_row(self) -> int:
        return max(self.rows) if self.rows else 0

    def cols_of(self, row: int) -> list[str]:
        """Colonne popolate di una riga, in ordine di foglio."""
        return sorted(self.row(row), key=col_to_index)


def read_sheet(path: str | Path, sheet_name: str | None = None) -> Sheet:
    """Legge un foglio. Senza `sheet_name` prende il primo."""
    path = Path(path)
    if not path.is_file():
        raise SourceError(f"File sorgente non trovato: {path}")
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        raise SourceError(
            f"{path.name}: non è un file .xlsx valido (l'archivio non si apre)."
        ) from None

    with z:
        try:
            wb = z.read("xl/workbook.xml").decode("utf8", "replace")
            rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8", "replace")
        except KeyError:
            raise SourceError(f"{path.name}: struttura .xlsx inattesa.") from None

        relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
        sheets: dict[str, str] = {}
        for m in re.finditer(r"<sheet ([^>]*?)/>", wb):
            attrs = m.group(1)
            nm = re.search(r'name="([^"]*)"', attrs)
            rid = re.search(r'r:id="(rId\d+)"', attrs)
            if nm and rid and rid.group(1) in relmap:
                target = relmap[rid.group(1)].lstrip("/")
                sheets[_unescape(nm.group(1))] = (
                    target if target.startswith("xl/") else "xl/" + target
                )
        if not sheets:
            raise SourceError(f"{path.name}: nessun foglio trovato.")

        if sheet_name is None:
            sheet_name = next(iter(sheets))
        if sheet_name not in sheets:
            raise SourceError(
                f"{path.name}: foglio {sheet_name!r} assente.\n"
                f"  Fogli presenti: {', '.join(repr(s) for s in sheets)}"
            )

        shared: list[str] = []
        if "xl/sharedStrings.xml" in z.namelist():
            data = z.read("xl/sharedStrings.xml").decode("utf8", "replace")
            for si in re.finditer(r"<si>(.*?)</si>", data, re.S):
                shared.append(
                    _unescape("".join(re.findall(r"<t[^>]*>(.*?)</t>", si.group(1), re.S)))
                )

        raw = z.read(sheets[sheet_name]).decode("utf8", "replace")
        calc = "xl/calcChain.xml" in z.namelist()

    out = Sheet(
        name=sheet_name,
        dimension=(re.search(r'<dimension ref="([^"]+)"', raw) or [None, ""])[1]
        if re.search(r'<dimension ref="([^"]+)"', raw)
        else "",
        n_formulas=len(re.findall(r"<f[^>]*[>/]", raw)),
        calc_chain_present=calc,
    )

    for rm in _ROW.finditer(raw):
        rownum = int(rm.group(1))
        body = rm.group(2)
        cells: dict[str, str] = {}
        for col, _r, attrs, inner in _CELL.findall(body):
            t = re.search(r't="([^"]+)"', attrs)
            v = re.search(r"<v>(.*?)</v>", inner, re.S)
            inline = re.search(r"<is>.*?<t[^>]*>(.*?)</t>", inner, re.S)
            if t and t.group(1) == "s" and v:
                # L'indice della stringa condivisa deve essere un intero: se non
                # lo e', il file e' corrotto e va detto, non ignorato. Prima
                # sollevava un ValueError nudo che il chiamante interpretava
                # come "foglio assente".
                try:
                    idx = int(v.group(1))
                except ValueError:
                    raise SourceError(
                        f"{path.name}, foglio {sheet_name!r}, cella {col}{rownum}: "
                        f"indice di stringa condivisa non valido "
                        f"({v.group(1)!r}).\n"
                        f"  Il file e' corrotto: probabilmente e' stato riscritto da "
                        f"uno strumento che non gestisce sharedStrings."
                    ) from None
                val = shared[idx] if idx < len(shared) else ""
            elif t and t.group(1) == "inlineStr" and inline:
                val = _unescape(inline.group(1))
            elif t and t.group(1) == "e" and v:
                # Errore Excel in cache (#N/D, #VALORE!): va visto, non ignorato.
                val = _unescape(v.group(1))
            elif v:
                val = _unescape(v.group(1))
            else:
                continue
            cells[col] = val
        if cells:
            out.rows[rownum] = cells
    return out


def read_sheet_names(path: str | Path) -> list[str]:
    path = Path(path)
    if not path.is_file():
        raise SourceError(f"File sorgente non trovato: {path}")
    with zipfile.ZipFile(path) as z:
        wb = z.read("xl/workbook.xml").decode("utf8", "replace")
    return [_unescape(m) for m in re.findall(r'<sheet name="([^"]*)"', wb)]
