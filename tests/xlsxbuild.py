"""Costruzione di .xlsx sintetici minimi, per i test dei lettori.

Serve a verificare le proprieta' dei lettori (`xlsxsource`, `tablesource`,
`wfmsource`) senza dipendere dai campioni reali — quelli servono ai golden test,
dove il risultato di ieri e' la specifica.

Due modi di scrivere il testo, perche' i lettori devono reggere entrambi:

- `inline=True` (default) usa `inlineStr`: comodo, un solo pezzo di XML;
- `inline=False` usa `sharedStrings`, che e' cio' che produce Excel davvero.

Se un test passa solo con uno dei due, il lettore ha un buco.
"""

from __future__ import annotations

import zipfile
from pathlib import Path

from fasterreports.core.xlsxsource import col_to_index

_CT_HEAD = """<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>"""
_RELS = """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

_WS_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
)
_WS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"


def esc(s: object) -> str:
    return (
        str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def _sheet_xml(grid: dict[str, object], shared: list[str] | None) -> str:
    """`grid`: {"A1": valore}. Numeri come numeri, testo come testo."""
    rows: dict[int, list[tuple[str, str]]] = {}
    for ref, val in grid.items():
        col = "".join(c for c in ref if c.isalpha())
        num = int("".join(c for c in ref if c.isdigit()))
        if isinstance(val, str) and val.startswith("="):
            # Una FORMULA, non testo: va in <f>, che e' dove la cercano gli
            # strumenti che analizzano il workbook. Scritta come testo, un test
            # sui riferimenti nelle formule sarebbe verde su un file che non ne
            # contiene nessuna.
            cell = f'<c r="{ref}"><f>{esc(val[1:])}</f></c>'
        elif isinstance(val, bool):
            cell = f'<c r="{ref}" t="b"><v>{int(val)}</v></c>'
        elif isinstance(val, (int, float)):
            cell = f'<c r="{ref}"><v>{val}</v></c>'
        elif shared is None:
            cell = f'<c r="{ref}" t="inlineStr"><is><t>{esc(val)}</t></is></c>'
        else:
            testo = str(val)
            if testo not in shared:
                shared.append(testo)
            cell = f'<c r="{ref}" t="s"><v>{shared.index(testo)}</v></c>'
        rows.setdefault(num, []).append((col, cell))

    body = ""
    for num in sorted(rows):
        celle = "".join(
            c for _, c in sorted(rows[num], key=lambda kv: col_to_index(kv[0]))
        )
        body += f'<row r="{num}">{celle}</row>'
    return (
        '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/'
        f'spreadsheetml/2006/main"><sheetData>{body}</sheetData></worksheet>'
    )


def make_workbook(
    path: Path,
    sheets: list[tuple[str, dict[str, object]]],
    *,
    inline: bool = True,
) -> Path:
    """Scrive un .xlsx con uno o piu' fogli, nell'ordine dato.

    L'ordine conta: `read_table` senza `sheet_name` prende il primo foglio, ed e'
    proprio quel comportamento che i test devono fissare.
    """
    if not sheets:
        raise ValueError("almeno un foglio")
    shared: list[str] | None = None if inline else []

    parti: list[tuple[str, str]] = []
    for i, (_nome, grid) in enumerate(sheets, start=1):
        parti.append((f"xl/worksheets/sheet{i}.xml", _sheet_xml(grid, shared)))

    ct = _CT_HEAD
    for i in range(1, len(sheets) + 1):
        ct += f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="{_WS_TYPE}"/>'
    if shared is not None:
        ct += (
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/'
            'vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
        )
    ct += "</Types>"

    rels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    wb_sheets = ""
    for i, (nome, _g) in enumerate(sheets, start=1):
        rels += f'<Relationship Id="rId{i}" Type="{_WS_REL}" Target="worksheets/sheet{i}.xml"/>'
        wb_sheets += f'<sheet name="{esc(nome)}" sheetId="{i}" r:id="rId{i}"/>'
    rels += "</Relationships>"

    wb = (
        '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/'
        'spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/'
        f'officeDocument/2006/relationships"><sheets>{wb_sheets}</sheets></workbook>'
    )

    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", _RELS)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", rels)
        for nome, xml in parti:
            z.writestr(nome, xml)
        if shared is not None:
            si = "".join(f"<si><t>{esc(s)}</t></si>" for s in shared)
            z.writestr(
                "xl/sharedStrings.xml",
                '<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/'
                f'spreadsheetml/2006/main" count="{len(shared)}" '
                f'uniqueCount="{len(shared)}">{si}</sst>',
            )
    return path


def make_xlsx(
    path: Path,
    sheet_name: str,
    grid: dict[str, object],
    *,
    inline: bool = True,
) -> Path:
    """Un solo foglio: il caso di gran lunga piu' comune nei test."""
    return make_workbook(path, [(sheet_name, grid)], inline=inline)
