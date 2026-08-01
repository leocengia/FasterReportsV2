"""Adattatori WFM: unpivot, ritaglio della settimana, blocchi affiancati.

Questi test costruiscono fogli .xlsx sintetici, così le proprietà si verificano
senza dipendere dai campioni — che invece servono al golden test.
"""

from __future__ import annotations

import zipfile
from datetime import date
from pathlib import Path

import pytest

from fasterreports.core.errors import SourceError
from fasterreports.core.wfmsource import (
    find_header_row,
    find_roster_blocks,
    parse_header_date,
    read_backoffice,
    read_roster,
    week_bounds,
    week_from_dates,
)
from fasterreports.core.xlsxsource import Sheet, col_to_index, read_sheet

# ---------------------------------------------------------------------------
# Costruzione di un .xlsx minimo (solo ciò che serve al nostro lettore)
# ---------------------------------------------------------------------------

_CT = """<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>"""
_RELS = """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
_WBRELS = """<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>"""


def _esc(s: str) -> str:
    return (
        str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def make_xlsx(path: Path, sheet_name: str, grid: dict[str, object]) -> Path:
    """`grid`: {"A1": valore}. Numeri come numeri, testo come inlineStr."""
    rows: dict[int, list[str]] = {}
    for ref, val in grid.items():
        col = "".join(c for c in ref if c.isalpha())
        num = int("".join(c for c in ref if c.isdigit()))
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            cell = f'<c r="{ref}"><v>{val}</v></c>'
        else:
            cell = f'<c r="{ref}" t="inlineStr"><is><t>{_esc(val)}</t></is></c>'
        rows.setdefault(num, []).append((col, cell))
    body = ""
    for num in sorted(rows):
        cells = "".join(c for _, c in sorted(rows[num], key=lambda kv: col_to_index(kv[0])))
        body += f'<row r="{num}">{cells}</row>'
    sheet = (
        '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/'
        f'spreadsheetml/2006/main"><sheetData>{body}</sheetData></worksheet>'
    )
    wb = (
        '<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/'
        'spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships"><sheets>'
        f'<sheet name="{_esc(sheet_name)}" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", _CT)
        z.writestr("_rels/.rels", _RELS)
        z.writestr("xl/workbook.xml", wb)
        z.writestr("xl/_rels/workbook.xml.rels", _WBRELS)
        z.writestr("xl/worksheets/sheet1.xml", sheet)
    return path


def roster_grid(dates=("20/07/2026", "21/07/2026"), agents=None, start_col="A"):
    """Roster a un blocco: A=Agent, B=Expected hours, C=Zona, D=Skill, E=Name, F=Surname."""
    agents = agents or [("ROSSI", "0800", "2", "HPO", "Mario", "Rossi",
                         ["0900_1300_1330_1730", "Off"])]
    grid: dict[str, object] = {}
    fixed = ["Agent", "Expected hours", "Zona", "Skill", "Name", "Surname"]
    base = col_to_index(start_col)
    from fasterreports.core.xlsxsource import index_to_col

    for i, h in enumerate(fixed):
        grid[f"{index_to_col(base + i)}2"] = h
    for j, d in enumerate(dates):
        grid[f"{index_to_col(base + len(fixed) + j)}2"] = d
    for r, (ag, eh, zona, skill, nome, cog, cells) in enumerate(agents, start=3):
        for i, v in enumerate((ag, eh, zona, skill, nome, cog)):
            grid[f"{index_to_col(base + i)}{r}"] = v
        for j, v in enumerate(cells):
            if v is not None:
                grid[f"{index_to_col(base + len(fixed) + j)}{r}"] = v
    return grid


# ---------------------------------------------------------------------------

def test_parse_header_date_testo():
    assert parse_header_date("20/07/2026") == date(2026, 7, 20)
    assert parse_header_date(" 1/1/2026 ") == date(2026, 1, 1)


def test_parse_header_date_seriale():
    assert parse_header_date(46223) == date(2026, 7, 20)
    assert parse_header_date("46223") == date(2026, 7, 20)


def test_parse_header_date_scarta_i_conteggi():
    """Un numero piccolo è un conteggio, non una data: nel back office ci sono
    colonne con 0, 10, 93."""
    assert parse_header_date(0) is None
    assert parse_header_date(107) is None
    assert parse_header_date("NO BOT") is None
    assert parse_header_date(None) is None


def test_week_bounds():
    assert week_bounds(46223) == (date(2026, 7, 20), date(2026, 7, 26))


def test_week_from_dates():
    lo, hi = week_from_dates(["2026-07-20 06:00:00", "2026-07-24 22:00:00"])
    assert (lo, hi) == (date(2026, 7, 20), date(2026, 7, 24))


def test_week_from_dates_vuoto():
    assert week_from_dates([None, ""]) is None


# --- unpivot ---------------------------------------------------------------

def test_unpivot_una_riga_per_agente_giorno(tmp_path):
    p = make_xlsx(tmp_path / "r.xlsx", "publish", roster_grid())
    src = read_roster(p, week=week_bounds(46223))
    assert src.headers[0] == "Nome agente"
    assert len(src.data) == 2  # un agente, due giorni
    lavora = [r for r in src.data if r[5] == "LAVORA"]
    assert len(lavora) == 1
    assert lavora[0][0] == "Mario Rossi"
    assert lavora[0][3] == pytest.approx(8.0)  # 09:00-17:30 meno 30' = 8h
    assert lavora[0][8] == "mario rossi"       # chiave


def test_ritaglio_della_settimana(tmp_path):
    dates = ("18/07/2026", "20/07/2026", "27/07/2026")
    grid = roster_grid(dates=dates, agents=[
        ("ROSSI", "0800", "2", "HPO", "Mario", "Rossi",
         ["0900_1700_    _    ", "0900_1700_    _    ", "0900_1700_    _    "]),
    ])
    p = make_xlsx(tmp_path / "r.xlsx", "publish", grid)
    src = read_roster(p, week=week_bounds(46223))  # 20-26/07
    assert len(src.data) == 1
    from fasterreports.core.coerce import to_datetime
    assert to_datetime(src.data[0][4]).date() == date(2026, 7, 20)


def test_filtro_skill_esatto(tmp_path):
    grid = roster_grid(agents=[
        ("A", "0800", "2", "HPO", "Uno", "Uno", ["0900_1700_    _    ", "Off"]),
        ("B", "0800", "2", "HPS", "Due", "Due", ["0900_1700_    _    ", "Off"]),
        ("C", "0800", "2", "HPO   *", "Tre", "Tre", ["0900_1700_    _    ", "Off"]),
    ])
    p = make_xlsx(tmp_path / "r.xlsx", "publish", grid)
    src = read_roster(p, week=week_bounds(46223))
    nomi = {r[0] for r in src.data}
    assert nomi == {"Uno Uno"}                       # HPS e HPO* fuori
    assert "tre tre" in src.notes.target_agents      # ma il marcato è tracciato
    assert "due due" not in src.notes.target_agents  # HPS non somiglia a HPO
    assert src.notes.skills_with_marker


def test_include_marked(tmp_path):
    grid = roster_grid(agents=[
        ("C", "0800", "2", "HPO   *", "Tre", "Tre", ["0900_1700_    _    ", "Off"]),
    ])
    p = make_xlsx(tmp_path / "r.xlsx", "publish", grid)
    src = read_roster(p, week=week_bounds(46223), include_marked=True)
    assert {r[0] for r in src.data} == {"Tre Tre"}


def test_blocchi_affiancati(tmp_path):
    """Due blocchi con popolazioni diverse, come nel file reale."""
    g1 = roster_grid(agents=[("A", "0800", "2", "HPO", "Uno", "Uno",
                             ["0900_1700_    _    ", "Off"])], start_col="A")
    g2 = roster_grid(agents=[("B", "0800", "2", "HPO", "Due", "Due",
                             ["0900_1700_    _    ", "Off"])], start_col="K")
    p = make_xlsx(tmp_path / "r.xlsx", "publish", {**g1, **g2})
    src = read_roster(p, week=week_bounds(46223))
    assert len(src.notes.blocks) == 2
    assert {r[0] for r in src.data} == {"Uno Uno", "Due Due"}


def test_agente_hpo_in_due_blocchi_e_tracciato(tmp_path):
    """Righe doppie: va segnalato, non subito."""
    g1 = roster_grid(agents=[("A", "0800", "2", "HPO", "Uno", "Uno",
                             ["0900_1700_    _    ", "Off"])], start_col="A")
    g2 = roster_grid(agents=[("A", "0800", "2", "HPO", "Uno", "Uno",
                             ["0900_1700_    _    ", "Off"])], start_col="K")
    p = make_xlsx(tmp_path / "r.xlsx", "publish", {**g1, **g2})
    src = read_roster(p, week=week_bounds(46223))
    assert len(set(src.notes.agents_in_blocks["uno uno"])) == 2
    assert len(src.data) == 4  # 2 giorni x 2 blocchi: e' il problema


def test_colonne_fisse_spostate(tmp_path):
    """Le colonne si trovano per nome: l'ordine dentro il blocco non conta."""
    from fasterreports.core.xlsxsource import index_to_col

    grid = {
        "A2": "Skill", "B2": "Surname", "C2": "Name", "D2": "Agent",
        "E2": "Zona", "F2": "Expected hours", "G2": "20/07/2026",
        "A3": "HPO", "B3": "Rossi", "C3": "Mario", "D3": "ROSSI",
        "E3": "2", "F3": "0800", "G3": "0900_1700_    _    ",
    }
    p = make_xlsx(tmp_path / "r.xlsx", "publish", grid)
    src = read_roster(p, week=week_bounds(46223))
    assert [r[0] for r in src.data] == ["Mario Rossi"]


def test_manca_una_colonna_fissa(tmp_path):
    grid = {
        "A2": "Agent", "B2": "Expected hours", "C2": "Zona", "D2": "Skill",
        "E2": "Name", "F2": "20/07/2026",  # manca Surname
        "A3": "ROSSI", "D3": "HPO", "E3": "Mario", "F3": "0900_1700_    _    ",
    }
    p = make_xlsx(tmp_path / "r.xlsx", "publish", grid)
    with pytest.raises(SourceError) as e:
        read_roster(p, week=week_bounds(46223))
    assert "Surname" in str(e.value)


def test_nessuna_colonna_agent(tmp_path):
    grid = {"A2": "Pippo", "B2": "20/07/2026", "A3": "x"}
    p = make_xlsx(tmp_path / "r.xlsx", "publish", grid)
    with pytest.raises(SourceError):
        read_roster(p, week=week_bounds(46223))


def test_settimana_obbligatoria_non_ritaglia_se_assente(tmp_path):
    """Senza settimana si prende tutto: il chiamante deve passarla."""
    grid = roster_grid(dates=("20/07/2026", "27/07/2026"), agents=[
        ("A", "0800", "2", "HPO", "Uno", "Uno",
         ["0900_1700_    _    ", "0900_1700_    _    "]),
    ])
    p = make_xlsx(tmp_path / "r.xlsx", "publish", grid)
    assert len(read_roster(p).data) == 2


# --- back office -----------------------------------------------------------

def backoffice_grid(rows, dates=(46223, 46224)):
    grid: dict[str, object] = {"A1": "RICHIEDI CAMBIO TURNO", "B1": 107}
    from fasterreports.core.xlsxsource import index_to_col

    for j, d in enumerate(dates):
        grid[f"{index_to_col(4 + j)}1"] = d
    for r, (sez, nome, cog, cells) in enumerate(rows, start=2):
        if sez is not None:
            grid[f"A{r}"] = sez
        if nome is not None:
            grid[f"B{r}"] = nome
        if cog is not None:
            grid[f"C{r}"] = cog
        for j, v in enumerate(cells):
            if v is not None:
                grid[f"{index_to_col(4 + j)}{r}"] = v
    return grid


def test_backoffice_unpivot(tmp_path):
    grid = backoffice_grid([("HPO", "Mario", "Rossi", ["1000_1030", "NO BOT"])])
    p = make_xlsx(tmp_path / "b.xlsx", "Only Cases Shifts", grid)
    src = read_backoffice(p, week=week_bounds(46223))
    assert len(src.data) == 2
    assert src.data[0][0] == "mario rossi"
    assert src.data[0][4] == "BOT"
    assert src.data[1][4] == "NO BOT"


def test_backoffice_salta_le_righe_di_totale(tmp_path):
    grid = backoffice_grid([
        ("HPO", "Mario", "Rossi", ["1000_1030", "NO BOT"]),
        ("Total BO Hrs", None, None, [10, 12]),
        (None, None, "intervals", [1, 2]),
    ])
    p = make_xlsx(tmp_path / "b.xlsx", "Only Cases Shifts", grid)
    src = read_backoffice(p, week=week_bounds(46223))
    assert {r[0] for r in src.data} == {"mario rossi"}
    assert src.notes.rows_skipped


def test_backoffice_filtro_allowed_keys(tmp_path):
    grid = backoffice_grid([
        ("HPO", "Mario", "Rossi", ["1000_1030", "NO BOT"]),
        ("HPO", "Lucia", "Verdi", ["1000_1030", "NO BOT"]),
    ])
    p = make_xlsx(tmp_path / "b.xlsx", "Only Cases Shifts", grid)
    src = read_backoffice(p, week=week_bounds(46223), allowed_keys={"mario rossi"})
    assert {r[0] for r in src.data} == {"mario rossi"}
    assert "lucia verdi" in src.notes.keys_not_allowed


def test_backoffice_alias_solo_verso_lo_spazio_consentito(tmp_path):
    """L'alias si usa per *entrare* nello spazio del roster, non per uscirne."""
    grid = backoffice_grid([
        ("HPO", "Alessandro", "Passierello", ["1000_1030", "NO BOT"]),
        ("HPO", "Nadia", "Ariefieva", ["1000_1030", "NO BOT"]),
    ])
    p = make_xlsx(tmp_path / "b.xlsx", "Only Cases Shifts", grid)
    aliases = {
        "alessandro passierello": "alessandro passariello",  # verso il roster
        "nadia ariefieva": "nadiia ariefieva",               # fuori dal roster
    }
    src = read_backoffice(
        p, week=week_bounds(46223), aliases=aliases,
        allowed_keys={"alessandro passariello", "nadia ariefieva"},
    )
    keys = {r[0] for r in src.data}
    assert keys == {"alessandro passariello", "nadia ariefieva"}
    assert "alessandro passierello" in src.notes.aliases_applied
    assert "nadia ariefieva" not in src.notes.aliases_applied


def test_backoffice_senza_date_blocca(tmp_path):
    grid = {"A1": "x", "B1": "Nome", "C1": "Cognome", "A2": "HPO"}
    p = make_xlsx(tmp_path / "b.xlsx", "Only Cases Shifts", grid)
    with pytest.raises(SourceError) as e:
        read_backoffice(p, week=week_bounds(46223))
    assert "intestazione" in str(e.value).lower()


# --- lettore xlsx ----------------------------------------------------------

def test_read_sheet_foglio_assente(tmp_path):
    p = make_xlsx(tmp_path / "x.xlsx", "uno", {"A1": "v"})
    with pytest.raises(SourceError) as e:
        read_sheet(p, "due")
    assert "'uno'" in str(e.value)


def test_read_sheet_file_non_xlsx(tmp_path):
    p = tmp_path / "finto.xlsx"
    p.write_text("non sono uno zip", encoding="utf-8")
    with pytest.raises(SourceError) as e:
        read_sheet(p)
    assert "non è un file .xlsx valido" in str(e.value)


def test_find_header_row():
    sheet = Sheet(name="s", rows={
        1: {"A": "titolo"},
        2: {"A": "Agent", "B": "Expected hours", "D": "Skill", "E": "Name", "F": "Surname"},
    })
    assert find_header_row(sheet, ("Agent", "Expected hours", "Zona", "Skill", "Name", "Surname")) == 2


def test_find_header_row_assente():
    sheet = Sheet(name="s", rows={1: {"A": "pippo"}})
    with pytest.raises(SourceError) as e:
        find_header_row(sheet, ("Agent", "Skill", "Name", "Surname"))
    assert "intestazione" in str(e.value)


def test_find_roster_blocks_ordine():
    sheet = Sheet(name="s", rows={
        2: {
            "A": "Agent", "B": "Expected hours", "C": "Zona", "D": "Skill",
            "E": "Name", "F": "Surname", "G": "20/07/2026",
            "K": "Agent", "L": "Expected hours", "M": "Zona", "N": "Skill",
            "O": "Name", "P": "Surname", "Q": "21/07/2026",
        }
    })
    blocks = find_roster_blocks(sheet, 2)
    assert len(blocks) == 2
    assert blocks[0].fixed["Name"] == "E"
    assert blocks[1].fixed["Name"] == "O"
    assert [c for c, _ in blocks[0].dates] == ["G"]
    assert [c for c, _ in blocks[1].dates] == ["Q"]
