"""Il lettore .xlsx: e' lo strumento con cui si verifica tutto il resto.

Questo file esiste per una ragione specifica. Il lettore aveva un difetto che non
produceva errori: produceva **numeri nella colonna sbagliata**. In un modulo che
serve a controllare i numeri di qualcun altro, e' il guasto peggiore possibile —
e infatti ha fatto passare per settimane un test golden che certificava una
differenza inesistente, e ha fatto leggere `NO BOT` come `466` in un audit.

Uno strumento di misura va misurato.
"""

from __future__ import annotations

import zipfile

import pytest

from fasterreports.core.errors import SourceError
from fasterreports.core.xlsxsource import index_to_col, read_sheet, read_sheet_names
from xlsxbuild import make_workbook, make_xlsx


def scrivi_grezzo(path, sheet_xml: str, shared: list[str] | None = None):
    """Costruisce un .xlsx con un corpo di foglio scritto a mano.

    Serve perche' il difetto stava nel parsing di forme XML che il costruttore
    normale non produce: la cella vuota autochiudente `<c r="A1" s="3"/>`, che
    Excel scrive per ogni cella formattata ma senza valore.
    """
    make_xlsx(path, "S", {"A1": "segnaposto"}, inline=shared is None)
    with zipfile.ZipFile(path) as z:
        parti = {n: z.read(n) for n in z.namelist()}
    parti["xl/worksheets/sheet1.xml"] = (
        '<?xml version="1.0"?><worksheet><sheetData>'
        f"{sheet_xml}"
        "</sheetData></worksheet>"
    ).encode("utf-8")
    if shared is not None:
        si = "".join(f"<si><t>{s}</t></si>" for s in shared)
        parti["xl/sharedStrings.xml"] = (
            f'<?xml version="1.0"?><sst count="{len(shared)}" '
            f'uniqueCount="{len(shared)}">{si}</sst>'
        ).encode("utf-8")
    with zipfile.ZipFile(path, "w") as z:
        for n, data in parti.items():
            z.writestr(n, data)
    return path


# ---------------------------------------------------------------------------
# Il difetto
# ---------------------------------------------------------------------------


def test_cella_vuota_autochiudente_non_si_mangia_le_successive(tmp_path):
    """Il guasto, nella sua forma minima.

    `<c r="C1" s="120"/>` non chiude nulla: il pattern che cercava `...</c>`
    partiva da C1 e arrivava fino al `</c>` di E1, prendendosi il valore di E1 e
    attribuendolo a C1. Le celle in mezzo scomparivano.

    Sul workbook vero questo faceva leggere `Slot inizio = 466` dove c'era
    `Stato BO = 'NO BOT'` (466 era l'indice della stringa condivisa).
    """
    p = scrivi_grezzo(
        tmp_path / "x.xlsx",
        '<row r="1">'
        '<c r="A1" t="s"><v>0</v></c>'
        '<c r="B1" s="2"><v>46235</v></c>'
        '<c r="C1" s="120"/>'
        '<c r="D1" s="120"/>'
        '<c r="E1" t="s"><v>1</v></c>'
        "</row>",
        shared=["viktoriia lavrinets", "NO BOT"],
    )
    row = read_sheet(p, "S").row(1)
    assert row == {"A": "viktoriia lavrinets", "B": "46235", "E": "NO BOT"}
    # In particolare: C non esiste, e NON vale l'indice della stringa di E.
    assert "C" not in row
    assert "466" not in str(row)


def test_valore_dopo_una_cella_vuota_resta_nella_sua_colonna(tmp_path):
    p = scrivi_grezzo(
        tmp_path / "x.xlsx",
        '<row r="1"><c r="A1" s="1"/><c r="B1"><v>7</v></c></row>',
    )
    assert read_sheet(p, "S").row(1) == {"B": "7"}


def test_piu_celle_vuote_di_fila(tmp_path):
    p = scrivi_grezzo(
        tmp_path / "x.xlsx",
        '<row r="1">'
        '<c r="A1"><v>1</v></c><c r="B1" s="1"/><c r="C1" s="1"/><c r="D1" s="1"/>'
        '<c r="E1"><v>5</v></c>'
        "</row>",
    )
    assert read_sheet(p, "S").row(1) == {"A": "1", "E": "5"}


def test_formula_che_restituisce_stringa_vuota_non_e_un_valore(tmp_path):
    """`IF($I="","",...)` su una riga vuota: nel file resta `<v/>`.

    E' il caso delle colonne P/Q di AT_DATASET oltre l'ultima riga di dati.
    """
    p = scrivi_grezzo(
        tmp_path / "x.xlsx",
        '<row r="1">'
        '<c r="A1"><v>3</v></c>'
        '<c r="P1" t="str"><f>IF($I1="","",1)</f><v/></c>'
        '<c r="Q1"><v>9</v></c>'
        "</row>",
    )
    row = read_sheet(p, "S").row(1)
    assert row["A"] == "3"
    assert row["Q"] == "9"          # non risucchiato da P
    assert row.get("P", "") == ""


# ---------------------------------------------------------------------------
# Le forme normali, perche' la correzione non deve rompere quelle
# ---------------------------------------------------------------------------


def test_stringa_condivisa_e_inline(tmp_path):
    p = scrivi_grezzo(
        tmp_path / "x.xlsx",
        '<row r="1">'
        '<c r="A1" t="s"><v>0</v></c>'
        '<c r="B1" t="inlineStr"><is><t>diretta</t></is></c>'
        "</row>",
        shared=["condivisa"],
    )
    assert read_sheet(p, "S").row(1) == {"A": "condivisa", "B": "diretta"}


def test_errore_excel_in_cache_arriva_come_tale(tmp_path):
    p = scrivi_grezzo(
        tmp_path / "x.xlsx",
        '<row r="1"><c r="A1" t="e"><v>#N/D</v></c><c r="B1"><v>1</v></c></row>',
    )
    assert read_sheet(p, "S").row(1) == {"A": "#N/D", "B": "1"}


def test_numeri_e_testo_con_caratteri_da_escapare(tmp_path):
    p = make_xlsx(
        tmp_path / "x.xlsx",
        "S",
        {"A1": "a < b & c", "B1": 3.5, "C1": "d'apostrofo"},
    )
    row = read_sheet(p, "S").row(1)
    assert row == {"A": "a < b & c", "B": "3.5", "C": "d'apostrofo"}


def test_colonne_oltre_la_z(tmp_path):
    grid = {f"{index_to_col(i)}1": i for i in (1, 26, 27, 52, 53, 703)}
    p = make_xlsx(tmp_path / "x.xlsx", "S", grid)
    row = read_sheet(p, "S").row(1)
    assert row["A"] == "1" and row["Z"] == "26" and row["AA"] == "27"
    assert row["AZ"] == "52" and row["BA"] == "53" and row["AAA"] == "703"


def test_cols_of_in_ordine_di_foglio(tmp_path):
    p = make_xlsx(tmp_path / "x.xlsx", "S", {"A1": 1, "Z1": 2, "AA1": 3, "B1": 4})
    assert read_sheet(p, "S").cols_of(1) == ["A", "B", "Z", "AA"]


def test_riga_completamente_vuota_non_entra(tmp_path):
    p = scrivi_grezzo(
        tmp_path / "x.xlsx",
        '<row r="1"><c r="A1"><v>1</v></c></row>'
        '<row r="2"><c r="A2" s="1"/><c r="B2" s="1"/></row>'
        '<row r="3"><c r="A3"><v>3</v></c></row>',
    )
    sh = read_sheet(p, "S")
    assert sorted(sh.rows) == [1, 3]
    assert sh.max_row == 3


# ---------------------------------------------------------------------------
# Errori parlanti
# ---------------------------------------------------------------------------


def test_file_non_zip(tmp_path):
    p = tmp_path / "finto.xlsx"
    p.write_bytes(b"non sono uno zip")
    for fn in (lambda: read_sheet(p), lambda: read_sheet_names(p)):
        with pytest.raises(SourceError) as e:
            fn()
        assert ".xlsx" in str(e.value)


def test_foglio_inesistente_elenca_i_presenti(tmp_path):
    p = make_workbook(tmp_path / "x.xlsx", [("Uno", {"A1": 1}), ("Due", {"A1": 2})])
    with pytest.raises(SourceError) as e:
        read_sheet(p, "Tre")
    assert "'Uno'" in str(e.value) and "'Due'" in str(e.value)
    assert read_sheet_names(p) == ["Uno", "Due"]


def test_indice_stringa_condivisa_non_intero(tmp_path):
    p = scrivi_grezzo(
        tmp_path / "x.xlsx",
        '<row r="1"><c r="A1" t="s"><v>12.0</v></c></row>',
        shared=["x"],
    )
    with pytest.raises(SourceError) as e:
        read_sheet(p, "S")
    msg = str(e.value)
    assert "A1" in msg and "'12.0'" in msg and "corrotto" in msg
