"""Dataset tabellari che arrivano come foglio Excel invece che come CSV.

Gli export reali sono misti: `SF DATABASE W31.csv` e `PSAT DATASET W31.csv` da un
lato, `AT DATASET W31.xlsx` e `ATwi DATASET W31.xlsx` dall'altro. Il formato
dipende da chi produce l'export, non da noi, quindi la pipeline deve reggere
entrambi senza che nessuno converta niente a mano.

Quello che questi test difendono e' che a valle **non cambi nulla**: dopo il
lettore, un foglio e un CSV sono la stessa coppia (headers, rows) e il matcher
lavora per nome come sempre.
"""

from __future__ import annotations

import pytest

from fasterreports.core.coerce import to_datetime, to_float
from fasterreports.core.errors import SourceError
from fasterreports.core.tablesource import (
    find_header_row,
    is_excel,
    read_table,
)
from fasterreports.core.xlsxsource import read_sheet
from xlsxbuild import make_workbook, make_xlsx

# ---------------------------------------------------------------------------
# Riconoscimento del formato
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "nome, atteso",
    [
        ("AT DATASET W31.xlsx", True),
        ("AT DATASET W31.XLSX", True),  # Windows non guarda la cassa
        ("Turni_W31.xlsm", True),
        ("modello.xltx", True),
        ("SF DATABASE W31.csv", False),
        ("export.txt", False),
        ("senza_estensione", False),
        # .xls vecchio formato: NON e' uno zip, il nostro lettore non lo apre.
        # Meglio dire "non e' Excel" e far fallire csvsource con un messaggio
        # sull'encoding, che aprire uno zip inesistente.
        ("vecchio.xls", False),
    ],
)
def test_is_excel(nome, atteso):
    assert is_excel(nome) is atteso


# ---------------------------------------------------------------------------
# Lettura di base
# ---------------------------------------------------------------------------


def test_legge_headers_e_righe(tmp_path):
    p = make_xlsx(
        tmp_path / "at.xlsx",
        "Sheet1",
        {
            "A1": "Case Number", "B1": "Start Time", "C1": "Handle Time",
            "A2": "5001", "B2": 46223.25, "C2": 305.5,
            "A3": "5002", "B3": 46223.5, "C3": 120,
        },
    )
    src = read_table(p)
    assert src.headers == ["Case Number", "Start Time", "Handle Time"]
    assert src.n_cols == 3
    righe = list(src.rows())
    assert len(righe) == 2
    assert righe[0][0] == "5001"


def test_stesso_esito_con_shared_strings(tmp_path):
    """Excel vero scrive il testo in sharedStrings, non inline.

    Se il lettore funzionasse solo con l'una o con l'altra forma, i test
    passerebbero e l'export reale no.
    """
    grid = {"A1": "Agent Name", "B1": "Skill", "A2": "Mario Rossi", "B2": "HPO"}
    inline = read_table(make_xlsx(tmp_path / "i.xlsx", "S", grid, inline=True))
    shared = read_table(make_xlsx(tmp_path / "s.xlsx", "S", grid, inline=False))
    assert inline.headers == shared.headers == ["Agent Name", "Skill"]
    assert list(inline.rows()) == list(shared.rows()) == [["Mario Rossi", "HPO"]]


def test_intestazioni_ripulite_dagli_spazi(tmp_path):
    """Gli export mettono spazi di troppo: il matcher normalizza, ma l'header
    grezzo serve pulito perche' il match esatto viene *prima* di quello
    normalizzato (vedi il caso 'Agent Name' / 'agent_name')."""
    p = make_xlsx(tmp_path / "x.xlsx", "S", {"A1": "  Case Number  ", "A2": "1"})
    assert read_table(p).headers == ["Case Number"]


def test_salta_le_righe_vuote_di_coda(tmp_path):
    """Un foglio "usato" conserva righe con formattazione ma senza valori: sono
    coda, non dati. Contarle come righe gonfierebbe il conteggio del preflight."""
    p = make_xlsx(
        tmp_path / "x.xlsx",
        "S",
        {
            "A1": "a", "B1": "b",
            "A2": "1", "B2": "2",
            "A3": "", "B3": "   ",      # riga solo spazi
            "A4": "3", "B4": "4",
        },
    )
    righe = list(read_table(p).rows())
    assert len(righe) == 2
    assert [r[0] for r in righe] == ["1", "3"]


def test_colonna_senza_intestazione_resta_nel_blocco(tmp_path):
    """Un buco nell'intestazione non deve spostare le colonne successive.

    E' il guasto che tutto il progetto combatte: se `C` senza nome facesse
    scalare `D`, i valori finirebbero nella colonna sbagliata in silenzio.
    """
    p = make_xlsx(
        tmp_path / "x.xlsx",
        "S",
        {"A1": "uno", "B1": "due", "D1": "quattro",
         "A2": "a", "B2": "b", "C2": "c", "D2": "d"},
    )
    src = read_table(p)
    assert src.headers == ["uno", "due", "", "quattro"]
    assert list(src.rows()) == [["a", "b", "c", "d"]]


def test_cella_mancante_diventa_none_non_uno_scorrimento(tmp_path):
    p = make_xlsx(
        tmp_path / "x.xlsx",
        "S",
        {"A1": "uno", "B1": "due", "C1": "tre",
         "A2": "a", "C2": "c"},
    )
    assert list(read_table(p).rows()) == [["a", None, "c"]]


# ---------------------------------------------------------------------------
# Riga di intestazione non in riga 1
# ---------------------------------------------------------------------------


def test_trova_l_intestazione_sotto_un_titolo(tmp_path):
    """Alcuni export mettono un titolo in riga 1 e le intestazioni sotto."""
    p = make_xlsx(
        tmp_path / "x.xlsx",
        "S",
        {
            "A1": "Report settimanale",           # una cella sola: non e' un header
            "A3": "Case Number", "B3": "Agent Name", "C3": "Skill",
            "A4": "5001", "B4": "Mario Rossi", "C4": "HPO",
        },
    )
    src = read_table(p)
    assert src.headers == ["Case Number", "Agent Name", "Skill"]
    assert list(src.rows()) == [["5001", "Mario Rossi", "HPO"]]


def test_header_row_esplicito_vince(tmp_path):
    """Quando l'euristica sbaglia si deve poter dire la riga a mano."""
    p = make_xlsx(
        tmp_path / "x.xlsx",
        "S",
        {"A1": "filtro", "B1": "attivo",
         "A2": "Case Number", "B2": "Agent Name",
         "A3": "1", "B3": "Rossi"},
    )
    src = read_table(p, header_row=2)
    assert src.headers == ["Case Number", "Agent Name"]
    assert list(src.rows()) == [["1", "Rossi"]]


def test_find_header_row_salta_le_righe_di_soli_numeri(tmp_path):
    p = make_xlsx(
        tmp_path / "x.xlsx", "S",
        {"A1": 1, "B1": 2, "C1": 3,
         "A2": "uno", "B2": "due", "C2": "tre"},
    )
    assert find_header_row(read_sheet(p, "S")) == 2


def test_find_header_row_foglio_senza_intestazioni(tmp_path):
    p = make_xlsx(tmp_path / "x.xlsx", "S", {"A1": 1, "A2": 2})
    with pytest.raises(SourceError) as e:
        read_table(p)
    msg = str(e.value)
    assert "intestazione" in msg          # cosa manca
    assert "prevalenza testo" in msg      # cosa cercava
    assert "header_row" in msg            # come si sistema


# ---------------------------------------------------------------------------
# Piu' fogli
# ---------------------------------------------------------------------------


def test_piu_fogli_prende_il_primo(tmp_path):
    p = make_workbook(
        tmp_path / "x.xlsx",
        [
            ("Dati", {"A1": "uno", "A2": "primo"}),
            ("Note", {"A1": "altro", "A2": "secondo"}),
        ],
    )
    src = read_table(p)
    assert src.sheet == "Dati"
    # Il foglio scelto finisce nella descrizione, cosi' il preflight lo dichiara
    # invece di indovinare in silenzio.
    assert "Dati" in src.delimiter


def test_sheet_name_seleziona_il_foglio(tmp_path):
    p = make_workbook(
        tmp_path / "x.xlsx",
        [("Dati", {"A1": "uno", "A2": "primo"}),
         ("Note", {"A1": "altro", "A2": "secondo"})],
    )
    src = read_table(p, sheet_name="Note")
    assert src.sheet == "Note"
    assert list(src.rows()) == [["secondo"]]


def test_foglio_inesistente_elenca_quelli_presenti(tmp_path):
    p = make_workbook(tmp_path / "x.xlsx", [("Dati", {"A1": "uno", "A2": "x"})])
    with pytest.raises(SourceError) as e:
        read_table(p, sheet_name="Sbagliato")
    assert "'Dati'" in str(e.value)


def test_file_non_zip(tmp_path):
    p = tmp_path / "finto.xlsx"
    p.write_bytes(b"non sono uno zip")
    with pytest.raises(SourceError) as e:
        read_table(p)
    assert ".xlsx" in str(e.value)


def test_file_assente(tmp_path):
    with pytest.raises(SourceError):
        read_table(tmp_path / "manca.xlsx")


# ---------------------------------------------------------------------------
# I valori: seriali e decimali passano dalle coercizioni esistenti
# ---------------------------------------------------------------------------


def test_le_date_arrivano_come_seriali_e_to_datetime_le_capisce(tmp_path):
    """In un CSV la data e' testo ambiguo (01/02 e' gennaio o febbraio?); in un
    foglio e' un numero, che e' *meno* rischioso. Non serve convertire qui:
    `to_datetime` gestisce entrambe le forme."""
    p = make_xlsx(
        tmp_path / "x.xlsx", "S",
        {"A1": "Start Time", "A2": 46223.25},   # 20/07/2026 06:00
    )
    (riga,) = read_table(p).rows()
    dt = to_datetime(riga[0])
    assert (dt.year, dt.month, dt.day, dt.hour) == (2026, 7, 20, 6)


def test_i_decimali_arrivano_col_punto_e_to_float_li_capisce(tmp_path):
    p = make_xlsx(tmp_path / "x.xlsx", "S", {"A1": "Handle Time", "A2": 305.5})
    (riga,) = read_table(p).rows()
    assert to_float(riga[0]) == 305.5


def test_errore_excel_in_cache_non_diventa_un_valore(tmp_path):
    """Un `#N/D` in cache deve arrivare come tale, non come zero o vuoto: e'
    un dato mancante travestito, ed e' il preflight che deve vederlo."""
    p = tmp_path / "x.xlsx"
    make_xlsx(p, "S", {"A1": "Handle Time", "A2": 1})
    import zipfile

    parti = {}
    with zipfile.ZipFile(p) as z:
        parti = {n: z.read(n) for n in z.namelist()}
    parti["xl/worksheets/sheet1.xml"] = (
        b'<?xml version="1.0"?><worksheet><sheetData>'
        b'<row r="1"><c r="A1" t="inlineStr"><is><t>Handle Time</t></is></c></row>'
        b'<row r="2"><c r="A2" t="e"><v>#N/D</v></c></row>'
        b"</sheetData></worksheet>"
    )
    with zipfile.ZipFile(p, "w") as z:
        for n, data in parti.items():
            z.writestr(n, data)

    (riga,) = read_table(p).rows()
    assert riga[0] == "#N/D"


# ---------------------------------------------------------------------------
# Interfaccia comune con csvsource: e' cio' che rende il resto della pipeline
# indifferente al formato
# ---------------------------------------------------------------------------


def test_espone_la_stessa_interfaccia_di_csvsource(tmp_path):
    from fasterreports.core.csvsource import read_csv_text

    grid = {"A1": "uno", "B1": "due", "A2": "a", "B2": "b"}
    tab = read_table(make_xlsx(tmp_path / "x.xlsx", "S", grid))
    csv = read_csv_text("uno,due\na,b\n")

    for attr in ("path", "headers", "encoding", "delimiter", "n_cols", "rows"):
        assert hasattr(tab, attr), attr
        assert hasattr(csv, attr), attr
    assert tab.headers == csv.headers
    assert list(tab.rows()) == list(csv.rows())
