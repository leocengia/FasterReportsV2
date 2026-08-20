"""Lo strumento che genera la fixture del contratto.

Va testato per la stessa ragione del lettore .xlsx: se sbaglia, non sbaglia da
solo. La fixture `workbook_W30.json` alimenta il test che verifica che ogni
colonna consumata dalle formule sia nel contratto — e quel test e' cieco su
tutto cio' che l'audit non vede.

E' successo davvero: le formule di 'AHT Outliers' citano le colonne come
riferimenti strutturati (`AHT_Data[Case Type]`), l'audit cercava solo
`SF_DATABASE!$Z`, e due colonne vere sono rimaste fuori dal contratto. Risultato:
il foglio 'AHT Outliers' e' uscito vuoto, senza che nulla lo segnalasse.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

from audit_workbook import Workbook, col_index, col_letters  # noqa: E402

SAMPLE = ROOT / "samples" / "omni-report" / "Omni Report W30.xlsm"


@pytest.fixture(scope="module")
def wb():
    if not SAMPLE.is_file():
        pytest.skip(f"campione assente: {SAMPLE}")
    return Workbook(SAMPLE)


def test_col_letters_e_inversa_di_col_index():
    for i in (1, 26, 27, 52, 53, 143, 703):
        assert col_index(col_letters(i)) == i


# ---------------------------------------------------------------------------
# Riferimenti strutturati: il buco che e' costato un foglio vuoto
# ---------------------------------------------------------------------------


def test_le_colonne_della_tabella_si_mappano_a_lettere(wb):
    tabelle = wb.table_columns()
    assert "AHT_Data" in tabelle
    foglio, lettere = tabelle["AHT_Data"]
    assert foglio == "SF_DATABASE"
    # La tabella parte da A1, quindi la n-esima colonna e' la n-esima lettera.
    assert lettere["Date Viewpoint"] == "A"
    assert lettere["Case Type"] == "Z"
    assert lettere["Employee Name"] == "BB"
    assert lettere["Primary Category"] == "BV"
    assert lettere["Case AHT (mins)"] == "DY"


def test_usage_vede_i_riferimenti_strutturati(wb):
    """`AHT_Data[Case Type]` deve contare come uso di `SF_DATABASE!Z`."""
    usage = wb.formula_usage(["SF_DATABASE"])
    assert "Z" in usage["SF_DATABASE"], "Case Type non rilevata"
    assert "BV" in usage["SF_DATABASE"], "Primary Category non rilevata"
    assert "AHT Outliers" in usage["SF_DATABASE"]["Z"]


def test_usage_vede_anche_i_riferimenti_per_lettera(wb):
    """La forma classica non deve essere andata persa nel farci stare l'altra."""
    usage = wb.formula_usage(["AT_DATASET", "PSAT_DATASET"])
    assert {"B", "F", "K", "L"} <= set(usage["AT_DATASET"])
    assert {"I", "DP"} <= set(usage["PSAT_DATASET"])


def test_nomi_di_foglio_fra_apici(wb):
    """`'Slot Only Cases'!$A$2`: senza gli apici il foglio sembrava letto da
    nessuno."""
    usage = wb.formula_usage(["Slot Only Cases", "Turni"])
    # 'Turni' e' letto da 'Helper Turni' per tutte le sue colonne.
    assert {"A", "B", "E", "F"} <= set(usage["Turni"])


def test_un_dataset_non_prende_le_colonne_di_un_altro(wb):
    """`AT_DATASET` e' sottostringa di `PSAT_DATASET`, `Turni` di
    `Helper Turni`: senza confine a sinistra le colonne si mescolano."""
    usage = wb.formula_usage(["AT_DATASET", "PSAT_DATASET"])
    # PSAT ha colonne DO/DP/DQ che AT non ha (AT arriva a Q).
    assert not ({"DO", "DP", "DQ"} & set(usage["AT_DATASET"]))


def test_tutti_i_dataset_del_contratto_sono_nei_default(contract):
    """Un dataset nel contratto ma non qui resta NON VERIFICATO.

    E' successo due volte. La prima: la fixture copriva solo i 4 CSV, e `Turni` e
    `Slot Only Cases` erano entrati nel contratto senza entrare qui. La seconda:
    `DUP_DATASET`, il 2026-08-19.

    Per questo il confronto non e' con una lista scritta a mano — che e' proprio
    la cosa che va fuori sincrono — ma col contratto.
    """
    from audit_workbook import DEFAULT_DATASETS

    mancanti = set(contract.datasets) - set(DEFAULT_DATASETS)
    assert not mancanti, (
        f"dataset nel contratto ma non in DEFAULT_DATASETS: {sorted(mancanti)}. "
        f"La fixture nascerebbe senza di loro, e le loro colonne non verrebbero "
        f"verificate contro il template."
    )
    in_piu = set(DEFAULT_DATASETS) - set(contract.datasets)
    assert not in_piu, (
        f"dataset in DEFAULT_DATASETS ma non nel contratto: {sorted(in_piu)}."
    )


def test_header_rows_legge_il_contratto(contract):
    """La riga delle intestazioni non si scrive due volte.

    `DUP_DATASET` e' il primo dataset con le intestazioni non in riga 1 (sotto il
    preambolo del report Salesforce). Quel numero e' dichiarato nel contratto —
    ed e' quello che il writer usa per scrivere. Se lo strumento ne tenesse una
    copia, le due potrebbero divergere e la fixture nascerebbe leggendo la riga
    sbagliata: intestazioni finte, e un contratto verificato contro di loro.
    """
    from audit_workbook import header_rows

    righe = header_rows()
    for ds in contract.datasets.values():
        assert righe.get(ds.sheet) == ds.header_row, (
            f"{ds.sheet}: il contratto dice riga {ds.header_row}, "
            f"lo strumento legge {righe.get(ds.sheet)}"
        )


# ---------------------------------------------------------------------------
# Celle e formule AUTO-CHIUSE: il buco che avrebbe inventato due intestazioni
# ---------------------------------------------------------------------------


def test_una_cella_vuota_formattata_non_ruba_il_valore_alla_successiva(tmp_path):
    """`<c r="A1" s="1"/>` non deve consumare il contenuto di `B1`.

    Il difetto misurato il 2026-08-19 su `DUP_DATASET`: la riga 14 ha `A` e `C`
    vuote-ma-formattate, e lo strumento riportava `A: '9136'`, `C: '1364'` —
    indici grezzi di sharedStrings — con `Full Name` e `Case Number` **spariti**.
    La fixture sarebbe nata con due intestazioni inventate, e il test che
    verifica il contratto le avrebbe prese per vere.

    Con `inline=False` perche' e' la forma che produce Excel davvero: e' proprio
    la risoluzione della stringa condivisa che saltava.
    """
    from xlsxbuild import VUOTA_FORMATTATA, make_workbook

    p = make_workbook(
        tmp_path / "vuote.xlsx",
        [("Foglio", {
            "A1": VUOTA_FORMATTATA,
            "B1": "Full Name",
            "C1": VUOTA_FORMATTATA,
            "D1": "Case Number",
            "E1": "Status",
        })],
        inline=False,
    )
    assert Workbook(p).header_row("Foglio") == {
        "B": "Full Name",
        "D": "Case Number",
        "E": "Status",
    }


def test_le_formule_condivise_non_incollano_celle_diverse(tmp_path):
    """`<f t="shared" si="6"/>` e' auto-chiuso: non e' un tag di apertura.

    Trattarlo come tale fa catturare tutto il testo fino al `</f>` successivo,
    cioe' unisce le formule di celle diverse in una stringa sola. In mezzo alle
    due c'e' una cella con formula condivisa: se il suo tag venisse letto come
    apertura, le tre celle uscirebbero come un unico testo.
    """
    from xlsxbuild import Condivisa, make_workbook

    p = make_workbook(
        tmp_path / "condivise.xlsx",
        [("Report", {
            "A1": "=SUM(Turni!$D$2:$D$99)",
            "B1": Condivisa(si=0, valore=1),          # erede, senza testo
            "C1": "=SUM(PSAT_DATASET!$DP$2:$DP$99)",
        })],
    )
    from audit_workbook import formule

    # Uguaglianza esatta, non "contiene": un testo incollato conterrebbe
    # comunque entrambe le sottostringhe, e un assert di appartenenza sarebbe
    # verde su un lettore rotto.
    assert formule(Workbook(p).xml("Report")) == [
        "SUM(Turni!$D$2:$D$99)",
        "SUM(PSAT_DATASET!$DP$2:$DP$99)",
    ]


def test_i_nomi_dei_fogli_sono_de_escapati(tmp_path):
    """`DC Agents & Categories`, non `DC Agents &amp; Categories`.

    `templatescan._sheet_targets` lo faceva gia': lo stesso foglio risultava con
    due nomi diversi a seconda di chi lo leggeva.
    """
    from xlsxbuild import make_workbook

    p = make_workbook(tmp_path / "amp.xlsx", [("DC Agents & Categories", {"A1": 1})])
    assert "DC Agents & Categories" in Workbook(p).sheets


def test_last_data_row_ignora_una_dimension_stantia(wb):
    """L'ultima riga con un valore, non `<dimension>`.

    In `SF_DATABASE` del template `<dimension>` dice `A1:EM3568` mentre l'ultima
    riga con un valore e' la 3389 — la stessa dell'intervallo di `AHT_Data`. Il
    confronto sulla dimensione stampava un ATTENZIONE per 179 righe di dati che
    non esistono: un falso allarme nello strumento che serve a trovare quelli veri.
    """
    dim = wb.dimension("SF_DATABASE")
    ultima = wb.last_data_row("SF_DATABASE")
    fine_dim = int("".join(c for c in dim.split(":")[-1] if c.isdigit()))
    assert 0 < ultima <= fine_dim
