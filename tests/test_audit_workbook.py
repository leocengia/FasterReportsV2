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


def test_i_sei_fogli_dati_sono_nei_default():
    """La fixture nasceva coprendo solo i 4 CSV: `Turni` e `Slot Only Cases`
    erano entrati nel contratto senza entrare qui, e restavano non verificati."""
    from audit_workbook import DEFAULT_DATASETS

    assert set(DEFAULT_DATASETS) == {
        "AT_DATASET",
        "ATwi_DATASET",
        "SF_DATABASE",
        "PSAT_DATASET",
        "Turni",
        "Slot Only Cases",
    }
