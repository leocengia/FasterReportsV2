from __future__ import annotations

import pytest

from fasterreports.core.contract import (
    col_to_index,
    index_to_col,
    load_contract,
    parse_contract,
)
from fasterreports.core.errors import ContractError, DuplicateTargetError

BASE = {
    "normalize": {},
    "datasets": {
        "D": {
            "sheet": "D",
            "header_row": 1,
            "data_start_col": "A",
            "fields": [{"canonical": "Alpha", "target_col": "A", "role": "input"}],
        }
    },
}


def _with_fields(fields, **ds_kw):
    raw = {"normalize": {}, "datasets": {"D": {
        "sheet": "D", "header_row": 1, "data_start_col": "A", "fields": fields, **ds_kw,
    }}}
    return raw


@pytest.mark.parametrize(
    "letters,index",
    [("A", 1), ("B", 2), ("Z", 26), ("AA", 27), ("AC", 29), ("BB", 54), ("DY", 129), ("EM", 143)],
)
def test_conversione_lettere(letters, index):
    assert col_to_index(letters) == index
    assert index_to_col(index) == letters


def test_lettera_non_valida():
    with pytest.raises(ContractError):
        col_to_index("A1")


def test_contratto_reale_si_carica(contract):
    assert set(contract.datasets) == {
        "AT_DATASET", "ATwi_DATASET", "SF_DATABASE", "PSAT_DATASET", "DUP_DATASET",
        "Turni", "Slot Only Cases",
    }
    at = contract.dataset("AT_DATASET")
    assert at.list_object is None
    assert at.max_template_row == 130000
    assert contract.dataset("SF_DATABASE").list_object == "AHT_Data"


def test_dup_dataset_e_il_solo_con_le_intestazioni_non_in_riga_1(contract):
    """`DUP_DATASET` e' il report Salesforce formattato, preambolo compreso.

    Le intestazioni in riga 14 non sono una preferenza: `Duplicates Helper` legge
    quel foglio cella per cella con offset fisso +13 (la sua riga 2 e'
    `DUP_DATASET!B15`), quindi 14 e' il contratto con quel foglio. Vedi anche
    `test_offset_helper_duplicates`, che lo verifica dall'altro capo — sul
    template vero.
    """
    dup = contract.dataset("DUP_DATASET")
    assert dup.header_row == 14
    assert dup.data_start_col == "B"
    assert dup.data_end_col == "Q"
    assert dup.optional, "nessun campo e' letto dal VBA: la fonte puo' mancare"
    assert dup.read_by_row, "letto cella per cella: il limite non e' un intervallo"
    assert dup.key_field == "Full Name"
    assert dup.stop_values == ("Total",)
    assert "DC Dashboard" in dup.dependent_sheets

    for altro in contract.datasets.values():
        if altro.name != "DUP_DATASET":
            assert altro.header_row == 1, (
                f"{altro.name} ha le intestazioni in riga {altro.header_row}: "
                f"se e' voluto, questo test va aggiornato — ma va anche "
                f"controllato che i suoi consumatori se ne siano accorti."
            )


def test_le_due_colonne_data_dei_duplicati_hanno_il_formato_dichiarato(contract):
    """Senza `date_format` non sarebbero nemmeno convertibili — e sarebbe la
    fortuna, non il progetto: `8/4/2026` si legge in due modi."""
    dup = contract.dataset("DUP_DATASET")
    per_nome = {f.canonical: f for f in dup.input_fields}
    for nome in ("Date/Time Opened", "Date/Time Closed"):
        f = per_nome[nome]
        assert f.dtype == "datetime", (
            f"{nome} come testo: Excel in locale italiano convertirebbe "
            f'"8/4/2026 3:59 PM" in data alla scrittura, e le formule del '
            f"template che lo parsavano con FIND(\"/\") darebbero #VALUE!"
        )
        assert f.date_format == "%m/%d/%Y %I:%M %p"


def test_reader_per_dataset(contract):
    """I 4 CSV passano da csvsource, i due fogli WFM dagli adattatori."""
    assert contract.dataset("AT_DATASET").reader == "csv"
    assert contract.dataset("Turni").reader == "wfm_roster"
    assert contract.dataset("Slot Only Cases").reader == "wfm_backoffice"


def test_reader_non_valido():
    raw = _with_fields([{"canonical": "A", "target_col": "A", "role": "input"}])
    raw["datasets"]["D"]["reader"] = "telepatia"
    with pytest.raises(ContractError) as e:
        parse_contract(raw)
    assert "reader=" in str(e.value)


def test_dataset_sconosciuto_elenca_i_disponibili(contract):
    with pytest.raises(ContractError) as e:
        contract.dataset("NON_ESISTE")
    assert "AT_DATASET" in str(e.value)


def test_due_campi_sulla_stessa_colonna_target():
    raw = _with_fields([
        {"canonical": "Alpha", "target_col": "B", "role": "input"},
        {"canonical": "Beta", "target_col": "B", "role": "input"},
    ])
    with pytest.raises(DuplicateTargetError) as e:
        parse_contract(raw)
    assert "stessa cella" in str(e.value)


def test_campo_canonico_duplicato():
    raw = _with_fields([
        {"canonical": "Alpha", "target_col": "A", "role": "input"},
        {"canonical": "Alpha", "target_col": "B", "role": "input"},
    ])
    with pytest.raises(ContractError):
        parse_contract(raw)


def test_role_non_valido():
    raw = _with_fields([{"canonical": "A", "target_col": "A", "role": "sconosciuto"}])
    with pytest.raises(ContractError) as e:
        parse_contract(raw)
    assert "role=" in str(e.value)


def test_dtype_non_valido():
    raw = _with_fields([{"canonical": "A", "target_col": "A", "role": "input", "dtype": "decimal"}])
    with pytest.raises(ContractError):
        parse_contract(raw)


def test_match_non_valido():
    raw = _with_fields([{"canonical": "A", "target_col": "A", "role": "input", "match": "fuzzy"}])
    with pytest.raises(ContractError) as e:
        parse_contract(raw)
    assert "match=" in str(e.value)


def test_campo_prima_di_data_start_col():
    raw = _with_fields(
        [{"canonical": "A", "target_col": "A", "role": "input"}], data_start_col="B"
    )
    with pytest.raises(ContractError) as e:
        parse_contract(raw)
    assert "prima di data_start_col" in str(e.value)


def test_derivato_con_source_inesistente():
    raw = _with_fields([
        {"canonical": "Alpha", "target_col": "A", "role": "input"},
        {"canonical": "Derivato", "target_col": "B", "role": "derived", "source": "Fantasma"},
    ])
    with pytest.raises(ContractError) as e:
        parse_contract(raw)
    assert "Fantasma" in str(e.value)


def test_dataset_senza_input():
    raw = _with_fields([
        {"canonical": "Derivato", "target_col": "A", "role": "derived"},
    ])
    with pytest.raises(ContractError) as e:
        parse_contract(raw)
    assert "role=input" in str(e.value)


def test_chiave_obbligatoria_mancante():
    raw = {
        "normalize": {},
        "datasets": {"D": {"sheet": "D", "header_row": 1, "data_start_col": "A"}},
    }
    with pytest.raises(ContractError) as e:
        parse_contract(raw)
    assert "fields" in str(e.value)


def test_contratto_senza_dataset():
    with pytest.raises(ContractError):
        parse_contract({"normalize": {}, "datasets": {}})


def test_file_inesistente(tmp_path):
    with pytest.raises(ContractError):
        load_contract(tmp_path / "niente.yml")


def test_last_input_index_ignora_i_derivati(contract):
    at = contract.dataset("AT_DATASET")
    # L = 12 e' l'ultimo input; P/Q (16/17) sono derivate.
    assert index_to_col(at.last_input_index) == "L"
