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
        "AT_DATASET", "ATwi_DATASET", "SF_DATABASE", "PSAT_DATASET"
    }
    at = contract.dataset("AT_DATASET")
    assert at.list_object is None
    assert at.max_template_row == 130000
    assert contract.dataset("SF_DATABASE").list_object == "AHT_Data"


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
