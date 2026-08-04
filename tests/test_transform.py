from __future__ import annotations

from datetime import datetime

import pytest

from fasterreports.core.csvsource import read_csv_text
from fasterreports.core.errors import CoercionError, EmptyDatasetError
from fasterreports.core.matcher import resolve_dataset
from fasterreports.core.transform import add_derived, build_block

AT_HEADERS = [
    "Agent Email", "Agent State", "Number of Active Contacts",
    "Start Time", "Total Time in seconds", "Productive Aux Flag (Yes / No)",
]


def _at_block(contract, rows, **kw):
    ds = contract.dataset("AT_DATASET")
    text = ",".join(AT_HEADERS) + "\n" + "\n".join(",".join(r) for r in rows) + "\n"
    src = read_csv_text(text)
    mapping = resolve_dataset(ds, list(src.headers), contract.normalize)
    return ds, build_block(ds, mapping, src.rows(), **kw)


def test_blocco_allineato_alle_lettere(contract):
    ds, block = _at_block(
        contract,
        [["mario@x.it", "Break", "0", "2026-07-28 06:00:00", "300", "No"]],
    )
    assert (block.start_col, block.end_col) == ("B", "L")
    riga = block.rows[0]
    # offset 0 = B; F=4, H=6, I=7, K=9, L=10
    assert riga[0] == "mario@x.it"
    assert riga[4] == "Break"
    assert riga[6] == 0.0
    assert riga[7] == datetime(2026, 7, 28, 6, 0)
    assert riga[9] == 300.0
    assert riga[10] == "No"
    # C, D, E, G, J restano vuote: nessuna formula le legge.
    assert [riga[i] for i in (1, 2, 3, 5, 8)] == [None] * 5


def test_derived_non_finisce_nel_blocco(contract):
    # In modalita' 'formula' P/Q restano formule del template: scriverci valori
    # le distruggerebbe.
    _, block = _at_block(
        contract, [["a@x.it", "Login", "1", "2026-07-28 06:00:00", "10", "Yes"]]
    )
    assert block.end_col == "L"  # non P ne' Q


def test_dataset_vuoto_blocca(contract):
    with pytest.raises(EmptyDatasetError) as e:
        _at_block(contract, [])
    assert "nessun dato utile" in str(e.value)


def test_troppi_valori_non_coercibili_bloccano(contract):
    rows = [
        ["a@x.it", "Break", "0", "2026-07-28 06:00:00", "non un numero", "No"],
        ["b@x.it", "Break", "0", "2026-07-28 07:00:00", "nemmeno", "No"],
    ]
    with pytest.raises(CoercionError) as e:
        _at_block(contract, rows, max_uncoercible_ratio=0.02)
    msg = str(e.value)
    assert "Total Time in seconds" in msg
    assert "non un numero" in msg  # gli esempi rifiutati sono nel messaggio
    assert "mappata su quella sbagliata" in msg


def test_pochi_valori_non_coercibili_passano_ma_sono_contati(contract):
    rows = [["a@x.it", "Break", "0", "2026-07-28 06:00:00", "10", "No"]] * 99
    rows.append(["b@x.it", "Break", "0", "2026-07-28 06:00:00", "boh", "No"])
    _, block = _at_block(contract, rows, max_uncoercible_ratio=0.02)
    st = block.stats["Total Time in seconds"]
    assert (st.total, st.bad) == (100, 1)
    assert st.bad_ratio == pytest.approx(0.01)


def test_le_celle_vuote_non_contano_come_scarti(contract):
    rows = [["a@x.it", "Break", "0", "2026-07-28 06:00:00", "", "No"]] * 10
    _, block = _at_block(contract, rows, max_uncoercible_ratio=0.0)
    st = block.stats["Total Time in seconds"]
    assert (st.blank, st.bad, st.bad_ratio) == (10, 0, 0.0)


def test_riga_piu_corta_dellheader_non_esplode(contract):
    ds = contract.dataset("ATwi_DATASET")
    text = "Agent Email,Talk Time in seconds,Wrap-up Time in seconds,Handle Time in seconds,Initiation Method\n"
    text += "a@x.it,10\n"  # riga troncata, capita negli export interrotti
    src = read_csv_text(text)
    mapping = resolve_dataset(ds, list(src.headers), contract.normalize)
    block = build_block(ds, mapping, src.rows())
    assert block.rows[0][1] == "a@x.it"
    assert block.rows[0][11] is None  # M, mancante nella riga


# --- modalita' derived=python -----------------------------------------------

def test_add_derived_replica_le_formule_del_template(contract):
    ds, block = _at_block(
        contract, [["a@x.it", "Login", "1", "2026-07-28 21:00:00", "60", "Yes"]]
    )
    out = add_derived(ds, block, offset_hours=9.0)
    assert out.end_col == "Q"
    riga = out.rows[0]
    p, q = riga[14], riga[15]  # P e Q, offset da B
    # 28/07 21:00 + 9h = 29/07 06:00 -> P = giorno, Q = 0.25
    assert q == pytest.approx(0.25)
    assert p == int(p)
    from fasterreports.core.coerce import EXCEL_EPOCH
    from datetime import timedelta
    assert EXCEL_EPOCH + timedelta(days=p) == datetime(2026, 7, 29)


def test_add_derived_lascia_vuoto_se_start_time_manca(contract):
    ds, block = _at_block(contract, [["a@x.it", "Login", "1", "", "60", "Yes"]])
    out = add_derived(ds, block, offset_hours=9.0)
    # Le formule del template fanno IF($I2="","",...): stesso comportamento.
    assert out.rows[0][14] is None and out.rows[0][15] is None
