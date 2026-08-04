from __future__ import annotations

from datetime import datetime

import pytest

from fasterreports.core.coerce import (
    Uncoercible,
    to_datetime,
    to_excel_fraction,
    to_float,
    to_int,
    to_str,
)


@pytest.mark.parametrize(
    "src,expected",
    [
        ("120", 120.0),
        ("1.5", 1.5),
        ("1,5", 1.5),          # export con locale italiano
        ("-3", -3.0),
        (" 42 ", 42.0),
        ("1e3", 1000.0),
        ("12,5%", 0.125),      # Co-Browse Usage % arriva anche cosi'
        ("0.42", 0.42),
        (7, 7.0),
    ],
)
def test_to_float(src, expected):
    assert to_float(src) == pytest.approx(expected)


@pytest.mark.parametrize("src", ["", "   ", None])
def test_vuoto_e_none_non_zero(src):
    # Distinzione importante: una cella vuota non e' uno zero. Scriverci 0
    # falserebbe le medie del motore.
    assert to_float(src) is None
    assert to_str(src) is None
    assert to_datetime(src) is None


@pytest.mark.parametrize("src", ["abc", "N/A", "12.34.56", "1.234,56", "--3", "1 2"])
def test_to_float_rifiuta(src):
    # '1.234,56' ha due separatori: indovinare quale sia il decimale
    # produrrebbe un numero plausibile e sbagliato.
    with pytest.raises(Uncoercible):
        to_float(src)


def test_to_int_rifiuta_i_decimali():
    assert to_int("5") == 5
    with pytest.raises(Uncoercible):
        to_int("5.5")


@pytest.mark.parametrize(
    "src,expected",
    [
        ("2026-07-28 14:30:00", datetime(2026, 7, 28, 14, 30)),
        ("2026-07-28", datetime(2026, 7, 28)),
        ("28/07/2026 14:30", datetime(2026, 7, 28, 14, 30)),
        ("2026-07-28T14:30:00Z", datetime(2026, 7, 28, 14, 30)),
    ],
)
def test_to_datetime(src, expected):
    assert to_datetime(src) == expected


def test_seriale_excel():
    # 46229.3124 e' un valore reale letto da ATwi_DATASET!E2 del W30
    # ('Connected to Agent Time'), quindi cade davvero nella W30 del 2026.
    got = to_datetime(46229.312407407408)
    assert got == datetime(2026, 7, 26, 7, 29, 52)
    # Andata e ritorno sul seriale 1 = 1899-12-31.
    assert to_datetime(1) == datetime(1899, 12, 31)


def test_to_datetime_rifiuta_il_testo():
    with pytest.raises(Uncoercible):
        to_datetime("ieri mattina")


def test_frazione_del_giorno():
    # E' la forma di 'Ora Milano' (AT_DATASET!Q).
    assert to_excel_fraction(datetime(2026, 1, 1, 6, 0, 0)) == pytest.approx(0.25)
    assert to_excel_fraction(datetime(2026, 1, 1, 12, 0, 0)) == pytest.approx(0.5)
