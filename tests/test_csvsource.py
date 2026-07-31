from __future__ import annotations

import pytest

from fasterreports.core.csvsource import read_csv, sniff_delimiter, sniff_encoding
from fasterreports.core.errors import SourceError


@pytest.mark.parametrize(
    "line,expected",
    [
        ("a,b,c", ","),
        ("a;b;c", ";"),
        ("a\tb\tc", "\t"),
        ("a|b|c", "|"),
        # Testo libero con virgole dentro le virgolette: il separatore e' ';'.
        ('a;b;"ottimo, davvero; grazie"', ";"),
    ],
)
def test_sniff_delimiter(line, expected):
    assert sniff_delimiter(line) == expected


def test_sniff_delimiter_su_una_sola_colonna():
    with pytest.raises(SourceError) as e:
        sniff_delimiter("solo_una_colonna")
    assert "separatore" in str(e.value)


def test_bom_rimosso(tmp_path):
    p = tmp_path / "bom.csv"
    p.write_bytes("﻿Agent Email,Foo\na@x.it,1\n".encode("utf-8"))
    src = read_csv(p)
    assert src.encoding == "utf-8-sig"
    assert src.headers[0] == "Agent Email"


def test_cp1252(tmp_path):
    p = tmp_path / "w.csv"
    p.write_bytes("Nome,Città\nMario,Torino\n".encode("cp1252"))
    src = read_csv(p)
    assert src.headers == ["Nome", "Città"]


def test_encoding_rilevato(tmp_path):
    p = tmp_path / "u.csv"
    p.write_text("a,b\n1,2\n", encoding="utf-8")
    assert sniff_encoding(p) == "utf-8-sig"  # compatibile, mangia il BOM se c'e'


def test_righe_vuote_ignorate(tmp_path):
    p = tmp_path / "v.csv"
    p.write_text("a,b\n1,2\n\n,\n3,4\n", encoding="utf-8")
    assert list(read_csv(p).rows()) == [["1", "2"], ["3", "4"]]


def test_file_mancante(tmp_path):
    with pytest.raises(SourceError) as e:
        read_csv(tmp_path / "assente.csv")
    assert "non trovato" in str(e.value)


def test_file_vuoto(tmp_path):
    p = tmp_path / "e.csv"
    p.write_text("", encoding="utf-8")
    with pytest.raises(SourceError) as e:
        read_csv(p)
    assert "vuoto" in str(e.value)


def test_prima_riga_vuota(tmp_path):
    p = tmp_path / "h.csv"
    p.write_text("\n\na,b\n", encoding="utf-8")
    with pytest.raises(SourceError) as e:
        read_csv(p)
    assert "intestazioni" in str(e.value)


def test_header_con_spazi_ripuliti(tmp_path):
    p = tmp_path / "s.csv"
    p.write_text("  Agent Email  , Foo \na@x.it,1\n", encoding="utf-8")
    assert read_csv(p).headers == ["Agent Email", "Foo"]


def test_streaming_non_carica_tutto(tmp_path):
    # AT_DATASET ha ~130k righe: le righe devono restare un iteratore.
    p = tmp_path / "big.csv"
    p.write_text("a,b\n" + "".join(f"{i},{i}\n" for i in range(1000)), encoding="utf-8")
    rows = read_csv(p).rows()
    assert next(rows) == ["0", "0"]
    assert sum(1 for _ in rows) == 999
