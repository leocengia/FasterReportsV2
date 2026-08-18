"""`8/10/2026`: 10 agosto o 8 ottobre?

Bug vero, trovato il 2026-08-18 aggiungendo `Date Viewpoint` al contratto:
`_DATETIME_FORMATS` provava il formato europeo prima di quello americano, quindi
la data della W33 dell'export Salesforce veniva letta come 8 ottobre — settimana
ISO 41 invece di 33. Nessun errore, nessun avviso: solo uno storico che sarebbe
finito otto settimane piu' avanti.

Il rimedio ha due meta', e servono entrambe:
  - `date_format` nel contratto: dove la lettura conta, si dichiara, non si indovina;
  - `data_ambigua`: dove non e' dichiarata, il preflight lo dice, invece di
    lasciare passare un lancio di moneta travestito da dato.
"""

from __future__ import annotations

from datetime import datetime

import pytest

from fasterreports.core.coerce import Uncoercible, coerce, data_ambigua, to_datetime
from fasterreports.core.contract import ContractError, load_contract
from fasterreports.core.errors import PipelineError  # noqa: F401  (usato dai test a valle)

AMERICANO = "%m/%d/%Y %I:%M:%S %p"


# ---------------------------------------------------------------------------
# Il formato dichiarato vince, e vince sempre
# ---------------------------------------------------------------------------


def test_col_formato_dichiarato_la_data_e_quella_giusta():
    v = to_datetime("8/10/2026 12:00:00 AM", AMERICANO)
    assert (v.year, v.month, v.day) == (2026, 8, 10)
    assert v.date().isocalendar()[1] == 33


def test_senza_formato_dichiarato_vince_il_primo_che_combacia():
    """Il comportamento storico, tenuto per non cambiare la lettura dei campi che
    oggi funzionano. E' registrato qui apposta: e' il motivo per cui una colonna
    ambigua DEVE dichiarare il formato."""
    v = to_datetime("8/10/2026 12:00:00 AM")
    assert (v.month, v.day) == (10, 8)  # letta all'europea


def test_il_formato_dichiarato_non_ripiega_sugli_altri():
    """Se l'export cambia forma, l'errore deve uscire subito. Ripiegare sulla
    lista dei formati noti rimetterebbe in gioco proprio l'ambiguita' che il
    formato dichiarato serve a togliere."""
    with pytest.raises(Uncoercible) as e:
        to_datetime("2026-08-10", AMERICANO)
    assert AMERICANO in str(e.value)


def test_coerce_passa_il_formato_solo_al_datetime():
    assert coerce("8/10/2026 12:00:00 AM", "datetime", AMERICANO).day == 10
    # Per gli altri dtype il formato non c'entra e non deve dare fastidio.
    assert coerce(" ciao ", "str") == "ciao"
    assert coerce("12,5", "float") == pytest.approx(12.5)


def test_valori_gia_datetime_passano_indenni():
    d = datetime(2026, 8, 10)
    assert to_datetime(d, AMERICANO) is d


def test_vuoto_resta_vuoto():
    assert to_datetime("", AMERICANO) is None
    assert to_datetime(None, AMERICANO) is None


# ---------------------------------------------------------------------------
# Riconoscere l'ambiguita': ne' troppo, ne' troppo poco
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("v", ["8/10/2026", "8/10/2026 12:00:00 AM", "1/2/26", "12-11-2026"])
def test_ambigue(v):
    assert data_ambigua(v)


@pytest.mark.parametrize(
    "v",
    [
        "2026-08-10 12:00:00",  # anno in testa: nessun dubbio
        "2026/08/10",
        "25/12/2026",           # 25 non puo' essere un mese
        "12/25/2026",
        "5/5/2026",             # stessa data in entrambe le letture
        "",
        None,
        45884,                  # seriale Excel
        datetime(2026, 8, 10),
    ],
)
def test_non_ambigue(v):
    assert not data_ambigua(v)


def test_ambiguita_non_dipende_dalla_parte_orario():
    assert data_ambigua("3/4/2026") == data_ambigua("3/4/2026 23:59:59")


# ---------------------------------------------------------------------------
# Il contratto: dichiararlo dove serve, e non dove non ha senso
# ---------------------------------------------------------------------------


def test_date_format_solo_su_dtype_datetime(tmp_path):
    p = tmp_path / "c.yml"
    p.write_text(
        "datasets:\n"
        "  X:\n"
        "    sheet: X\n"
        "    header_row: 1\n"
        "    data_start_col: A\n"
        "    fields:\n"
        "      - canonical: 'Q'\n"
        "        target_col: A\n"
        "        role: input\n"
        "        dtype: str\n"
        '        date_format: "%m/%d/%Y"\n',
        encoding="utf-8",
    )
    with pytest.raises(ContractError) as e:
        load_contract(p)
    assert "dtype=datetime" in str(e.value)


def test_date_format_senza_direttive_e_un_errore(tmp_path):
    """`date_format: americano` non e' un formato: e' un'intenzione. Meglio
    fermarsi al caricamento del contratto che leggere le date a caso."""
    p = tmp_path / "c.yml"
    p.write_text(
        "datasets:\n"
        "  X:\n"
        "    sheet: X\n"
        "    header_row: 1\n"
        "    data_start_col: A\n"
        "    fields:\n"
        "      - canonical: 'Q'\n"
        "        target_col: A\n"
        "        role: input\n"
        "        dtype: datetime\n"
        "        date_format: americano\n",
        encoding="utf-8",
    )
    with pytest.raises(ContractError) as e:
        load_contract(p)
    assert "strptime" in str(e.value)


def test_il_contratto_reale_dichiara_il_formato_di_date_viewpoint():
    """La riga che impedisce alla W33 di finire nella W41."""
    from pathlib import Path

    c = load_contract(Path(__file__).resolve().parents[1] / "config" / "columns.yml")
    fld = next(
        f for f in c.dataset("SF_DATABASE").input_fields if f.canonical == "Date Viewpoint"
    )
    assert fld.dtype == "datetime"
    assert fld.date_format, "senza date_format la settimana viene letta al contrario"
    letta = coerce("8/10/2026 12:00:00 AM", "datetime", fld.date_format)
    assert letta.date().isocalendar()[:2] == (2026, 33)
    assert letta.weekday() == 0, "Date Viewpoint e' il lunedi' della settimana"
