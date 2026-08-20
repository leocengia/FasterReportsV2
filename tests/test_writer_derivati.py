"""Le due scritture derivate, provate con un finto workbook.

`writer.py` parla con xlwings e qui Excel non c'e', quindi la prova vera resta
quella sulla macchina di Leonardo. Ma cio' che questi test coprono non e' xlwings:
e' *quali intervalli* vengono toccati e quali no — ed e' esattamente il punto
delicato, perche' `CaseType Deepdive` punta alle righe di 'Helper CaseType' per
posizione e ha formattazione fatta a mano. Una scrittura una riga piu' su
metterebbe i case type sotto le etichette sbagliate senza nessun errore.

Il doppio e' volutamente stupido: registra le chiamate, non simula Excel.
"""

from __future__ import annotations

import pytest

from fasterreports.core.aht_history import RigaStorico, righe_foglio
from fasterreports.core.errors import PipelineError
from fasterreports.omni.writer import write_aht_history


class FintoRange:
    def __init__(self, sht, addr):
        self.sht, self.addr = sht, addr
        self._nf = sht.formati.get(addr, "")

    # --- lettura --------------------------------------------------------
    @property
    def value(self):
        return self.sht.valori.get(self.addr)

    @value.setter
    def value(self, v):
        self.sht.scritture.append((self.addr, v))
        self.sht.valori[self.addr] = v

    @property
    def number_format(self):
        return self._nf

    @number_format.setter
    def number_format(self, v):
        self.sht.formati_scritti.append((self.addr, v))

    def clear_contents(self):
        self.sht.pulizie.append(self.addr)

    def end(self, direction):
        assert direction == "up"
        col = "".join(c for c in self.addr if c.isalpha())
        return FintoCella(self.sht.ultima_riga_per_colonna.get(col, 1))


class FintoCella:
    def __init__(self, row):
        self.row = row


class FintoFoglio:
    def __init__(self, ultime=None, formati=None, valori=None, ultima_usata=1):
        self.scritture: list = []
        self.pulizie: list[str] = []
        self.formati_scritti: list = []
        self.formati = formati or {}
        self.valori = valori or {}
        self.ultima_riga_per_colonna = ultime or {}
        self._ultima_usata = ultima_usata

    def range(self, addr):
        return FintoRange(self, addr)

    @property
    def used_range(self):
        return type("R", (), {"last_cell": FintoCella(self._ultima_usata)})()

    @property
    def cells(self):
        return type("R", (), {"last_cell": FintoCella(1048576)})()


class FintoBook:
    def __init__(self, fogli):
        self._fogli = fogli

    @property
    def sheets(self):
        book = self

        class Sheets:
            def __getitem__(self, nome):
                if nome not in book._fogli:
                    raise KeyError(nome)
                return book._fogli[nome]

            def __iter__(self):
                return iter(
                    type("S", (), {"name": n})() for n in book._fogli
                )

        return Sheets()


STORICO = [
    RigaStorico(2026, 32, "Phone", "EVC", 10, 12.0),
    RigaStorico(2026, 33, "Phone", "EVC", 12, 13.0),
]


# ---------------------------------------------------------------------------
# AHT History
# ---------------------------------------------------------------------------


def test_scrive_intestazione_e_dati_in_un_colpo():
    sht = FintoFoglio(formati={"A2": '"W"0', "E2": "0.00"}, ultima_usata=3)
    res = write_aht_history(FintoBook({"AHT History": sht}), righe_foglio(STORICO))

    (addr, valori), = sht.scritture
    assert addr == "A1:G3"
    assert valori[0][0] == "week"          # intestazione, colonna A
    assert valori[0][5:] == ["iso_year", "week_key"]
    assert res.rows_written == 2
    assert res.range_written == "A2:G3"


def test_pulisce_oltre_i_dati_nuovi():
    """Uno storico piu' corto di prima (puo' capitare correggendo il CSV a mano)
    lascerebbe righe vecchie in fondo, e le formule del trend leggono fino a riga
    100000: le vedrebbero."""
    sht = FintoFoglio(ultima_usata=900)
    write_aht_history(FintoBook({"AHT History": sht}), righe_foglio(STORICO))
    assert sht.pulizie == ["A1:G900"]


def test_propaga_i_formati_letti_dalla_riga_2():
    """Il formato vive nel template, non nel codice: si legge da riga 2 e si
    estende. Cosi' cambiare `"W"0` nel template continua a funzionare."""
    sht = FintoFoglio(formati={"A2": '"W"0', "E2": "0.00"}, ultima_usata=3)
    write_aht_history(FintoBook({"AHT History": sht}), righe_foglio(STORICO))

    scritti = dict(sht.formati_scritti)
    assert scritti["A3:A3"] == '"W"0'
    assert scritti["E3:E3"] == "0.00"


def test_non_propaga_niente_se_ce_solo_lintestazione():
    sht = FintoFoglio(ultima_usata=1)
    write_aht_history(FintoBook({"AHT History": sht}), righe_foglio([]))
    assert sht.formati_scritti == []


def test_foglio_assente_dice_cosa_manca_e_perche():
    with pytest.raises(PipelineError) as e:
        write_aht_history(FintoBook({"Altro": FintoFoglio()}), righe_foglio(STORICO))
    assert "AHT History" in str(e.value)
    assert "AHT Trend WoW" in str(e.value)


# ---------------------------------------------------------------------------
# Helper CaseType: appende, e non tocca mai le righe che ci sono
# ---------------------------------------------------------------------------


def helper(ultima_b=63, ultima_d=163):
    """Un 'Helper CaseType' con la lista curata fino a riga 63 (colonna B) e le
    formule preparate fino a 163 (colonna D) — cioe' il template DOPO l'estensione
    descritta in docs/template-modifiche-wip.md §2."""
    return FintoFoglio(ultime={"B": ultima_b, "D": ultima_d})

