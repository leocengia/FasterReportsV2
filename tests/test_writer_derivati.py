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
from fasterreports.omni.writer import (
    append_helper_casetype,
    leggi_coppie_helper,
    write_aht_history,
)


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


def test_appende_dalla_prima_riga_libera():
    """Le righe 2..63 sono la lista curata a cui punta 'CaseType Deepdive': la
    scrittura deve cominciare a 64, non prima."""
    sht = helper()
    res = append_helper_casetype(
        FintoBook({"Helper CaseType": sht}),
        [("Phone", "Call Assignment"), ("Non-live", "Collections")],
    )
    (addr, valori), = sht.scritture
    assert addr == "A64:B65"
    assert valori == [["Phone", "Call Assignment"], ["Non-live", "Collections"]]
    assert res.rows_written == 2


def test_non_pulisce_mai_niente():
    """Non c'e' nessun `clear` in questo percorso: tutto il resto del foglio sono
    formule del template."""
    sht = helper()
    append_helper_casetype(FintoBook({"Helper CaseType": sht}), [("Phone", "X")])
    assert sht.pulizie == []


def test_scrive_solo_le_colonne_A_e_B():
    sht = helper()
    append_helper_casetype(FintoBook({"Helper CaseType": sht}), [("Phone", "X")])
    (addr, _), = sht.scritture
    assert addr.startswith("A") and ":B" in addr


def test_niente_da_appendere_niente_scritture():
    sht = helper()
    res = append_helper_casetype(FintoBook({"Helper CaseType": sht}), [])
    assert sht.scritture == []
    assert res.rows_written == 0


def test_si_ferma_dove_finiscono_le_formule_e_lo_dice():
    """Una coppia scritta oltre la capienza avrebbe nome e canale e nessun numero
    accanto: presente e invisibile insieme, cioe' lo stesso difetto che questo
    passaggio serve a togliere."""
    sht = helper(ultima_b=63, ultima_d=65)  # spazio per 2 righe: 64 e 65
    nuove = [("Phone", f"Tipo {i}") for i in range(5)]
    res = append_helper_casetype(FintoBook({"Helper CaseType": sht}), nuove)

    (addr, valori), = sht.scritture
    assert addr == "A64:B65"
    assert len(valori) == 2
    assert res.rows_written == 2
    assert res.warnings and "3 case type NON" in res.warnings[0]
    assert "riga 65" in res.warnings[0]


def test_template_non_ancora_esteso_non_scrive_e_spiega_come_rimediare():
    """Lo stato del template WIP al 2026-08-18: formule e lista finiscono entrambe
    a riga 63, quindi non c'e' una singola riga in cui un case type nuovo verrebbe
    calcolato. Il build NON scrive e dice cosa fare — scrivere comunque
    produrrebbe righe con nome e canale e nessun numero accanto."""
    sht = helper(ultima_b=63, ultima_d=63)
    res = append_helper_casetype(FintoBook({"Helper CaseType": sht}), [("Phone", "X")])
    assert sht.scritture == []
    assert res.rows_written == 0
    assert "riga 63" in res.warnings[0]
    assert "trascina le formule" in res.warnings[0].lower()


def test_foglio_assente_avvisa_ma_non_ferma_il_report():
    """E' un foglio di supporto a un report secondario: il resto dell'Omni Report
    e' valido, e fermarlo per questo sarebbe sproporzionato."""
    res = append_helper_casetype(FintoBook({"Altro": FintoFoglio()}), [("Phone", "X")])
    assert res.rows_written == 0
    assert "Helper CaseType" in res.warnings[0]


# ---------------------------------------------------------------------------
# Rilettura della lista curata
# ---------------------------------------------------------------------------


def test_legge_le_coppie_esistenti():
    sht = FintoFoglio(
        ultime={"B": 4},
        valori={"A2:B4": [["Phone", "EVC"], ["Phone", "Alfa"], ["Non-live", "EVC"]]},
    )
    assert leggi_coppie_helper(FintoBook({"Helper CaseType": sht})) == [
        ("Phone", "EVC"), ("Phone", "Alfa"), ("Non-live", "EVC")
    ]


def test_una_riga_sola_non_si_appiattisce():
    """xlwings restituisce una lista piatta quando l'intervallo e' di una riga:
    senza il caso speciale, 'Phone' ed 'EVC' diventerebbero due righe rotte."""
    sht = FintoFoglio(ultime={"B": 2}, valori={"A2:B2": ["Phone", "EVC"]})
    assert leggi_coppie_helper(FintoBook({"Helper CaseType": sht})) == [("Phone", "EVC")]


def test_lista_vuota_se_il_foglio_ha_solo_lintestazione():
    sht = FintoFoglio(ultime={"B": 1})
    assert leggi_coppie_helper(FintoBook({"Helper CaseType": sht})) == []


def test_salta_le_righe_a_meta():
    sht = FintoFoglio(
        ultime={"B": 4},
        valori={"A2:B4": [["Phone", "EVC"], [None, "Alfa"], ["Non-live", None]]},
    )
    assert leggi_coppie_helper(FintoBook({"Helper CaseType": sht})) == [("Phone", "EVC")]
