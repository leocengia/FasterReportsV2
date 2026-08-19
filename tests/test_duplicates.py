"""La sezione Duplicate Cases: dal download al blocco.

`DUP_DATASET` e' il primo dataset che non e' una tabella pulita, e il primo che un
foglio consumatore legge CELLA PER CELLA. Questi test fissano le tre cose che ne
derivano e che nessun altro dataset esercita:

- le intestazioni stanno in riga 14, sotto il preambolo del report Salesforce;
- la coda dei totali (`Total | Sum`, `Count`) non e' un record;
- l'offset fra `DUP_DATASET` e `Duplicates Helper` e' fisso, quindi una riga
  inserita nel preambolo del TEMPLATE sposterebbe tutti i dati di uno senza
  produrre un solo errore.
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

import pytest

from fasterreports.core.matcher import resolve_dataset
from fasterreports.core.tablesource import read_table
from fasterreports.core.transform import build_block

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "template" / "Omni_Report_TEMPLATE.xlsm"

sys.path.insert(0, str(ROOT / "tools"))

from test_tablesource import _dup_export  # noqa: E402  — l'export sintetico


def _blocco(contract, path):
    ds = contract.dataset("DUP_DATASET")
    src = read_table(path)
    mapping = resolve_dataset(ds, list(src.headers), contract.normalize)
    return ds, src, build_block(ds, mapping, src.rows())


def test_il_contratto_aggancia_l_export_sintetico(contract, tmp_path):
    """Le 15 colonne del report si agganciano alle lettere che il contratto dice."""
    ds, _src, block = _blocco(contract, _dup_export(tmp_path, righe=3))
    assert block.n_rows == 3
    assert block.start_col == "B" and block.end_col == "Q"


def test_la_coda_dei_totali_non_diventa_un_agente(contract, tmp_path):
    """`Total | Sum | 72,13` e `Count | 178` sono la coda del report, non dati.

    Senza `key_field`/`stop_values` uscirebbero due agenti: uno di nome 'Total' e
    uno senza nome. Il template si difende da solo
    (`IF(OR(B15="",B15="Total"),"",...)`), ma scriverle sarebbe sbagliato
    comunque — e soprattutto il conteggio delle righe direbbe 5 invece di 3.
    """
    ds, _src, block = _blocco(contract, _dup_export(tmp_path, righe=3, coda=True))
    assert block.n_rows == 3
    agenti = [r[0] for r in block.rows]
    assert agenti == ["Agente 0", "Agente 1", "Agente 2"]
    assert "Total" not in agenti
    # E lo dice, invece di scartare in silenzio.
    assert block.dropped
    assert any("chiusura" in d for d in block.dropped)


def test_un_export_senza_coda_non_ha_bisogno_di_niente_di_diverso(contract, tmp_path):
    ds, _src, block = _blocco(contract, _dup_export(tmp_path, righe=3, coda=False))
    assert block.n_rows == 3
    assert block.dropped == []


def test_le_date_diventano_datetime_veri(contract, tmp_path):
    """`8/4/2026 3:59 PM` -> 4 agosto, non 8 aprile.

    Senza `date_format` non sarebbe nemmeno convertibile: la lista dei formati
    noti non copre `%I:%M %p` senza secondi. Con il formato dichiarato la lettura
    e' decisa, che e' la lezione di `Date Viewpoint`.
    """
    from datetime import datetime

    ds, _src, block = _blocco(contract, _dup_export(tmp_path, righe=1))
    per_col = {f.canonical: f.target_index - ds.start_index for f in ds.input_fields}
    aperto = block.rows[0][per_col["Date/Time Opened"]]
    chiuso = block.rows[0][per_col["Date/Time Closed"]]
    assert aperto == datetime(2026, 8, 4, 15, 59)
    assert chiuso == datetime(2026, 8, 4, 16, 24)
    # Nessuna riga segnalata come ambigua: il formato e' dichiarato.
    assert block.stats["Date/Time Opened"].ambigue == 0
    assert block.stats["Date/Time Opened"].bad == 0


# ---------------------------------------------------------------------------
# Il vincolo col template: l'offset fisso
# ---------------------------------------------------------------------------


@pytest.mark.skipif(not TEMPLATE.is_file(), reason="template assente")
def test_offset_helper_duplicates(contract):
    """`Duplicates Helper!A2` deve leggere `DUP_DATASET!B15`, e non altro.

    E' il guasto piu' insidioso di questa sezione, ed e' il solo che sposterebbe
    TUTTI i dati senza produrre un errore: 'Duplicates Helper' legge cella per
    cella con offset fisso, quindi una riga inserita nel preambolo di
    `DUP_DATASET` farebbe leggere a ogni formula la riga sbagliata. Numeri
    plausibili, attribuiti alla persona sbagliata.

    Il test lega i due capi: la riga dichiarata nel contratto (`header_row`) e la
    riga che la prima formula del foglio consumatore va davvero a leggere.
    """
    ds = contract.dataset("DUP_DATASET")
    with zipfile.ZipFile(TEMPLATE) as z:
        wb = z.read("xl/workbook.xml").decode("utf8", "replace")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8", "replace")
        relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
        target = None
        for m in re.finditer(r"<sheet ([^>]*)/>", wb):
            nome = re.search(r'name="([^"]*)"', m.group(1))
            rid = re.search(r'r:id="(rId\d+)"', m.group(1))
            if nome and rid and nome.group(1) == "Duplicates Helper":
                target = "xl/" + relmap[rid.group(1)].lstrip("/").removeprefix("xl/")
        assert target, "il template non ha il foglio 'Duplicates Helper'"
        raw = z.read(target).decode("utf8", "replace")

    cella = re.search(r'<c r="A2"[^>]*>(.*?)</c>', raw, re.S)
    assert cella, "'Duplicates Helper'!A2 non ha contenuto"
    prima_riga_letta = int(re.search(r"DUP_DATASET!\$?[A-Z]{1,2}\$?(\d+)", cella.group(1)).group(1))
    assert prima_riga_letta == ds.header_row + 1, (
        f"'Duplicates Helper'!A2 legge DUP_DATASET riga {prima_riga_letta}, ma il "
        f"contratto scrive i dati dalla riga {ds.header_row + 1}. Qualcuno ha "
        f"inserito o togliato una riga nel preambolo di DUP_DATASET: tutti i dati "
        f"della sezione duplicati risulterebbero spostati di "
        f"{prima_riga_letta - ds.header_row - 1} righe, senza nessun errore."
    )


@pytest.mark.skipif(not TEMPLATE.is_file(), reason="template assente")
def test_la_capienza_di_dup_dataset_si_legge_dal_template(contract):
    """1000 righe, e il numero viene dal template — non da una costante qui.

    `scan_row_limits` non lo vedrebbe: cerca intervalli, e qui ci sono
    riferimenti a cella singola. Per questo `DUP_DATASET` dichiara `read_by_row`.
    """
    from fasterreports.core.templatescan import binding_limits, scan_cell_refs

    ds = contract.dataset("DUP_DATASET")
    assert ds.read_by_row
    limiti = scan_cell_refs(TEMPLATE, ("DUP_DATASET",))
    assert limiti, "nessun riferimento a cella trovato: il foglio consumatore c'e'?"
    vincolante = binding_limits(limiti)["DUP_DATASET"]
    assert vincolante.max_row == 1014
    assert vincolante.sheet == "Duplicates Helper"
    # 1014 - 14 (l'intestazione) = 1000 righe dati.
    assert vincolante.max_row - ds.header_row == 1000


def test_un_riferimento_a_riga_fissa_diventerebbe_un_limite_falso(tmp_path):
    """Perche' `read_by_row` e' per-dataset e non una scansione globale.

    Una formula che punta a UNA riga scelta a mano — il caso classico e' un
    "commento della settimana" preso da una riga fissa — non e' un limite di
    capienza: e' una selezione. Applicare la scansione a cella singola a tutti i
    dataset la leggerebbe come limite, e `_check_row_limits` bloccherebbe ogni
    settimana con piu' righe di quella: un guasto inventato dal controllo, che e'
    il tipo peggiore.

    Nota: nel template di oggi un riferimento cosi' NON c'e' (le formule di
    'Recap PSAT Positive' usano intervalli fino a riga 1000; la nota storica in
    columns.yml su `DQ130` descrive un workbook precedente). Quindi questo test
    lo costruisce: la proprieta' da fissare e' dello scanner, non di questo file.
    """
    from xlsxbuild import make_workbook

    from fasterreports.core.templatescan import scan_cell_refs, scan_row_limits

    p = make_workbook(
        tmp_path / "riga_fissa.xlsx",
        [
            ("Recap", {
                "A1": "=PSAT_DATASET!$DQ$130",                     # selezione a mano
                "A2": "=COUNTA(PSAT_DATASET!$DP$2:$DP$5000)",      # vero limite
            }),
            ("PSAT_DATASET", {"A1": "x"}),
        ],
    )
    # La scansione per intervalli vede il limite vero, e ignora la riga fissa.
    assert [l.max_row for l in scan_row_limits(p, ("PSAT_DATASET",))] == [5000]
    # Quella a cella singola vede 130 — ed e' per questo che non si applica a
    # tutti: qui produrrebbe un limite di 130 dove il vero e' 5000.
    assert [l.max_row for l in scan_cell_refs(p, ("PSAT_DATASET",))] == [130]


def test_solo_dup_dataset_dichiara_read_by_row(contract):
    """Il flag va messo solo dove il foglio consumatore legge davvero cosi'.

    Vedi il test qui sopra per cosa costa metterlo dove non serve.
    """
    letti_per_riga = {n for n, d in contract.datasets.items() if d.read_by_row}
    assert letti_per_riga == {"DUP_DATASET"}


# ---------------------------------------------------------------------------
# I controlli: settimana e capienze
# ---------------------------------------------------------------------------

from datetime import datetime  # noqa: E402

from fasterreports.core.coherence import BLOCCA, SEGNALA, check_sources  # noqa: E402


def _chiusi(*giorni: str) -> list:
    return [datetime.fromisoformat(g) for g in giorni]


def _finding(rep, frammento):
    return next((f for f in rep.findings if frammento in f.check), None)


def test_settimana_duplicati_allineata_non_dice_niente():
    rep = check_sources(
        dup_closed=_chiusi("2026-08-11T10:00", "2026-08-13T15:00"),
        week_inferred=(2026, 33),
    )
    assert _finding(rep, "duplicati") is None
    assert rep.ok


def test_settimana_duplicati_sfasata_blocca():
    """W32 contro W33: e' il caso che si vedeva nel template.

    Deciso che le due settimane devono coincidere, quindi non SEGNALA: BLOCCA.
    Un report che porta la sezione duplicati di un'altra settimana non ha nessuna
    etichetta che lo dica.
    """
    rep = check_sources(
        dup_closed=_chiusi("2026-08-04T10:00", "2026-08-07T15:00"),
        week_inferred=(2026, 33),
    )
    f = _finding(rep, "duplicati e' di un'altra settimana")
    assert f is not None and f.level == BLOCCA
    assert not rep.ok
    # Entrambi i periodi nel messaggio: senza, non si sa quale file riscaricare.
    testo = f.summary + " ".join(f.details)
    assert "W32" in testo and "W33" in testo
    assert "2026-08-04" in testo


def test_un_caso_a_cavallo_della_mezzanotte_non_sposta_la_settimana():
    """La moda dei giorni, non il min/max.

    Lo stesso motivo per cui `_resolve_week` fa cosi' su `AT_DATASET`: un solo
    caso chiuso appena dentro la settimana precedente non deve far risultare tutto
    l'export della settimana sbagliata — che sarebbe un blocco falso, e i blocchi
    falsi si imparano a ignorare.
    """
    rep = check_sources(
        dup_closed=_chiusi(
            "2026-08-09T23:58",  # domenica: W32
            "2026-08-11T09:00", "2026-08-12T09:00", "2026-08-13T09:00",  # W33
        ),
        week_inferred=(2026, 33),
    )
    assert _finding(rep, "duplicati e' di un'altra settimana") is None


def test_senza_la_settimana_del_report_il_controllo_si_salta():
    rep = check_sources(dup_closed=_chiusi("2026-08-04T10:00"), week_inferred=None)
    f = _finding(rep, "settimana dei duplicati")
    assert f is not None and f.level == SEGNALA
    assert rep.ok


def test_capienza_all_82_percento_segnala():
    """28 agenti su 34: e' la misura vera della W32."""
    rep = check_sources(
        dup_capienze={"agenti con casi duplicati": 34},
        dup_conteggi={"agenti con casi duplicati": 28},
    )
    f = _finding(rep, "posto quasi finito")
    assert f is not None and f.level == SEGNALA
    assert "28 su 34" in f.summary and "82%" in f.summary
    assert rep.ok


def test_capienza_sotto_la_soglia_tace():
    rep = check_sources(
        dup_capienze={"record type": 12},
        dup_conteggi={"record type": 7},
    )
    assert _finding(rep, "posto") is None


def test_capienza_superata_blocca_e_dice_quante_voci_si_perdono():
    rep = check_sources(
        dup_capienze={"agenti con casi duplicati": 34},
        dup_conteggi={"agenti con casi duplicati": 37},
    )
    f = _finding(rep, "posto finito")
    assert f is not None and f.level == BLOCCA
    assert "37" in f.summary and "34" in f.summary
    assert "3 voci in eccesso" in f.hint
    assert not rep.ok


def test_le_capienze_si_misurano_sul_template_vero(contract):
    """I numeri del piano, riletti dal file invece che ricopiati."""
    if not TEMPLATE.is_file():
        pytest.skip("template assente")
    from dataclasses import replace as _replace

    from fasterreports.omni.orchestrate import _dup_capienze
    from fasterreports.omni.settings import load_settings

    s = load_settings(ROOT / "config" / "settings.yml")
    capienze = _dup_capienze(_replace(s, template=TEMPLATE))
    assert capienze["agenti con casi duplicati"] == 34
    assert capienze["record type"] == 12
    assert capienze["Type (case type)"] == 26
    assert capienze["parent con piu' di un duplicato"] == 20
    assert capienze["agenti (aree dei grafici)"] == 45
    assert capienze["Case Origin distinti (menu del filtro)"] == 15


def test_i_conteggi_vengono_dai_dati(contract, tmp_path):
    """Quanti distinti servono, dal blocco appena letto."""
    from fasterreports.omni.orchestrate import _dup_conteggi

    _ds, _src, block = _blocco(contract, _dup_export(tmp_path, righe=3))
    conteggi = _dup_conteggi(contract, {"DUP_DATASET": block})
    assert conteggi["agenti con casi duplicati"] == 3
    assert conteggi["record type"] == 1        # tutte 'Technical'
    assert conteggi["Type (case type)"] == 1   # tutte 'Booking Information'
    # I parent sintetici sono tutti diversi: nessun cluster.
    assert conteggi["parent con piu' di un duplicato"] == 0
