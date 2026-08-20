"""L'orchestratore: dalle sorgenti in cartella al rapporto di preflight.

Sono le proprieta' che i test sui singoli pezzi non coprono, perche' nascono dal
modo in cui i pezzi sono cuciti insieme. Le due qui sotto sono guasti veri,
trovati facendo girare la pipeline sugli export reali della W31:

- il rapporto moriva sulla prima fonte assente, invece di elencarle tutte;
- il confronto fra la settimana digitata e quella dei dati girava solo se le
  fonti WFM c'erano, cosi' `--week 30` su un export della W31 passava.
"""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path

import pytest

from fasterreports.core.contract import load_contract
from fasterreports.omni.orchestrate import run_preflight
from fasterreports.omni.settings import load_settings

ROOT = Path(__file__).resolve().parents[1]

AT_HEADERS = [
    "Agent Email",
    "Agent State",
    "Number of Active Contacts",
    "Start Time",
    "Total Time in seconds",
    "Productive Aux Flag (Yes / No)",
]


@pytest.fixture(scope="module")
def contratto():
    return load_contract(ROOT / "config" / "columns.yml")


@pytest.fixture
def cfg(tmp_path):
    """Settings reali, con input/ e template dirottati sul tmp.

    Il template inesistente porta l'offset fuso a 0: cosi' le date di prova
    valgono come sono scritte, senza il +9 di Seattle->Milano.
    """
    s = load_settings(ROOT / "config" / "settings.yml", root=ROOT)
    (tmp_path / "input").mkdir()
    return replace(
        s,
        input_dir=tmp_path / "input",
        output_dir=tmp_path / "output",
        template=tmp_path / "template_assente.xlsm",
    )


def scrivi_at_csv(cfg, righe, *, nome="AT DATASET W31.csv") -> Path:
    p = cfg.input_dir / nome
    corpo = [",".join(AT_HEADERS)]
    corpo += [",".join(str(c) for c in r) for r in righe]
    p.write_text("\n".join(corpo) + "\n", encoding="utf-8")
    return p


def riga_at(quando: str, email="mario.rossi@example.com"):
    return [email, "Available", 1, quando, 300, "Yes"]


def _ds(report, nome):
    return next(d for d in report.datasets if d.name == nome)


# ---------------------------------------------------------------------------
# Il rapporto sopravvive alle fonti che mancano
# ---------------------------------------------------------------------------


def test_fonti_assenti_elencate_tutte(contratto, cfg):
    """Sei righe, non un'eccezione alla prima.

    Prima il ramo che scriveva l'errore chiamava di nuovo `input_path`, che
    solleva proprio perche' il file non c'e': il preflight moriva mentre
    descriveva il guasto, e chi lanciava vedeva una fonte sola invece di sapere
    quali mancavano.

    Le fonti `optional` sono l'eccezione dichiarata (`PSAT_DATASET`,
    `DUP_DATASET`): la loro riga e' SALTATA (non BLOCCATA) e il resto del report
    procede lo stesso attorno a loro — vedi i test dedicati piu' sotto per il
    perche'. Per ognuna resta un blocco VUOTO, che e' cio' che fa PULIRE il
    foglio: un residuo della settimana prima sarebbe un report sbagliato, non uno
    che manca.
    """
    report, blocks = run_preflight(contratto, cfg, week_number=31)
    assert len(report.datasets) == len(contratto.datasets)
    assert not report.ok  # le fonti obbligatorie mancano comunque
    opzionali = {n for n, d in contratto.datasets.items() if d.optional}
    assert set(blocks) == opzionali, "solo i blocchi vuoti delle opzionali"
    for d in report.datasets:
        if d.name in opzionali:
            assert not d.error and d.skipped_reason
        else:
            assert d.error, d.name
        assert "nessun file corrispondente" in d.source


def test_una_fonte_sola_non_impedisce_di_leggerla(contratto, cfg):
    scrivi_at_csv(cfg, [riga_at("2026-07-28 12:00:00")])
    report, blocks = run_preflight(contratto, cfg, week_number=31)
    assert _ds(report, "AT_DATASET").ok
    assert "AT_DATASET" in blocks
    # Le fonti OPZIONALI assenti sono saltate, non bloccate. Restano bloccate le
    # obbligatorie che mancano — cioe' tutte tranne AT_DATASET, che c'e'.
    # Il numero viene dal contratto: scriverlo a mano vorrebbe dire aggiornarlo a
    # ogni dataset nuovo, e nel frattempo il test misurerebbe la cosa sbagliata.
    obbligatorie_mancanti = sum(
        1 for d in contratto.datasets.values() if not d.optional and d.name != "AT_DATASET"
    )
    assert sum(1 for d in report.datasets if d.error) == obbligatorie_mancanti
    for nome, ds in contratto.datasets.items():
        if ds.optional:
            assert _ds(report, nome).skipped_reason, f"{nome} doveva essere saltata"
            assert not _ds(report, nome).error


# ---------------------------------------------------------------------------
# La fonte opzionale: assente si segnala, sbagliata blocca comunque
# ---------------------------------------------------------------------------


def test_fonte_opzionale_assente_segnala_e_non_blocca_il_resto(contratto, cfg):
    """Il caso vero della W32: l'export dei sondaggi non esisteva affatto.

    Nessuna regola di malpractice legge PSAT_DATASET (nessun `consumers` inizia
    per `VBA:`): bloccare l'intero report per una fonte che il motore non usa
    sarebbe fermare le regole su AT/SF/Turni per un problema che non le
    riguarda.
    """
    scrivi_at_csv(cfg, [riga_at("2026-07-28 12:00:00")])
    report, blocks = run_preflight(contratto, cfg, week_number=31, only={"AT_DATASET", "PSAT_DATASET"})
    psat = _ds(report, "PSAT_DATASET")
    assert psat.ok  # skipped_reason non e' un errore
    assert psat.skipped_reason
    assert "SALTATO" in report.render()
    assert "BLOCCATO" not in report.render().split("PSAT_DATASET")[1].split("###")[0]


def test_fonte_opzionale_assente_produce_un_blocco_vuoto_non_lassenza(contratto, cfg):
    """Il blocco vuoto e' cio' che fa PULIRE il foglio invece di lasciarlo con
    l'ultima settimana scritta: misurato, il template porta residui (130 righe
    in PSAT_DATASET) proprio perche' non e' mai stato svuotato apposta."""
    scrivi_at_csv(cfg, [riga_at("2026-07-28 12:00:00")])
    _, blocks = run_preflight(contratto, cfg, week_number=31, only={"AT_DATASET", "PSAT_DATASET"})
    assert "PSAT_DATASET" in blocks
    vuoto = blocks["PSAT_DATASET"]
    assert vuoto.n_rows == 0
    assert vuoto.dataset == "PSAT_DATASET"


def test_fonte_opzionale_presente_funziona_come_sempre(contratto, cfg):
    """optional=true non deve cambiare nulla quando il file C'E': deve solo
    smettere di bloccare quando non c'e'."""
    import csv

    scrivi_at_csv(cfg, [riga_at("2026-07-28 12:00:00")])
    ds = contratto.dataset("PSAT_DATASET")
    headers = [f.canonical for f in ds.input_fields]
    p = cfg.input_dir / "PSAT DATASET W31.csv"
    with p.open("w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(headers)
        w.writerow(["Mario Rossi", "2026-07-28", "12345", "9", "ottimo"])

    report, blocks = run_preflight(contratto, cfg, week_number=31, only={"AT_DATASET", "PSAT_DATASET"})
    psat = _ds(report, "PSAT_DATASET")
    assert psat.ok and not psat.skipped_reason
    assert blocks["PSAT_DATASET"].n_rows == 1


def test_fonte_opzionale_con_colonne_sbagliate_blocca_comunque(contratto, cfg):
    """Un file che C'E' ma ha le colonne sbagliate non e' "assente questa
    settimana": e' un problema nei dati, e deve bloccare come per qualunque
    altra fonte — optional copre solo il caso "il file non c'e'"."""
    scrivi_at_csv(cfg, [riga_at("2026-07-28 12:00:00")])
    p = cfg.input_dir / "PSAT DATASET W31.csv"
    p.write_text("colonna_a_caso,altra\nx,y\n", encoding="utf-8")

    report, blocks = run_preflight(contratto, cfg, week_number=31, only={"AT_DATASET", "PSAT_DATASET"})
    psat = _ds(report, "PSAT_DATASET")
    assert psat.error and not psat.skipped_reason
    assert "PSAT_DATASET" not in blocks


def test_optional_su_dataset_letto_dal_vba_non_si_carica():
    """Difesa nel caricamento del contratto: un dataset che il VBA consuma non
    puo' essere dichiarato opzionale, altrimenti la sua assenza produrrebbe
    numeri sbagliati in silenzio invece che assenti."""
    from fasterreports.core.errors import ContractError
    from fasterreports.core.contract import parse_contract

    raw = {
        "datasets": {
            "X": {
                "sheet": "X",
                "header_row": 1,
                "data_start_col": "A",
                "optional": True,
                "fields": [{
                    "canonical": "Nome",
                    "target_col": "A",
                    "role": "input",
                    "consumers": ["VBA:AddXRules"],
                }],
            }
        }
    }
    with pytest.raises(ContractError, match="VBA:AddXRules"):
        parse_contract(raw)


# ---------------------------------------------------------------------------
# La settimana: dai dati, non dal numero digitato
# ---------------------------------------------------------------------------


def test_settimana_ricavata_dai_dati(contratto, cfg):
    scrivi_at_csv(
        cfg,
        [riga_at("2026-07-27 09:00:00"), riga_at("2026-08-02 18:00:00")],
    )
    report, _ = run_preflight(contratto, cfg, week_number=31)
    assert report.week == (2026, 31)
    from datetime import date

    assert report.week_bounds == (date(2026, 7, 27), date(2026, 8, 2))
    # E finisce nel testo del rapporto: e' il numero su cui si ritagliano turni
    # e slot, chi legge deve poterlo confrontare con quello che si aspettava.
    reso = report.render()
    assert "W31 2026" in reso
    assert "27/07/2026" in reso and "02/08/2026" in reso


def test_settimana_dichiarata_diversa_blocca_anche_senza_le_fonti_wfm(contratto, cfg):
    """Il guasto: `--week 30` su un export della W31 rispondeva OK.

    Il controllo viveva nella sezione coerenza, che girava solo con roster e back
    office in cartella. Senza quelli sarebbe uscito un `Omni_Report_W30.xlsm`
    pieno di dati della 31 — un report con l'etichetta sbagliata, che e' peggio
    di un report che manca perche' viene archiviato.
    """
    scrivi_at_csv(cfg, [riga_at("2026-07-28 12:00:00")])
    report, _ = run_preflight(contratto, cfg, week_number=30)

    assert _ds(report, "AT_DATASET").ok        # i dati sono a posto
    assert report.coherence is not None        # ma la coerenza gira comunque
    assert not report.coherence.ok             # e blocca
    assert not report.ok
    testo = report.render()
    assert "--week 30" in testo and "31" in testo


def test_settimana_dichiarata_coerente_non_blocca(contratto, cfg):
    scrivi_at_csv(cfg, [riga_at("2026-07-28 12:00:00")])
    report, _ = run_preflight(contratto, cfg, week_number=31)
    motivi = [f.check for f in report.coherence.findings if f.level == "BLOCCA"]
    assert "settimana dichiarata" not in motivi


def test_senza_week_nessun_confronto_ma_la_settimana_si_sa(contratto, cfg):
    """`--week` e' un riscontro, non la fonte della verita': senza, si procede."""
    scrivi_at_csv(cfg, [riga_at("2026-07-28 12:00:00")])
    report, _ = run_preflight(contratto, cfg, week_number=None)
    assert report.week == (2026, 31)
    motivi = [f.check for f in report.coherence.findings if f.level == "BLOCCA"]
    assert "settimana dichiarata" not in motivi


# ---------------------------------------------------------------------------
# Formati misti: e' come arrivano gli export veri
# ---------------------------------------------------------------------------


def test_lo_stesso_dataset_come_csv_o_come_xlsx(contratto, cfg, tmp_path):
    """SF e PSAT arrivano in .csv, AT e ATwi in .xlsx. Il formato lo decide
    l'estensione: a valle il blocco deve essere identico."""
    from xlsxbuild import make_xlsx
    from fasterreports.core.xlsxsource import index_to_col

    righe = [riga_at("2026-07-28 12:00:00"), riga_at("2026-07-29 08:30:00")]

    scrivi_at_csv(cfg, righe)
    _, da_csv = run_preflight(contratto, cfg, week_number=31, only={"AT_DATASET"})
    (cfg.input_dir / "AT DATASET W31.csv").unlink()

    grid: dict[str, object] = {
        f"{index_to_col(i)}1": h for i, h in enumerate(AT_HEADERS, start=1)
    }
    for r, valori in enumerate(righe, start=2):
        for i, v in enumerate(valori, start=1):
            grid[f"{index_to_col(i)}{r}"] = v
    make_xlsx(cfg.input_dir / "AT DATASET W31.xlsx", "Sheet1", grid)
    _, da_xlsx = run_preflight(contratto, cfg, week_number=31, only={"AT_DATASET"})

    assert da_csv["AT_DATASET"].rows == da_xlsx["AT_DATASET"].rows


# ---------------------------------------------------------------------------
# --only: ricarica un sottoinsieme, non ricomincia da capo
# ---------------------------------------------------------------------------


def test_build_parziale_senza_workbook_esistente_si_rifiuta(contratto, cfg):
    """Il guasto: `--only` ripartiva dal template.

    I fogli non ricaricati sarebbero rimasti pieni dei dati della settimana del
    template — un report mezzo W31 e mezzo W30, senza che nulla lo dicesse. E il
    ciclo di scrittura cercava i blocchi di tutti i dataset, quindi
    andava in KeyError proprio nel caso per cui `--only` esiste.
    """
    from fasterreports.core.errors import PipelineError
    from fasterreports.omni.orchestrate import build

    scrivi_at_csv(cfg, [riga_at("2026-07-28 12:00:00")])
    with pytest.raises(PipelineError) as e:
        build(contratto, cfg, "31", week_number=31, only={"AT_DATASET"})
    msg = str(e.value)
    assert "--only" in msg
    assert "giro completo" in msg              # cosa fare
    assert "settimana del template" in msg     # perche'


def test_build_completo_non_pretende_il_workbook(contratto, cfg):
    """Il giro completo riparte dal template: e' l'altro ramo, e non deve
    inciampare nel controllo appena aggiunto."""
    from fasterreports.omni.orchestrate import build

    scrivi_at_csv(cfg, [riga_at("2026-07-28 12:00:00")])
    # Le altre cinque fonti mancano, quindi si fermera' al preflight — che e'
    # un esito, non un'eccezione: il rapporto e' il prodotto.
    res = build(contratto, cfg, "31", week_number=31)
    assert res.workbook is None
    assert res.preflight.is_file()
    assert "--only" not in res.report.render()


def test_due_settimane_in_cartella_bloccano_il_dataset(contratto, cfg):
    scrivi_at_csv(cfg, [riga_at("2026-07-28 12:00:00")], nome="AT DATASET W30.csv")
    scrivi_at_csv(cfg, [riga_at("2026-07-28 12:00:00")], nome="AT DATASET W31.csv")
    report, blocks = run_preflight(contratto, cfg, week_number=31)
    d = _ds(report, "AT_DATASET")
    assert d.error and "W30" in d.error and "W31" in d.error
    assert "AT_DATASET" not in blocks
