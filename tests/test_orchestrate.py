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
    quali delle sei mancavano.
    """
    report, blocks = run_preflight(contratto, cfg, week_number=31)
    assert len(report.datasets) == len(contratto.datasets)
    assert not report.ok
    assert blocks == {}
    for d in report.datasets:
        assert d.error, d.name
        assert "nessun file corrispondente" in d.source


def test_una_fonte_su_sei_non_impedisce_di_leggerla(contratto, cfg):
    scrivi_at_csv(cfg, [riga_at("2026-07-28 12:00:00")])
    report, blocks = run_preflight(contratto, cfg, week_number=31)
    assert _ds(report, "AT_DATASET").ok
    assert "AT_DATASET" in blocks
    # le altre cinque restano bloccate, ognuna col suo perche'
    assert sum(1 for d in report.datasets if d.error) == len(contratto.datasets) - 1


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


def test_due_settimane_in_cartella_bloccano_il_dataset(contratto, cfg):
    scrivi_at_csv(cfg, [riga_at("2026-07-28 12:00:00")], nome="AT DATASET W30.csv")
    scrivi_at_csv(cfg, [riga_at("2026-07-28 12:00:00")], nome="AT DATASET W31.csv")
    report, blocks = run_preflight(contratto, cfg, week_number=31)
    d = _ds(report, "AT_DATASET")
    assert d.error and "W30" in d.error and "W31" in d.error
    assert "AT_DATASET" not in blocks
