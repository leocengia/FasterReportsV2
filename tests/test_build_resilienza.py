"""Il build resiste a un file bloccato, a un errore COM transitorio, e non
lascia mai un file a meta' scrittura con il nome del file buono.

Tre guasti diversi con la stessa causa di fondo: scrivere in un workbook Excel
non e' un'operazione atomica vista da fuori, a meno che qualcosa non la renda
tale apposta.
"""

from __future__ import annotations

from dataclasses import replace
from pathlib import Path

import pytest

from fasterreports.core.contract import load_contract
from fasterreports.core.errors import PipelineError
from fasterreports.omni.orchestrate import (
    _bloccato_da_excel,
    _con_retry,
    _e_transitorio,
    _scrivi_con_copia_atomica,
    build,
)
from fasterreports.omni.settings import load_settings

ROOT = Path(__file__).resolve().parents[1]

AT_HEADERS = [
    "Agent Email", "Agent State", "Number of Active Contacts",
    "Start Time", "Total Time in seconds", "Productive Aux Flag (Yes / No)",
]


@pytest.fixture(scope="module")
def contratto():
    return load_contract(ROOT / "config" / "columns.yml")


@pytest.fixture
def cfg(tmp_path):
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


# ---------------------------------------------------------------------------
# Il file di lock di Excel
# ---------------------------------------------------------------------------


def test_bloccato_da_excel_rileva_il_lock(tmp_path):
    p = tmp_path / "Omni_Report_W31.xlsm"
    assert not _bloccato_da_excel(p)
    (tmp_path / "~$Omni_Report_W31.xlsm").touch()
    assert _bloccato_da_excel(p)


def test_build_si_ferma_se_il_file_di_output_e_aperto(contratto, cfg):
    """Il lock di Excel blocca PRIMA di qualunque altra cosa — prima ancora
    del controllo su --only, prima di toccare il template. Sovrascrivere un
    file che un collega ha ancora aperto produce un salvataggio a meta' o un
    errore COM poco chiaro, non un messaggio leggibile."""
    scrivi_at_csv(cfg, [riga_at("2026-07-28 12:00:00")])
    cfg.output_dir.mkdir(parents=True, exist_ok=True)
    (cfg.output_dir / "~$Omni_Report_W31.xlsm").touch()

    with pytest.raises(PipelineError) as e:
        build(contratto, cfg, "31", week_number=31, only={"AT_DATASET"})
    msg = str(e.value)
    assert "aperto in Excel" in msg
    assert "~$Omni_Report_W31.xlsm" in msg


def test_build_senza_lock_procede_oltre(contratto, cfg):
    """Senza il file di lock l'errore e' un altro (manca il workbook per il
    giro parziale): prova che il controllo del lock non scatta a vuoto."""
    scrivi_at_csv(cfg, [riga_at("2026-07-28 12:00:00")])
    with pytest.raises(PipelineError) as e:
        build(contratto, cfg, "31", week_number=31, only={"AT_DATASET"})
    assert "aperto in Excel" not in str(e.value)
    assert "giro completo" in str(e.value)


# ---------------------------------------------------------------------------
# Retry sugli errori COM transitori
# ---------------------------------------------------------------------------


def test_e_transitorio_riconosce_i_marcatori_noti():
    assert _e_transitorio(Exception("RPC_E_CALL_REJECTED: pupupu"))
    assert _e_transitorio(Exception("Call was rejected by callee."))
    assert not _e_transitorio(Exception("Foglio 'X' assente."))


def test_con_retry_ritenta_solo_gli_errori_transitori():
    tentativi = []

    def fallisce_due_volte_poi_ok():
        tentativi.append(1)
        if len(tentativi) < 3:
            raise Exception("Call was rejected by callee.")
        return "fatto"

    assert _con_retry(fallisce_due_volte_poi_ok, attesa_iniziale=0.001) == "fatto"
    assert len(tentativi) == 3


def test_con_retry_non_ritenta_un_errore_non_transitorio():
    """Un errore che non e' rumore COM (un file mancante, un foglio che non
    c'e') deve fallire SUBITO: ritentarlo perderebbe solo tempo prima di
    fallire comunque."""
    tentativi = []

    def fallisce_per_un_altro_motivo():
        tentativi.append(1)
        raise ValueError("qualcos'altro, non e' COM")

    with pytest.raises(ValueError):
        _con_retry(fallisce_per_un_altro_motivo, attesa_iniziale=0.001)
    assert len(tentativi) == 1


def test_con_retry_alza_bandiera_bianca_dopo_i_tentativi():
    tentativi = []

    def fallisce_sempre_per_un_motivo_transitorio():
        tentativi.append(1)
        raise Exception("Server execution failed")

    with pytest.raises(Exception):
        _con_retry(
            fallisce_sempre_per_un_motivo_transitorio, tentativi=3, attesa_iniziale=0.001
        )
    assert len(tentativi) == 3


# ---------------------------------------------------------------------------
# Scrittura atomica: il file buono non esiste mai a meta'
#
# Verificabile per intero SENZA Excel: `_scrivi_con_copia_atomica` non sa
# niente di xlwings, prende solo un `sorgente`, un `out_path`, e una funzione
# `scrivi` da chiamare sulla copia temporanea.
# ---------------------------------------------------------------------------


def test_promuove_al_nome_buono_solo_al_successo(tmp_path):
    sorgente = tmp_path / "modello.txt"
    sorgente.write_text("originale", encoding="utf-8")
    out = tmp_path / "risultato.txt"

    _scrivi_con_copia_atomica(
        sorgente, out, lambda temp: temp.write_text("scritto per davvero", encoding="utf-8")
    )

    assert out.read_text(encoding="utf-8") == "scritto per davvero"
    assert not (tmp_path / "risultato.building.txt").exists()


def test_non_tocca_il_file_buono_se_il_lavoro_fallisce_a_meta(tmp_path):
    out = tmp_path / "risultato.txt"
    out.write_text("versione precedente, buona", encoding="utf-8")
    sorgente = tmp_path / "modello.txt"
    sorgente.write_text("nuovo modello", encoding="utf-8")

    def scrittura_a_meta_poi_fallisce(temp: Path) -> None:
        temp.write_text("scrittura a meta'...", encoding="utf-8")
        raise RuntimeError("la macro e' fallita a meta' strada")

    with pytest.raises(RuntimeError):
        _scrivi_con_copia_atomica(sorgente, out, scrittura_a_meta_poi_fallisce)

    # Il file con il nome buono e' ESATTAMENTE quello di prima: nessuno
    # potrebbe scambiarlo per il risultato, sbagliato, del giro fallito.
    assert out.read_text(encoding="utf-8") == "versione precedente, buona"
    assert not (tmp_path / "risultato.building.txt").exists()


def test_su_file_nuovo_un_fallimento_non_lascia_nessun_residuo(tmp_path):
    """Caso 'giro completo': out_path non esiste ancora. Se il lavoro
    fallisce, non deve apparire ne' con il nome buono ne' con quello
    temporaneo — altrimenti resterebbe un file fantasma nella cartella."""
    out = tmp_path / "risultato.txt"
    sorgente = tmp_path / "modello.txt"
    sorgente.write_text("modello", encoding="utf-8")

    def fallisce(temp: Path) -> None:
        raise RuntimeError("no")

    with pytest.raises(RuntimeError):
        _scrivi_con_copia_atomica(sorgente, out, fallisce)

    assert not out.exists()
    assert not (tmp_path / "risultato.building.txt").exists()
