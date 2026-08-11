"""`--week auto`: niente da digitare, niente prompt che aspetta un tasto.

Serve per lanciare il comando senza intervento umano (un'attività pianificata
di Windows), e per farlo bene deve rispondere sempre alla stessa domanda,
qualunque sia il giorno in cui gira: qual è l'ultima settimana lunedì-domenica
già conclusa?
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pytest

import fasterreports.omni.cli as cli_mod
from fasterreports.omni.cli import _settimana_auto, _week_number

ROOT = Path(__file__).resolve().parents[1]


@pytest.mark.parametrize(
    "oggi, attesa",
    [
        (date(2026, 8, 3), 31),   # lunedì -> la settimana appena chiusa (27/7-2/8)
        (date(2026, 8, 9), 31),   # domenica della stessa settimana ISO 32
        (date(2026, 8, 10), 32),  # lunedì successivo -> scatta la 32
        (date(2026, 1, 1), 52),   # a cavallo d'anno: la 2025-W52 finiva il 28/12
    ],
)
def test_settimana_auto_e_sempre_lultima_conclusa(oggi, attesa):
    assert _settimana_auto(oggi) == attesa


def test_settimana_auto_non_dipende_dal_giorno_della_settimana():
    """Lanciato di lunedì o di venerdì deve dare la STESSA settimana, se la
    settimana corrente non è ancora `oggi` a cavallo della successiva."""
    lunedi = _settimana_auto(date(2026, 8, 3))
    venerdi = _settimana_auto(date(2026, 8, 7))
    assert lunedi == venerdi


def test_week_number_non_conosce_auto():
    """`_settimana_auto` risolve 'auto' PRIMA che arrivi qui: `_week_number`
    si occupa solo di un numero (o W-numero) già concreto."""
    with pytest.raises(Exception):
        _week_number("auto")


# ---------------------------------------------------------------------------
# La sostituzione in main(): 'auto' diventa un numero prima di tutto il resto
# ---------------------------------------------------------------------------


def test_main_sostituisce_auto_prima_di_aprire_qualunque_cartella(monkeypatch, tmp_path):
    """Se la sostituzione non avvenisse subito, 'auto' finirebbe nel nome del
    file di preflight e in --week passato a valle — invece deve arrivare gia'
    risolta a un numero concreto."""
    monkeypatch.setattr(cli_mod, "_settimana_auto", lambda: 33)

    rc = cli_mod.main([
        "preflight", "--week", "auto",
        "--input", str(tmp_path / "input"),
        "--output", str(tmp_path / "output"),
    ])

    # Cartella input vuota/inesistente: preflight blocca su tutte le fonti,
    # ma e' un esito atteso — quello che importa e' CHE NUMERO ha usato.
    assert rc == 1
    # Il nome del file e' la prova: e' quello che dimostra che 'auto' e' stato
    # risolto PRIMA di usare la settimana per qualunque altra cosa. Il corpo
    # del rapporto non contiene per forza il numero: con tutte le fonti
    # assenti (come qui) non si arriva nemmeno a leggere AT_DATASET, quindi
    # 'Settimana dai dati' non compare affatto — atteso, non un problema.
    assert (tmp_path / "output" / "preflight_W33.txt").is_file(), (
        "il file va nominato con la settimana risolta, non 'auto'"
    )


def test_main_con_settimana_numerica_non_la_tocca(monkeypatch, tmp_path):
    """Un numero vero non deve passare da `_settimana_auto`."""
    chiamato = []
    monkeypatch.setattr(cli_mod, "_settimana_auto", lambda: chiamato.append(1) or 0)

    cli_mod.main([
        "preflight", "--week", "31",
        "--input", str(tmp_path / "input"),
        "--output", str(tmp_path / "output"),
    ])

    assert not chiamato
    assert (tmp_path / "output" / "preflight_W31.txt").is_file()
