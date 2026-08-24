"""La sezione Only Cases: due fogli che il programma non scrive ma deve sorvegliare.

`OC Eventi` e `Only Cases Dashboard` sono arrivati nel template il 2026-08-21,
disegnati e formattati a mano. La pipeline non ci scrive dentro: si ricalcolano da
`AT_DATASET`. Quello che puo' andare storto, quindi, non e' la scrittura — sono le
due cose che arrivano da sole quando cresce il volume:

- **il tetto delle formule.** I due fogli sono nati leggendo `AT_DATASET` fino a
  riga 60000, dove il resto del template arriva a 130000. Il limite piu' basso
  vince su tutti (`coherence._check_row_limits` prende il vincolante), quindi un
  foglio nuovo puo' dimezzare il margine di TUTTO il report senza che se ne parli.
- **la capienza degli elenchi.** L'elenco cresce da se' (array dinamico), le
  colonne accanto no: la voce in eccesso compare senza nessun numero accanto.

Il foglio si difende gia' da solo, con due celle che dicono «ATTENZIONE: capienza
… esaurita». Questi test verificano che il programma misuri LE STESSE capienze —
se un giorno le formule venissero tirate senza aggiornare la guardia, o viceversa,
i due numeri divergerebbero e si vedrebbe qui.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import replace as _replace
from pathlib import Path

import pytest

from fasterreports.core.capienze import Scaffale, capienze
from fasterreports.core.coherence import BLOCCA, SEGNALA, check_sources
from fasterreports.core.onlycases import (
    FOGLIO_DASHBOARD,
    FOGLIO_HELPER,
    SCAFFALI,
    STATO_DI_DEFAULT,
    conteggi,
    punti_da_misurare,
)

ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "template" / "Omni_Report_TEMPLATE.xlsm"

pytestmark = pytest.mark.skipif(not TEMPLATE.is_file(), reason="template assente")


def _finding(rep, frammento):
    return next((f for f in rep.findings if frammento in f.check), None)


def _settings():
    from fasterreports.omni.settings import load_settings

    return _replace(load_settings(ROOT / "config" / "settings.yml"), template=TEMPLATE)


# ---------------------------------------------------------------------------
# I fogli nel template vero
# ---------------------------------------------------------------------------


def test_i_due_fogli_ci_sono_e_non_hanno_la_settimana_nel_nome():
    """Il nome del foglio e' un contratto col codice, quindi non puo' scadere.

    Nasceva come 'Only Cases W33'. Rinominarlo ogni settimana in Excel funziona —
    Excel aggiorna le formule da se' — ma il controllo delle capienze cerca il
    foglio per nome, e il lunedi' dopo non lo troverebbe piu'. Un controllo che
    sparisce in silenzio e' peggio di un controllo che non c'e': il preflight
    continua a dire OK.
    """
    from fasterreports.core.templatescan import _sheet_targets

    with zipfile.ZipFile(TEMPLATE) as z:
        fogli = set(_sheet_targets(z, TEMPLATE))

    assert FOGLIO_HELPER in fogli
    assert FOGLIO_DASHBOARD in fogli
    con_settimana = sorted(f for f in fogli if re.search(r"\bW\d{1,2}\b", f))
    assert not con_settimana, (
        f"fogli col numero di settimana nel nome: {con_settimana}.\n"
        f"  Vanno rinominati stabili: la settimana si legge dalle celle, che si "
        f"ricalcolano.\n"
        f"  Per 'Only Cases' c'e' lo strumento: "
        f"python tools/patch_template_only_cases.py template/Omni_Report_TEMPLATE.xlsm"
    )


def test_i_fogli_nuovi_non_abbassano_il_tetto_di_AT_DATASET(contract):
    """Un foglio di presentazione non deve stringere il margine di tutto il report.

    E' successo: nati a 60000 quando `AT_DATASET` ha le formule `P`/`Q`
    pre-riempite fino a 130000, i due fogli avevano portato il limite vincolante
    da 130000 a 60000 — cioe' il BLOCCA del preflight si sarebbe spostato a poco
    piu' del doppio del volume attuale, per una sezione secondaria.
    """
    from fasterreports.core.templatescan import binding_limits, scan_row_limits

    ds = contract.dataset("AT_DATASET")
    vinc = binding_limits(scan_row_limits(TEMPLATE, ("AT_DATASET",)))["AT_DATASET"]
    assert vinc.max_row >= ds.max_template_row, (
        f"il limite piu' stretto su AT_DATASET e' {vinc.max_row} "
        f"({vinc.sheet}!{vinc.ref}), sotto le {ds.max_template_row} righe di "
        f"formule del template."
    )


def test_la_capienza_misurata_e_quella_che_dichiara_il_foglio():
    """Il programma e la guardia dentro il foglio devono dire lo stesso numero.

    La dashboard ha due celle che si controllano da sole:

        B8 = COUNT('OC Eventi'!$A$5:$A$3004)     -> 3000 eventi
        B9 = COUNTA($A$12:$A$71)                 -> 60 agenti

    e accanto un `IF(...>=3000, "ATTENZIONE...")`. Sono la dichiarazione di
    capienza scritta da chi ha fatto il foglio. Il preflight misura la stessa cosa
    per un'altra strada — l'ultima riga con una formula — e i due numeri devono
    coincidere: se divergono, o le formule sono state tirate senza aggiornare la
    guardia, o la colonna che misuriamo non e' piu' quella giusta.
    """
    from fasterreports.core.templatescan import _sheet_targets, _unescape

    misurate = capienze(SCAFFALI, _estensioni())
    assert misurate == {"eventi Only Cases": 3000, "agenti Only Cases": 60}

    with zipfile.ZipFile(TEMPLATE) as z:
        raw = z.read(_sheet_targets(z, TEMPLATE)[FOGLIO_DASHBOARD]).decode("utf8", "replace")
    formule = _unescape(raw)

    # `COUNT('OC Eventi'!$A$5:$A$3004)` -> l'elenco eventi arriva a 3004.
    m = re.search(r"COUNT\('OC Eventi'!\$A\$(\d+):\$A\$(\d+)\)", formule)
    assert m, "la dashboard non dichiara piu' la capienza degli eventi in B8"
    assert int(m.group(2)) - int(m.group(1)) + 1 == misurate["eventi Only Cases"]

    # `COUNTA($A$12:$A$71)` -> l'elenco agenti arriva a 71.
    m = re.search(r"COUNTA\(\$A\$(\d+):\$A\$(\d+)\)", formule)
    assert m, "la dashboard non dichiara piu' la capienza degli agenti in B9"
    assert int(m.group(2)) - int(m.group(1)) + 1 == misurate["agenti Only Cases"]


def test_la_colonna_misurata_non_prosegue_oltre_l_elenco():
    """Perche' si misura `G` e non `E`, che sarebbe la scelta naturale.

    Sotto l'elenco degli agenti (`A12:J71`) la dashboard ha un SECONDO blocco: il
    dettaglio giornaliero del singolo agente, righe 79-94. `B`, `C`, `D`, `E`, `F`
    hanno formule anche li', quindi misurarle direbbe 83 posti invece di 60 — e il
    controllo tacerebbe fino a ventitre' agenti oltre il vero, che e' precisamente
    il caso per cui esiste.
    """
    est = _estensioni_di(FOGLIO_DASHBOARD, "BCDEFGHIJ", prima_riga=12)
    assert est["E"] > 72, "il secondo blocco non c'e' piu': rivedere lo scaffale"
    for col in "GHIJ":
        assert est[col] == 72, f"{col} non si ferma piu' alla riga del totale"


def test_lo_stato_si_legge_dal_foglio_non_dal_codice():
    """`Only Cases Dashboard`!B5 dice quale stato agente e' "only cases".

    Sta li' e non in Python perche' e' un parametro del foglio: chi lo cambia in
    Excel deve ottenere che il conteggio del preflight cambi con lui.
    """
    from fasterreports.omni.orchestrate import _oc_stato

    assert _oc_stato(_settings()) == STATO_DI_DEFAULT == "Available Cases"


def test_le_capienze_si_misurano_sul_template_vero():
    from fasterreports.omni.orchestrate import _oc_capienze

    assert _oc_capienze(_settings()) == {
        "eventi Only Cases": 3000,
        "agenti Only Cases": 60,
    }


# ---------------------------------------------------------------------------
# I conteggi: quanto chiede la settimana
# ---------------------------------------------------------------------------

# (Agent Email, Agent State) nella posizione che hanno nel blocco AT_DATASET:
# il blocco parte da B, quindi Email e' 0 e State e' 4.
_OFFSET = {"Agent Email": 0, "Agent State": 4}


def _riga(email: str, stato: str) -> list:
    return [email, "", "", "", stato, "", "", "", "", "", "", "", ""]


def test_si_contano_solo_le_righe_nello_stato_giusto():
    """`FILTER` nel foglio guarda uno stato solo, e il conteggio deve fare uguale.

    Contare tutte le righe di `AT_DATASET` direbbe 26 541 dove il foglio ne mostra
    1 386: il controllo bloccherebbe ogni settimana per un problema che non
    esiste, e chi lo vede imparerebbe a ignorarlo.
    """
    righe = [
        _riga("a@x.it", "Available Cases"),
        _riga("a@x.it", "Available Cases"),
        _riga("b@x.it", "Available Cases"),
        _riga("c@x.it", "Break"),
        _riga("d@x.it", "Offline"),
    ]
    assert conteggi(righe, _OFFSET, "Available Cases") == {
        "eventi Only Cases": 3,
        "agenti Only Cases": 2,
    }


def test_uno_stato_che_non_c_e_da_zero_e_non_esplode():
    righe = [_riga("a@x.it", "Break")]
    assert conteggi(righe, _OFFSET, "Available Cases") == {
        "eventi Only Cases": 0,
        "agenti Only Cases": 0,
    }


def test_senza_la_colonna_dello_stato_non_si_conta_niente():
    """Meglio nessun numero che un numero costruito su una colonna assente."""
    assert conteggi([_riga("a@x.it", "Available Cases")], {"Agent Email": 0}, "x") == {}


def test_i_conteggi_vengono_dal_blocco(contract):
    """Il giro completo: dal blocco AT_DATASET alle etichette degli scaffali."""
    from fasterreports.core.transform import Block
    from fasterreports.omni.orchestrate import _oc_conteggi

    block = Block(
        dataset="AT_DATASET",
        rows=[_riga("a@x.it", "Available Cases"), _riga("b@x.it", "Break")],
        start_col="B",
        end_col="N",
        stats={},
    )
    assert _oc_conteggi(contract, _settings(), {"AT_DATASET": block}) == {
        "eventi Only Cases": 1,
        "agenti Only Cases": 1,
    }


# ---------------------------------------------------------------------------
# Il controllo, con i numeri veri
# ---------------------------------------------------------------------------


def test_capienza_sotto_la_soglia_tace():
    """1386 eventi su 3000 e 33 agenti su 60: la misura vera del W30."""
    rep = check_sources(
        oc_capienze={"eventi Only Cases": 3000, "agenti Only Cases": 60},
        oc_conteggi={"eventi Only Cases": 1386, "agenti Only Cases": 33},
    )
    assert _finding(rep, "posto") is None
    assert rep.ok


def test_capienza_vicina_segnala_prima_che_morda():
    rep = check_sources(
        oc_capienze={"agenti Only Cases": 60},
        oc_conteggi={"agenti Only Cases": 51},
    )
    f = _finding(rep, "posto quasi finito")
    assert f is not None and f.level == SEGNALA
    assert "51 su 60" in f.summary
    assert rep.ok


def test_capienza_superata_blocca_e_indica_dove_tirare_le_formule():
    rep = check_sources(
        oc_capienze={"eventi Only Cases": 3000},
        oc_conteggi={"eventi Only Cases": 3200},
    )
    f = _finding(rep, "posto finito")
    assert f is not None and f.level == BLOCCA
    assert "200 voci in eccesso" in f.hint
    assert "core/onlycases.py" in f.hint  # il rimedio giusto, non quello dei duplicati
    assert not rep.ok


def test_le_due_sezioni_non_si_confondono():
    """Duplicate Cases e Only Cases usano lo stesso controllo, con rimedi diversi."""
    rep = check_sources(
        dup_capienze={"agenti con casi duplicati": 34},
        dup_conteggi={"agenti con casi duplicati": 37},
        oc_capienze={"eventi Only Cases": 3000},
        oc_conteggi={"eventi Only Cases": 3200},
    )
    trovati = [f for f in rep.findings if "posto finito" in f.check]
    assert len(trovati) == 2
    rimedi = {f.check: f.hint for f in trovati}
    assert "piano-duplicates.md" in rimedi["posto finito per agenti con casi duplicati"]
    assert "onlycases.py" in rimedi["posto finito per eventi Only Cases"]


def test_uno_scaffale_con_la_coda_scarta_la_riga_del_totale():
    """La proprieta' della dataclass, senza passare dal template."""
    con_totale = Scaffale("x", "F", "G", 12, campo=None, coda=1)
    senza = Scaffale("x", "F", "G", 12, campo=None)
    assert con_totale.capienza(72) == 60
    assert senza.capienza(72) == 61


# ---------------------------------------------------------------------------


def _estensioni():
    from fasterreports.core.templatescan import scan_formula_extent

    return scan_formula_extent(TEMPLATE, punti_da_misurare())


def _estensioni_di(foglio: str, colonne: str, prima_riga: int) -> dict[str, int]:
    from fasterreports.core.templatescan import scan_formula_extent

    est = scan_formula_extent(
        TEMPLATE, tuple((foglio, c, prima_riga) for c in colonne)
    )
    return {c: est[(foglio, c)] for c in colonne if (foglio, c) in est}
