"""I controlli che accompagnano lo storico AHT.

Tre guasti, tutti silenziosi se nessuno li cerca:
  - una colonna data letta al contrario (la settimana finisce altrove);
  - l'export SF di una settimana diversa dal resto del report (un file con
    dentro due settimane e nessuna etichetta a dirlo);
  - un case type nei dati che il template non elenca (numeri veri che non
    compaiono in nessun foglio).
"""

from __future__ import annotations

from datetime import datetime

from fasterreports.core.coherence import BLOCCA, SEGNALA, check_sources
from fasterreports.core.transform import ColumnStats

LUNEDI_W33 = datetime(2026, 8, 10)


def trova(rep, pezzo: str):
    return [f for f in rep.findings if pezzo in f.check]


# ---------------------------------------------------------------------------
# Date ambigue non dichiarate
# ---------------------------------------------------------------------------


def test_segnala_una_colonna_data_ambigua():
    st = ColumnStats(canonical="Date Viewpoint", target_col="A")
    st.ambigue = 2833
    st.esempi_ambigui = ["8/10/2026 12:00:00 AM"]
    rep = check_sources(column_stats={"SF_DATABASE": {"Date Viewpoint": st}})

    (f,) = trova(rep, "date ambigue")
    assert f.level == SEGNALA
    assert "2833" in f.summary
    assert "date_format" in f.hint
    # SEGNALA, non BLOCCA: la lettura potrebbe anche essere quella giusta. Quello
    # che non va e' non saperlo.
    assert rep.ok


def test_non_segnala_niente_se_nessuna_data_e_ambigua():
    st = ColumnStats(canonical="Start Time", target_col="I")
    rep = check_sources(column_stats={"AT_DATASET": {"Start Time": st}})
    assert trova(rep, "date ambigue") == []


# ---------------------------------------------------------------------------
# Date Viewpoint: un lunedi', e quello giusto
# ---------------------------------------------------------------------------


def test_date_viewpoint_regolare_non_dice_niente():
    rep = check_sources(date_viewpoint=[LUNEDI_W33] * 100, week_inferred=(2026, 33))
    assert trova(rep, "Date Viewpoint") == []
    assert trova(rep, "altra settimana") == []


def test_blocca_se_date_viewpoint_non_e_un_lunedi():
    """E' il modo in cui si intercetta un cambio di formato dell'export: 10/08
    letto come 8 ottobre cade di giovedi'."""
    rep = check_sources(date_viewpoint=[datetime(2026, 10, 8)] * 100)

    (f,) = trova(rep, "non cade di lunedi")
    assert f.level == BLOCCA
    assert "giovedi" in f.summary
    assert "date_format" in f.hint
    assert not rep.ok


def test_blocca_se_lexport_sf_e_di_unaltra_settimana():
    """Un report con dentro due settimane diverse e nessuna etichetta che lo dica
    e' esattamente il file che viene archiviato e poi riletto come buono."""
    rep = check_sources(date_viewpoint=[LUNEDI_W33] * 50, week_inferred=(2026, 31))

    (f,) = trova(rep, "altra settimana")
    assert f.level == BLOCCA
    assert "W33" in f.summary and "W31" in f.summary
    assert not rep.ok


def test_segnala_se_date_viewpoint_ha_piu_di_un_giorno():
    rep = check_sources(
        date_viewpoint=[LUNEDI_W33] * 50 + [datetime(2026, 8, 17)] * 20,
        week_inferred=(2026, 33),
    )
    (f,) = trova(rep, "non e' un solo giorno")
    assert f.level == SEGNALA
    assert "2" in f.summary


def test_il_giorno_prevalente_decide():
    """Poche righe sporche non devono ribaltare il controllo: vince la moda, la
    stessa scelta che `infer_iso_week` fa per AT_DATASET."""
    rep = check_sources(
        date_viewpoint=[LUNEDI_W33] * 2000 + [datetime(2026, 8, 17)],
        week_inferred=(2026, 33),
    )
    assert trova(rep, "altra settimana") == []
    assert trova(rep, "non cade di lunedi") == []


def test_senza_date_viewpoint_il_controllo_si_salta():
    rep = check_sources(date_viewpoint=None, week_inferred=(2026, 33))
    assert trova(rep, "Date Viewpoint") == []


def test_senza_settimana_inferita_controlla_solo_il_lunedi():
    """Con `--only SF_DATABASE` non si legge AT_DATASET, quindi la settimana di
    confronto non c'e'. Il controllo sul lunedi' si puo' fare comunque."""
    rep = check_sources(date_viewpoint=[LUNEDI_W33] * 10, week_inferred=None)
    assert rep.ok
    assert trova(rep, "altra settimana") == []


# ---------------------------------------------------------------------------
# Case type nuovi
# ---------------------------------------------------------------------------


def test_segnala_i_case_type_nuovi_elencandoli():
    rep = check_sources(casetype_nuovi=[("Phone", "Call Assignment"), ("Non-live", "Collections")])

    (f,) = trova(rep, "case type nuovi")
    assert f.level == SEGNALA
    assert "2 combinazioni" in f.summary
    assert any("Phone | Call Assignment" in d for d in f.details)
    assert "CaseType Deepdive" in f.hint
    # Non blocca: i numeri finiscono comunque nell'helper, non si perde nulla.
    assert rep.ok


def test_distingue_i_nuovi_che_entrano_nel_trend_da_quelli_esclusi():
    """La conseguenza che conta e' sul trend: un case type nuovo NON escluso
    compare da questa settimana nelle heat map, dove prima non c'era, e il
    grafico cambia forma senza che nulla lo dica."""
    rep = check_sources(
        casetype_nuovi=[("Phone", "Call Assignment"), ("Non-live", "Collections")],
        casetype_esclusi=("Call Assignment",),
    )
    (f,) = trova(rep, "case type nuovi")

    assert "1 entrano nel trend" in f.summary
    entra = [d for d in f.details if "ENTRA" in d]
    assert entra == ["Non-live | Collections   -> ENTRA nel trend"]
    assert any("Call Assignment" in d and "escluso dal trend" in d for d in f.details)
    assert "casetype_esclusi" in f.hint


def test_se_tutti_i_nuovi_sono_esclusi_non_si_parla_di_trend():
    """Niente cambia nelle heat map: il suggerimento su come escluderli sarebbe
    rumore."""
    rep = check_sources(
        casetype_nuovi=[("Phone", "Call Assignment")],
        casetype_esclusi=("call assignment",),  # confronto insensibile al caso
    )
    (f,) = trova(rep, "case type nuovi")
    assert "entrano nel trend" not in f.summary
    assert "ENTRA" not in "".join(f.details)
    assert "casetype_esclusi" not in f.hint


def test_nessun_case_type_nuovo_non_dice_niente():
    assert trova(check_sources(casetype_nuovi=[]), "case type nuovi") == []
    assert trova(check_sources(casetype_nuovi=None), "case type nuovi") == []
