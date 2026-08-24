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


def test_segnala_i_case_type_non_in_helper_casetype():
    """Conseguenza limitata a un foglio: 'CaseType Deepdive'."""
    rep = check_sources(
        casetype_nuovi=[("Phone", "Call Assignment"), ("Non-live", "Collections")]
    )
    (f,) = trova(rep, "case type non in 'Helper CaseType'")
    assert f.level == SEGNALA
    assert "2 combinazioni" in f.summary
    assert any("Phone | Call Assignment" in d for d in f.details)
    assert "CaseType Deepdive" in f.hint
    assert rep.ok


def test_le_due_liste_sono_due_segnalazioni_diverse():
    """Sono due domande diverse, si rispondono in due file diversi.

    'Helper CaseType' (nel template) -> il dettaglio in 'CaseType Deepdive'.
    `aht_history.casetype_heatmap` (in settings.yml) -> le righe delle heat map.
    Confonderle e' l'errore da cui e' partita questa correzione.
    """
    rep = check_sources(
        casetype_nuovi=[("Phone", "Collections")],
        casetype_heatmap=("EVC",),
        casetype_pesi={("Phone", "Collections"): (1, 3.0)},
    )
    deepdive = trova(rep, "case type non in 'Helper CaseType'")
    heatmap = trova(rep, "case type fuori dalle heat map")
    assert len(deepdive) == 1 and len(heatmap) == 1
    assert "CaseType Deepdive" in deepdive[0].hint
    assert "AHT Trend WoW" in heatmap[0].hint
    assert "casetype_heatmap" in heatmap[0].hint


def test_i_case_type_fuori_dalle_heatmap_portano_i_numeri():
    """Chiedere una decisione senza dare i numeri e' chiedere di indovinare.

    Nella W33 'Call Assignment' aveva 30 casi: 'marginale' non e' una parola che
    si possa usare senza guardare quel numero.
    """
    rep = check_sources(
        casetype_heatmap=("EVC",),
        casetype_pesi={
            ("Phone", "Call Assignment"): (30, 4.5),
            ("Non-live", "Call Assignment"): (2, 6.0),
            ("Phone", "Collections"): (1, 12.0),
            ("Phone", "EVC"): (100, 15.0),
        },
    )
    (f,) = trova(rep, "case type fuori dalle heat map")
    # Per NOME: 'Call Assignment' e' una voce sola, con i due canali sommati.
    assert "2 case type" in f.summary and "33 casi in tutto" in f.summary
    assert f.details == [
        "Call Assignment   32 casi   <- MAI VISTO PRIMA",
        "Collections   1 casi   <- MAI VISTO PRIMA",
    ]
    # 'EVC' e' in lista: non compare.
    assert not any("EVC" in d for d in f.details)


def test_i_gia_decisi_stanno_in_fondo_e_i_nuovi_si_vedono():
    """La segnalazione che non cambia mai e' una segnalazione che si smette di leggere.

    `casetype_heatmap_esclusi` elenca i case type su cui la decisione e' gia'
    stata presa. Continuano a comparire — una scelta fatta a un caso alla
    settimana va riguardata se diventano trenta — ma separati, altrimenti la prima
    riga davvero nuova sparirebbe dentro un elenco che cresce e non cambia.
    """
    rep = check_sources(
        casetype_heatmap=("EVC",),
        casetype_heatmap_esclusi=("Collections", "Call Assignment"),
        casetype_pesi={
            ("Phone", "Call Assignment"): (30, 4.5),
            ("Phone", "Collections"): (1, 12.0),
            ("Phone", "Specialty Functions"): (8, 2.0),
        },
    )
    (f,) = trova(rep, "case type fuori dalle heat map")
    assert "di cui 1 mai visti prima" in f.summary
    assert f.details == [
        "Specialty Functions   8 casi   <- MAI VISTO PRIMA",
        "gia' decisi (aht_history.casetype_heatmap_esclusi):",
        "  Call Assignment   30 casi",
        "  Collections   1 casi",
    ]


def test_se_sono_tutti_gia_decisi_non_si_annuncia_niente_di_nuovo():
    rep = check_sources(
        casetype_heatmap=("EVC",),
        casetype_heatmap_esclusi=("Collections",),
        casetype_pesi={("Phone", "Collections"): (1, 12.0)},
    )
    (f,) = trova(rep, "case type fuori dalle heat map")
    assert "mai visti prima" not in f.summary
    assert not any("MAI VISTO PRIMA" in d for d in f.details)


def test_se_tutti_i_case_type_sono_in_lista_non_si_dice_niente():
    rep = check_sources(
        casetype_heatmap=("EVC",), casetype_pesi={("Phone", "EVC"): (10, 1.0)}
    )
    assert trova(rep, "case type fuori dalle heat map") == []


def test_nessun_case_type_nuovo_non_dice_niente():
    for arg in ([], None):
        rep = check_sources(casetype_nuovi=arg)
        assert trova(rep, "case type non in 'Helper CaseType'") == []
