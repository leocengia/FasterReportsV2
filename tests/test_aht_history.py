"""Lo storico settimanale AHT: aggregazione, memoria, idempotenza.

E' la parte del progetto in cui un errore non si vede: un volume contato due
volte o una settimana archiviata sotto il numero sbagliato producono un trend
plausibile e falso, e lo storico e' l'unica cosa che non si ricostruisce
rilanciando il programma — gli export delle settimane passate non ci sono piu'.
Quindi qui si prova tutto, e senza Excel: il modulo non ne ha bisogno.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from fasterreports.core.aht_history import (
    CANALE_LIVE,
    CANALE_NON_LIVE,
    COLONNE,
    RigaStorico,
    aggrega,
    carica,
    mappa_canale,
    righe_foglio,
    scrivi,
    settimane,
    unisci,
)
from fasterreports.core.errors import PipelineError


# ---------------------------------------------------------------------------
# Il canale: due vocabolari, una traduzione sola
# ---------------------------------------------------------------------------


@pytest.mark.parametrize(
    "raw, atteso",
    [
        ("Phone", CANALE_LIVE),
        ("phone", CANALE_LIVE),
        ("  PHONE  ", CANALE_LIVE),
        ("Other", CANALE_NON_LIVE),
        ("Email", CANALE_NON_LIVE),
        ("", CANALE_NON_LIVE),
        (None, CANALE_NON_LIVE),
    ],
)
def test_mappa_canale(raw, atteso):
    assert mappa_canale(raw) == atteso


def test_non_live_non_e_other():
    """L'etichetta dello storico e' 'Non-live', non 'Other'. Non e' un dettaglio
    estetico: le formule di 'AHT Trend WoW' filtrano per stringa esatta, e con
    'Other' non troverebbero niente — un trend vuoto senza errori."""
    assert CANALE_NON_LIVE == "Non-live"


# ---------------------------------------------------------------------------
# Aggregazione: volume e AHT si contano in modo diverso, e conta
# ---------------------------------------------------------------------------


def test_volume_conta_tutte_le_righe_aht_solo_quelle_numeriche():
    """Un caso senza AHT e' comunque un caso: entra nel volume. Ma contarlo come
    zero nella media abbasserebbe l'AHT di un case type piccolo senza motivo.
    Nel W33 reale erano 123 righe su 2833."""
    righe = [
        ("Phone", "Booking Information", 10.0),
        ("Phone", "Booking Information", 20.0),
        ("Phone", "Booking Information", ""),      # nessun AHT
        ("Phone", "Booking Information", None),    # nessun AHT
    ]
    (r,) = aggrega(righe, iso_year=2026, week=33)
    assert r.volume == 4
    assert r.aht == pytest.approx(15.0)


def test_aht_zero_se_nessuna_riga_lo_ha():
    righe = [("Other", "Collections", None), ("Other", "Collections", "")]
    (r,) = aggrega(righe, iso_year=2026, week=33)
    assert r.volume == 2
    assert r.aht == 0.0


def test_separa_i_canali():
    righe = [
        ("Phone", "EVC", 10.0),
        ("Other", "EVC", 30.0),
    ]
    out = aggrega(righe, iso_year=2026, week=33)
    assert {(r.channel, r.volume, r.aht) for r in out} == {
        ("Phone", 1, 10.0),
        ("Non-live", 1, 30.0),
    }


def test_scarta_le_righe_senza_case_type():
    """Un caso senza tipo non appartiene a nessun case type: sotto l'etichetta
    vuota diventerebbe una riga fantasma nello storico e nel foglio."""
    righe = [("Phone", "", 10.0), ("Phone", None, 10.0), ("Phone", "  ", 10.0),
             ("Phone", "EVC", 5.0)]
    out = aggrega(righe, iso_year=2026, week=33)
    assert [(r.case_type, r.volume) for r in out] == [("EVC", 1)]


def test_nessuna_combinazione_a_volume_zero():
    """Si aggregano le righe che ci sono, non il prodotto cartesiano canali x
    case type: una combinazione a volume zero non puo' nascere."""
    righe = [("Phone", "EVC", 5.0), ("Other", "Refund Request", 7.0)]
    out = aggrega(righe, iso_year=2026, week=33)
    assert all(r.volume > 0 for r in out)
    assert len(out) == 2


def test_accetta_aht_con_la_virgola_decimale():
    """Gli export arrivano indifferentemente con 1,5 o 1.5."""
    (r,) = aggrega([("Phone", "EVC", "12,5")], iso_year=2026, week=33)
    assert r.aht == pytest.approx(12.5)


def test_aht_non_numerico_non_conta_come_zero():
    righe = [("Phone", "EVC", "N/A"), ("Phone", "EVC", 10.0)]
    (r,) = aggrega(righe, iso_year=2026, week=33)
    assert r.volume == 2
    assert r.aht == pytest.approx(10.0)


def test_ordine_deterministico():
    """Due giri sugli stessi dati in ordine diverso danno lo stesso risultato:
    l'ordine in cui l'export elenca i casi non deve entrare nel foglio."""
    a = [("Phone", "Zeta", 1.0), ("Other", "Alfa", 2.0), ("Phone", "Alfa", 3.0)]
    assert aggrega(a, iso_year=2026, week=33) == aggrega(list(reversed(a)), iso_year=2026, week=33)


# ---------------------------------------------------------------------------
# La chiave dell'anno: senza, a gennaio il trend si rompe in silenzio
# ---------------------------------------------------------------------------


def test_week_key_tiene_separati_due_anni():
    """La W05 del 2027 e la W05 del 2026 sono settimane diverse. Sulla sola
    `week` un SUMIFS le sommerebbe insieme: volume raddoppiato e AHT mediata
    fra due anni, senza nessun errore visibile."""
    a = RigaStorico(2026, 5, "Phone", "EVC", 10, 12.0)
    b = RigaStorico(2027, 5, "Phone", "EVC", 20, 30.0)
    assert a.week_key != b.week_key
    assert a.chiave != b.chiave
    assert len(unisci([a], [b])) == 2


def test_week_key_ordina_bene_a_cavallo_danno():
    """Sulla sola settimana, la W1 del 2027 finirebbe SOTTO la W52 del 2026 e il
    trend mostrerebbe per mesi solo le settimane vecchie."""
    righe = [
        RigaStorico(2027, 1, "Phone", "EVC", 1, 1.0),
        RigaStorico(2026, 52, "Phone", "EVC", 1, 1.0),
        RigaStorico(2026, 53, "Phone", "EVC", 1, 1.0),
    ]
    assert settimane(righe) == [202652, 202653, 202701]


# ---------------------------------------------------------------------------
# Idempotenza: rigenerare la stessa settimana sostituisce, non raddoppia
# ---------------------------------------------------------------------------


def test_rigenerare_la_stessa_settimana_non_raddoppia():
    prima = aggrega([("Phone", "EVC", 10.0)], iso_year=2026, week=33)
    dopo = aggrega([("Phone", "EVC", 10.0), ("Phone", "EVC", 20.0)], iso_year=2026, week=33)

    uno = unisci([], prima)
    due = unisci(uno, dopo)

    assert len(due) == 1
    # Vince l'ultimo giro: e' quello fatto sull'export corretto.
    assert due[0].volume == 2
    assert due[0].aht == pytest.approx(15.0)


def test_unisci_e_stabile_su_piu_giri():
    righe = aggrega([("Phone", "EVC", 10.0), ("Other", "Alfa", 5.0)], iso_year=2026, week=33)
    uno = unisci([], righe)
    assert unisci(uno, righe) == uno == unisci(unisci(uno, righe), righe)


def test_unisci_non_cancella_le_settimane_vecchie():
    """Lo storico e' cumulativo. La finestra delle ultime 11 settimane la decide
    'AHT Trend WoW' con TAKE, non questo modulo."""
    vecchia = RigaStorico(2026, 22, "Phone", "EVC", 5, 10.0)
    nuova = RigaStorico(2026, 33, "Phone", "EVC", 7, 11.0)
    out = unisci([vecchia], [nuova])
    assert settimane(out) == [202622, 202633]


def test_unisci_non_cancella_un_case_type_che_sparisce():
    """Un case type senza casi questa settimana resta nello storico delle
    settimane in cui ne aveva: la sua riga vecchia non va toccata."""
    storico = [RigaStorico(2026, 32, "Phone", "Contract Update", 3, 20.0)]
    out = unisci(storico, [RigaStorico(2026, 33, "Phone", "EVC", 1, 5.0)])
    assert len(out) == 2


# ---------------------------------------------------------------------------
# Il file: andata e ritorno senza perdere niente
# ---------------------------------------------------------------------------


def test_storico_assente_e_uno_storico_vuoto(tmp_path):
    """La prima settimana in assoluto parte da zero, e non e' un guasto."""
    assert carica(tmp_path / "mai_scritto.csv") == []


def test_andata_e_ritorno(tmp_path):
    righe = aggrega(
        [("Phone", "EVC", 10.5), ("Other", "Alfa Beta, con virgola", 3.25)],
        iso_year=2026, week=33,
    )
    p = scrivi(tmp_path / "aht_history.csv", righe)
    assert carica(p) == righe


def test_scrittura_atomica_non_lascia_il_file_temporaneo(tmp_path):
    p = scrivi(tmp_path / "s.csv", [RigaStorico(2026, 33, "Phone", "EVC", 1, 1.0)])
    assert p.is_file()
    assert not (tmp_path / "s.csv.building").exists()


def test_intestazione_col_nome_giusto(tmp_path):
    p = scrivi(tmp_path / "s.csv", [RigaStorico(2026, 33, "Phone", "EVC", 1, 1.0)])
    assert p.read_text(encoding="utf-8").splitlines()[0] == ",".join(COLONNE)


def test_storico_con_colonne_sbagliate_si_ferma(tmp_path):
    p = tmp_path / "s.csv"
    p.write_text("week,channel,volume\n33,Phone,5\n", encoding="utf-8")
    with pytest.raises(PipelineError) as e:
        carica(p)
    assert "case_type" in str(e.value)


def test_storico_con_una_riga_illeggibile_si_ferma(tmp_path):
    """Meglio fermarsi che proseguire scartando in silenzio una settimana: lo
    storico e' memoria, e un buco non si recupera."""
    p = tmp_path / "s.csv"
    p.write_text(
        ",".join(COLONNE) + "\n"
        "33,Phone,EVC,5,12.0,2026,202633\n"
        "34,Phone,EVC,cinque,12.0,2026,202634\n",
        encoding="utf-8",
    )
    with pytest.raises(PipelineError) as e:
        carica(p)
    assert "riga 3" in str(e.value)


# ---------------------------------------------------------------------------
# Le righe per il foglio
# ---------------------------------------------------------------------------


def test_righe_foglio_ha_lintestazione_e_tutto_lo_storico():
    """Tutto, non le ultime 11 settimane: 'AHT Trend WoW'!B4 sceglie da se' le
    11 piu' recenti con TAKE, quindi tagliare qui butterebbe dati senza mostrare
    niente di piu'."""
    righe = [RigaStorico(2026, w, "Phone", "EVC", 1, 1.0) for w in range(1, 21)]
    out = righe_foglio(righe)
    assert out[0] == list(COLONNE)
    assert len(out) == 21
    assert [r[0] for r in out[1:]] == list(range(1, 21))


def test_righe_foglio_mette_la_settimana_come_numero():
    """La colonna A tiene un numero (33) e il formato "W"0 la fa leggere W33.
    Scriverci la stringa 'W33' romperebbe SUMIFS e SORT del template."""
    (_, riga) = righe_foglio([RigaStorico(2026, 33, "Phone", "EVC", 1, 1.0)])
    assert riga[0] == 33
    assert isinstance(riga[0], int)


def test_ordine_delle_colonne_del_foglio():
    (_, riga) = righe_foglio([RigaStorico(2026, 33, "Non-live", "EVC", 7, 12.5)])
    assert riga == [33, "Non-live", "EVC", 7, 12.5, 2026, 202633]


# ---------------------------------------------------------------------------
# I dati veri della W33: il conteggio deve tornare
# ---------------------------------------------------------------------------


def test_su_dati_reali_i_conti_tornano():
    """Numeri verificati a mano sull'export della W33: 828 casi Phone di
    'Supplier Initiated Traveler Contact' con AHT medio 10.44."""
    fixture = Path(__file__).parent / "fixtures" / "sf_w33_aggregati.csv"
    if not fixture.is_file():
        pytest.skip("fixture degli aggregati W33 non presente")
    import csv

    with open(fixture, newline="", encoding="utf-8") as f:
        atteso = {
            (r["channel"], r["case_type"]): (int(r["volume"]), float(r["aht"]))
            for r in csv.DictReader(f)
        }
    assert atteso[("Phone", "Supplier Initiated Traveler Contact")][0] == 828
    assert atteso[("Phone", "Supplier Initiated Traveler Contact")][1] == pytest.approx(
        10.4425, abs=1e-3
    )


# ---------------------------------------------------------------------------
# La lista delle heat map: l'archivio tiene tutto, la vista mostra i 27 scelti
#
# Fino al 2026-08-20 qui c'erano i test di `casetype_esclusi`, una lista di case
# type che `aggrega` teneva fuori DALL'ARCHIVIO. Sono stati sostituiti perche' il
# meccanismo era rovesciato: con una lista di esclusi, un case type mai visto
# prima entrava nelle heat map da se' — e' quello che e' successo nella W33.
# ---------------------------------------------------------------------------

HEATMAP = ["EVC", "Booking Information"]


def _st(*coppie, week=33):
    return [
        RigaStorico(2026, week, canale, ct, 10, 12.0) for canale, ct in coppie
    ]


def test_la_vista_tiene_solo_i_case_type_delle_heatmap():
    """`AHT Trend WoW`!A5 e' `UNIQUE(FILTER('AHT History'!C...))`: mostra tutto
    quello che trova nel foglio. Quindi e' il foglio che va filtrato."""
    from fasterreports.core.aht_history import filtra_heatmap

    storico = _st(("Phone", "EVC"), ("Phone", "Collections"), ("Non-live", "EVC"))
    tenute, fuori = filtra_heatmap(storico, HEATMAP)
    assert [(r.channel, r.case_type) for r in tenute] == [
        ("Non-live", "EVC"), ("Phone", "EVC"),
    ]
    assert fuori == ["Collections"]


def test_il_filtro_e_per_NOME_e_vale_per_entrambi_i_canali():
    """Nelle heat map del file legacy gli stessi 27 case type comparivano in
    entrambi i canali: un case type che c'e' per Phone e non per Non-live e'
    un'assenza di dati, non una scelta."""
    from fasterreports.core.aht_history import filtra_heatmap

    tenute, fuori = filtra_heatmap(
        _st(("Phone", "EVC"), ("Non-live", "EVC")), ["EVC"]
    )
    assert len(tenute) == 2 and fuori == []


def test_senza_lista_non_si_filtra_niente():
    from fasterreports.core.aht_history import filtra_heatmap

    storico = _st(("Phone", "EVC"), ("Phone", "Collections"))
    tenute, fuori = filtra_heatmap(storico, None)
    assert len(tenute) == 2 and fuori == []


def test_righe_foglio_applica_la_lista_delle_heatmap():
    from fasterreports.core.aht_history import righe_foglio

    righe = righe_foglio(_st(("Phone", "EVC"), ("Phone", "Collections")), HEATMAP)
    assert [r[2] for r in righe[1:]] == ["EVC"]


def test_aggrega_non_filtra_niente():
    """L'archivio prende tutto: e' l'unica cosa che non si ricostruisce.

    Fino al 2026-08-20 `aggrega` teneva fuori otto case type, e quei dati non
    esistono piu' per le settimane in cui e' girato cosi'.
    """
    righe = aggrega(
        [("Phone", "Call Assignment", 4.0), ("Phone", "EVC", 10.0)],
        iso_year=2026,
        week=33,
    )
    assert sorted(r.case_type for r in righe) == ["Call Assignment", "EVC"]


def test_l_archivio_tiene_tutto_e_la_lista_e_reversibile():
    """E' il guadagno del filtro a valle: aggiungere un case type alla lista lo
    fa comparire con TUTTE le settimane che l'archivio ha.

    Col filtro in `aggrega` avrebbe avuto un dato su dodici colonne — cioe'
    esattamente l'artefatto per cui il filtro esiste.
    """
    from fasterreports.core.aht_history import filtra_heatmap, righe_foglio

    archivio = (
        _st(("Phone", "Collections"), week=31)
        + _st(("Phone", "Collections"), week=32)
        + _st(("Phone", "Collections"), week=33)
    )
    assert righe_foglio(archivio, HEATMAP) == [list(COLONNE)]
    tenute, _ = filtra_heatmap(archivio, HEATMAP + ["Collections"])
    assert sorted(r.week for r in tenute) == [31, 32, 33]


def test_riscrivere_una_settimana_la_sostituisce_per_intero():
    """Il caso della W33: due coppie erano gia' finite nel CSV.

    Con la sostituzione riga per riga sarebbero rimaste per sempre — il giro
    nuovo non le produce, quindi non le sovrascrive, quindi nessuno le tocca
    piu'. Silenzioso e permanente.
    """
    prima = _st(("Phone", "EVC"), ("Phone", "Collections"), week=33)
    dopo = _st(("Phone", "EVC"), week=33)
    unito = unisci(prima, dopo)
    assert [(r.week, r.case_type) for r in unito] == [(33, "EVC")]


def test_riscrivere_una_settimana_non_tocca_le_altre():
    prima = _st(("Phone", "EVC"), week=32) + _st(("Phone", "Collections"), week=33)
    unito = unisci(prima, _st(("Phone", "EVC"), week=33))
    assert sorted((r.week, r.case_type) for r in unito) == [
        (32, "EVC"), (33, "EVC"),
    ]


def test_la_lista_delle_heatmap_copre_tutto_lo_storico():
    """La prova che la lista non fa sparire righe da undici settimane.

    Si verifica sui FILE VERI, `config/settings.yml` contro
    `data/aht_history.csv`: ogni case type dello storico deve essere nella lista
    delle heat map. Se non lo fosse, applicare il filtro farebbe sparire delle
    righe — e undici settimane di curatela a mano non si ricostruiscono.

    Misurato il 2026-08-20: la lista ha 27 nomi, lo storico W22..W32 ne usa
    esattamente 27, e sono gli stessi 27 delle heat map del file Excel legacy.
    Non sono i 31 di 'Helper CaseType': quel foglio serve a 'CaseType Deepdive',
    e i 4 in piu' (`Booking Research (Rates)`, `Bulk Update Request`,
    `Contract Update`, `Traveler Outreach`) non vanno nelle heat map.
    """
    import csv

    from fasterreports.omni.settings import load_settings

    ROOT = Path(__file__).resolve().parents[1]
    cfg = ROOT / "config" / "settings.yml"
    csv_path = ROOT / "data" / "aht_history.csv"
    if not (cfg.is_file() and csv_path.is_file()):
        pytest.skip("config o storico assenti")

    settings = load_settings(cfg)
    ammessi = {c.strip() for c in settings.casetype_heatmap}
    esclusi = {c.strip() for c in settings.casetype_heatmap_esclusi}
    assert len(ammessi) >= 20, f"lista sospettosamente corta: {len(ammessi)}"
    assert "Rates & Inventory Changes" in ammessi, (
        "il nome con la & non e' arrivato intero: il filtro butterebbe via un "
        "case type vero"
    )

    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        storico = {r["case_type"].strip() for r in csv.DictReader(f)}

    # Ogni case type dell'archivio dev'essere stato GUARDATO: o si vede nelle heat
    # map, o si e' deciso di tenerlo fuori. Quello che non deve succedere e' che
    # ne arrivi uno nuovo e nessuno se ne accorga — sparirebbe dalla vista senza
    # che nessuno l'abbia deciso. Per questo il test fallisce alla prima settimana
    # che ne porta uno inedito: e' il momento in cui la decisione va presa.
    ignoti = sorted(storico - ammessi - esclusi)
    assert not ignoti, (
        f"{len(ignoti)} case type dello storico non sono in nessuna delle due "
        f"liste di aht_history:\n"
        f"  {ignoti}\n"
        f"  Vanno in `casetype_heatmap` se li vuoi nelle heat map, in\n"
        f"  `casetype_heatmap_esclusi` se non li vuoi. L'archivio li tiene\n"
        f"  comunque: la scelta riguarda solo cosa si vede."
    )


def test_la_lista_delle_heatmap_non_e_quella_di_helper_casetype():
    """Sono due liste diverse, e confonderle e' l'errore che ho fatto.

    'Helper CaseType' alimenta 'CaseType Deepdive' e ne elenca 31; le heat map
    ne vogliono 27. I 4 di differenza entrerebbero nelle heat map senza che
    nessuno li abbia chiesti — che e' il difetto da cui e' partita la correzione.
    """
    from fasterreports.omni.orchestrate import _read_casetype_helper
    from fasterreports.omni.settings import load_settings

    ROOT = Path(__file__).resolve().parents[1]
    template = ROOT / "template" / "Omni_Report_TEMPLATE.xlsm"
    cfg = ROOT / "config" / "settings.yml"
    if not (template.is_file() and cfg.is_file()):
        pytest.skip("template o config assenti")

    helper = _read_casetype_helper(template)
    assert helper is not None
    nomi_helper = {t.strip() for _c, t in helper}
    ammessi = {c.strip() for c in load_settings(cfg).casetype_heatmap}

    assert ammessi < nomi_helper, (
        "la lista delle heat map dovrebbe essere un sottoinsieme PROPRIO di "
        "'Helper CaseType': se coincidessero, tanto valeva leggere da la'"
    )
    assert sorted(nomi_helper - ammessi) == [
        "Booking Research (Rates)",
        "Bulk Update Request",
        "Contract Update",
        "Traveler Outreach",
    ]
