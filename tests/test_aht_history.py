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
