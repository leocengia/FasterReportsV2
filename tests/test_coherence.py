"""Controlli di coerenza fra le fonti.

Per ciascuno: che cosa trova, e se **blocca** o **segnala**. La distinzione è la
sostanza del modulo — un controllo che segnala quando dovrebbe bloccare lascia
passare numeri sbagliati.
"""

from __future__ import annotations

from datetime import date

import pytest

from fasterreports.core.coherence import BLOCCA, SEGNALA, check_sources
from fasterreports.core.wfmsource import SourceNotes


def _finding(rep, frammento):
    for f in rep.findings:
        if frammento in f.check:
            return f
    return None


def turni_row(nome, data=46223, stato="LAVORA", ore=8.0, inizio=0.375, fine=0.75):
    # [nome, skill, contratto, ore, data, stato, inizio, fine, chiave]
    from fasterreports.core.names import normalize_name

    return [nome, "HPO", None, ore, data, stato, inizio, fine, normalize_name(nome)]


def slot_row(chiave, data=46223, inizio=0.5, fine=0.52, stato="BOT"):
    return [chiave, data, inizio, fine, stato]


# --- skill -----------------------------------------------------------------

def test_skill_marcata_blocca_per_default():
    notes = SourceNotes(
        skills_seen={"HPO": 36, "HPO                *": 1},
        skills_with_marker={"HPO                *": 1},
        target_agents={"mario rossi": "HPO", "nora ed dahir": "HPO                *"},
    )
    rep = check_sources(roster_notes=notes)
    f = _finding(rep, "somigliano")
    assert f is not None and f.level == BLOCCA
    assert "nora ed dahir" in " ".join(f.details)
    assert not rep.ok
    # L'errore dice cosa fare, non solo cosa è successo.
    assert "include_marked_skills" in f.hint


def test_skill_marcata_solo_segnala_se_inclusa():
    notes = SourceNotes(
        skills_seen={"HPO": 36, "HPO                *": 1},
        skills_with_marker={"HPO                *": 1},
        target_agents={"nora ed dahir": "HPO                *"},
    )
    rep = check_sources(roster_notes=notes, include_marked=True)
    assert _finding(rep, "somigliano").level == SEGNALA
    assert rep.ok


def test_skill_pulite_non_generano_blocchi():
    notes = SourceNotes(skills_seen={"HPO": 36, "RETAIL": 27},
                        target_agents={"mario rossi": "HPO"})
    rep = check_sources(roster_notes=notes)
    assert _finding(rep, "somigliano") is None
    assert rep.ok


def test_hps_non_viene_confuso_con_hpo():
    """`HPS` esiste nel roster reale: non deve far scattare il controllo."""
    notes = SourceNotes(skills_seen={"HPO": 36, "HPS": 18}, target_agents={})
    rep = check_sources(roster_notes=notes)
    assert _finding(rep, "somigliano") is None


# --- blocchi ---------------------------------------------------------------

def test_agente_in_due_blocchi_blocca():
    notes = SourceNotes(
        skills_seen={"HPO": 2},
        agents_in_blocks={"mario rossi": ["A–F", "K–P"]},
        blocks=["A–F", "K–P"],
    )
    rep = check_sources(roster_notes=notes)
    f = _finding(rep, "piu' blocchi")
    assert f is not None and f.level == BLOCCA
    assert "righe DOPPIE" in f.hint


def test_agente_in_un_solo_blocco_va_bene():
    notes = SourceNotes(agents_in_blocks={"mario rossi": ["A–F", "A–F"]}, blocks=["A–F"])
    rep = check_sources(roster_notes=notes)
    assert _finding(rep, "piu' blocchi") is None


# --- insiemi di agenti ----------------------------------------------------

def test_slot_senza_turni_blocca():
    rep = check_sources(
        turni_rows=[turni_row("Mario Rossi")],
        slot_rows=[slot_row("mario rossi"), slot_row("nora ed dahir")],
    )
    f = _finding(rep, "slot ma senza turni")
    assert f is not None and f.level == BLOCCA
    assert f.details == ["nora ed dahir"]
    assert "orario di default" in f.hint


def test_turni_senza_slot_segnala():
    rep = check_sources(
        turni_rows=[turni_row("Mario Rossi"), turni_row("Lucia Verdi")],
        slot_rows=[slot_row("mario rossi")],
    )
    f = _finding(rep, "turni ma senza slot")
    assert f is not None and f.level == SEGNALA
    assert f.details == ["lucia verdi"]
    assert rep.ok


def test_insiemi_allineati_nessuna_segnalazione():
    rep = check_sources(
        turni_rows=[turni_row("Mario Rossi")], slot_rows=[slot_row("mario rossi")],
    )
    assert _finding(rep, "slot ma senza turni") is None
    assert _finding(rep, "turni ma senza slot") is None


# --- anagrafiche ----------------------------------------------------------

def test_agente_senza_email_blocca():
    rep = check_sources(
        turni_rows=[turni_row("Mario Rossi"), turni_row("Lucia Verdi")],
        email_agenti={"mario rossi"},
    )
    f = _finding(rep, "senza email")
    assert f is not None and f.level == BLOCCA
    assert f.details == ["lucia verdi"]
    assert "NESSUNA regola" in f.hint


def test_chi_ha_lavorato_casi_ma_non_ha_email_viene_segnalato():
    """Il caso trovato sulla W31 vera: un nome in SF_DATABASE e non in 'Email
    Agenti'. Nessun roster serve per vederlo — bastano i quattro dataset."""
    rep = check_sources(
        case_owners={"SF_DATABASE": ["Mario Rossi", "Leonardo Cengia", "Mario Rossi"]},
        email_agenti={"mario rossi"},
    )
    f = _finding(rep, "SF_DATABASE senza email")
    assert f is not None
    # SEGNALA, non BLOCCA: le regole sui casi escono comunque, con email vuota.
    assert f.level == SEGNALA
    assert f.details == ["leonardo cengia"]
    assert "Anagrafica" in f.hint


def test_case_owners_tutti_con_email_nessuna_segnalazione():
    rep = check_sources(
        case_owners={"SF_DATABASE": ["Mario Rossi"]},
        email_agenti={"mario rossi"},
    )
    assert _finding(rep, "SF_DATABASE senza email") is None


def test_case_owners_senza_email_agenti_non_inventa_un_problema():
    """Template assente = elenco email non leggibile. Segnalare 37 agenti
    "senza email" sarebbe rumore su un dato che non abbiamo."""
    rep = check_sources(case_owners={"SF_DATABASE": ["Mario Rossi"]}, email_agenti=set())
    assert _finding(rep, "SF_DATABASE senza email") is None


def test_case_owners_ogni_fonte_ha_la_sua_riga():
    rep = check_sources(
        case_owners={
            "SF_DATABASE": ["Leonardo Cengia"],
            "PSAT_DATASET": ["Lucia Verdi"],
        },
        email_agenti={"mario rossi"},
    )
    assert _finding(rep, "SF_DATABASE senza email") is not None
    assert _finding(rep, "PSAT_DATASET senza email") is not None


def test_agente_senza_contratto_solo_segnala():
    """Nessun calcolo dipende dal contratto: segnalare basta."""
    rep = check_sources(turni_rows=[turni_row("Mario Rossi")], contratti={})
    f = _finding(rep, "senza contratto")
    assert f is not None and f.level == SEGNALA
    assert rep.ok


# --- settimana ------------------------------------------------------------

def test_settimana_sbagliata_blocca():
    rep = check_sources(
        turni_rows=[turni_row("Mario Rossi", data=46216)],  # settimana prima
        at_dates=(date(2026, 7, 20), date(2026, 7, 26)),
    )
    f = _finding(rep, "settimana di Turni")
    assert f is not None and f.level == BLOCCA
    assert "settimana scorsa" in f.hint


def test_settimana_giusta_passa():
    rep = check_sources(
        turni_rows=[turni_row("Mario Rossi", data=46223)],
        at_dates=(date(2026, 7, 20), date(2026, 7, 26)),
    )
    assert _finding(rep, "settimana di Turni") is None


def test_giorni_scoperti_segnalano():
    rep = check_sources(
        turni_rows=[turni_row("Mario Rossi", data=46223),
                    turni_row("Mario Rossi", data=46224)],
        slot_rows=[slot_row("mario rossi", data=46223)],
    )
    f = _finding(rep, "giorni coperti")
    assert f is not None and f.level == SEGNALA


# --- plausibilita' dei turni ---------------------------------------------

@pytest.mark.parametrize(
    "ore,inizio,fine",
    [
        (0.0, 0.375, 0.375),   # zero ore
        (20.0, 0.1, 0.95),     # oltre 16
        (8.0, 0.75, 0.375),    # fine prima dell'inizio
    ],
)
def test_turni_implausibili_bloccano(ore, inizio, fine):
    rep = check_sources(
        turni_rows=[turni_row("Mario Rossi", ore=ore, inizio=inizio, fine=fine)]
    )
    f = _finding(rep, "implausibili")
    assert f is not None and f.level == BLOCCA


def test_lavora_con_campi_vuoti_blocca():
    row = turni_row("Mario Rossi")
    row[3] = None  # Ore/gg mancante su una riga LAVORA
    rep = check_sources(turni_rows=[row])
    f = _finding(rep, "implausibili")
    assert f is not None and f.level == BLOCCA
    assert "campi vuoti" in " ".join(f.details)


def test_ferie_off_con_campi_vuoti_va_bene():
    """Sulle righe non-LAVORA i campi vuoti sono il comportamento voluto."""
    rep = check_sources(turni_rows=[
        turni_row("Mario Rossi", stato="FERIE-OFF", ore=None, inizio=None, fine=None)
    ])
    assert _finding(rep, "implausibili") is None


def test_turno_valido_passa():
    rep = check_sources(turni_rows=[turni_row("Mario Rossi", ore=8.0, inizio=0.375, fine=0.75)])
    assert _finding(rep, "implausibili") is None


# --- segnalazioni informative -------------------------------------------

def test_stati_non_turno_segnalati():
    notes = SourceNotes(non_shift_states={"training": ["G3 (x)"], "flessibilita": ["H4 (y)"]})
    rep = check_sources(roster_notes=notes)
    f = _finding(rep, "non sono turni")
    assert f is not None and f.level == SEGNALA
    assert "training" in f.summary


def test_marcatori_segnalati():
    notes = SourceNotes(non_shift_states={"training": ["G3"]},
                        shift_markers={"O finale": 1839})
    rep = check_sources(roster_notes=notes)
    f = _finding(rep, "marcatori")
    assert f is not None and f.level == SEGNALA
    assert "1839" in f.summary


def test_cache_vecchia_segnalata():
    rep = check_sources(roster_notes=SourceNotes(stale_cache=True))
    f = _finding(rep, "valori in cache")
    assert f is not None and f.level == SEGNALA
    assert "ricalcola" in f.hint


def test_request_segnalate():
    rep = check_sources(backoffice_notes=SourceNotes(slot_requests=["CE46 (x)"]))
    f = _finding(rep, "cambio turno")
    assert f is not None and f.level == SEGNALA


def test_alias_mancanti_segnalati():
    rep = check_sources(slot_rows=[slot_row("mario rossi")], aliases_available=False)
    f = _finding(rep, "alias")
    assert f is not None and f.level == SEGNALA
    assert "template" in f.hint


# --- report ---------------------------------------------------------------

def test_report_vuoto_dice_che_va_bene():
    rep = check_sources()
    assert rep.ok
    assert "nessuna anomalia" in "\n".join(rep.render())


def test_render_include_livello_e_suggerimento():
    rep = check_sources(
        turni_rows=[turni_row("Mario Rossi")], email_agenti=set(),
    )
    text = "\n".join(rep.render())
    assert "[BLOCCA]" in text
    assert "->" in text  # il suggerimento è nel testo


def test_troncamento_dei_dettagli():
    rows = [turni_row(f"Agente {i}") for i in range(30)]
    rep = check_sources(turni_rows=rows, email_agenti=set())
    text = "\n".join(rep.render())
    assert "e altri" in text


# --- la settimana scelta e' quella dei dati? -------------------------------

W30 = (date(2026, 7, 20), date(2026, 7, 26))


def test_settimana_giusta_con_sbordo_minimo_segnala():
    """Il caso reale del W30: 3 righe su 26540 finiscono oltre il bordo."""
    times = ["2026-07-20 06:00:00"] * 100 + ["2026-07-26 20:00:00"]
    rep = check_sources(at_start_times=times, week=W30, timezone_offset_hours=9.0)
    f = _finding(rep, "settimana scelta")
    assert f is not None and f.level == SEGNALA
    assert rep.ok
    assert "9 ore" in f.hint  # spiega perche' e' normale


def test_settimana_tutta_dentro():
    times = ["2026-07-20 06:00:00", "2026-07-21 06:00:00"]
    rep = check_sources(at_start_times=times, week=W30, timezone_offset_hours=9.0)
    f = _finding(rep, "settimana scelta")
    assert f.level == SEGNALA and "tutte" in f.summary


def test_settimana_completamente_sbagliata_blocca():
    """`--week 29` con dati della 30: nessuna riga dentro."""
    times = ["2026-07-20 06:00:00"] * 50
    sbagliata = (date(2026, 7, 13), date(2026, 7, 19))
    rep = check_sources(at_start_times=times, week=sbagliata, timezone_offset_hours=9.0)
    f = _finding(rep, "settimana scelta")
    assert f.level == BLOCCA
    assert "NESSUNA riga" in f.summary
    assert not rep.ok


def test_troppi_dati_fuori_settimana_blocca():
    times = ["2026-07-20 06:00:00"] * 5 + ["2026-08-10 06:00:00"] * 5
    rep = check_sources(at_start_times=times, week=W30, timezone_offset_hours=9.0)
    f = _finding(rep, "settimana scelta")
    assert f.level == BLOCCA
    assert "50%" in f.summary


def test_offset_fuso_conta():
    """Senza l'offset, un turno serale di Seattle cadrebbe nel giorno sbagliato."""
    # 19/07 16:00 Seattle + 9h = 20/07 01:00 Milano -> dentro la W30.
    times = ["2026-07-19 16:00:00"] * 10
    con = check_sources(at_start_times=times, week=W30, timezone_offset_hours=9.0)
    senza = check_sources(at_start_times=times, week=W30, timezone_offset_hours=0.0)
    assert _finding(con, "settimana scelta").level == SEGNALA
    assert _finding(senza, "settimana scelta").level == BLOCCA


def test_settimana_dichiarata_coincide():
    rep = check_sources(week_declared=30, week_inferred=(2026, 30))
    f = _finding(rep, "settimana dedotta")
    assert f.level == SEGNALA and "coincide" in f.summary
    assert rep.ok


def test_settimana_dichiarata_diversa_blocca():
    """Numeri giusti col nome di un'altra settimana: problema che si scopre mesi dopo."""
    rep = check_sources(week_declared=29, week_inferred=(2026, 30))
    f = _finding(rep, "settimana dichiarata")
    assert f.level == BLOCCA
    assert "--week 30" in f.hint  # dice quale usare
    assert not rep.ok


def test_settimana_senza_dichiarazione_solo_informa():
    rep = check_sources(week_inferred=(2026, 30))
    f = _finding(rep, "settimana dedotta")
    assert f.level == SEGNALA and "30" in f.summary


def test_skill_non_esatta_in_turni_blocca():
    """Il guasto trovato sul workbook generato della W31.

    `include_marked_skills: true` includeva i 5 agenti con `HPO   *` in 'Turni'
    scrivendo la skill grezza. Il FILTER di 'Helper Turni' confronta per
    uguaglianza esatta, quindi restavano fuori dai calcoli: 'Turni' 37 agenti,
    'Helper Turni' 32. E i due controlli sugli insiemi di agenti erano verdi,
    perche' Turni e Slot combaciavano — la riga c'era, semplicemente non serviva
    a nulla.
    """
    rows = [turni_row("Mario Rossi"), turni_row("Lucia Verdi")]
    rows[1][1] = "HPO                *"
    rep = check_sources(turni_rows=rows, wanted_skills=("HPO",))
    f = _finding(rep, "FILTER non riconoscera")
    assert f is not None and f.level == BLOCCA
    assert "lucia verdi" in f.details[0]
    assert "uguaglianza esatta" in f.hint


def test_skill_esatta_non_segnala_nulla():
    rep = check_sources(turni_rows=[turni_row("Mario Rossi")], wanted_skills=("HPO",))
    assert _finding(rep, "FILTER non riconoscera") is None


def test_chi_ha_casi_ma_nessun_turno_viene_segnalato():
    """La terza direzione dello stesso guasto, misurata sulla W31.

    'lucia serafini' ha 1 caso in SF e nessun turno HPO: in 'Report Agenti'
    compare con 'Ore previste' = 0, e in AddLoginRows il suo orario atteso
    ripiega su defaultStart perche' un turno non c'e'.
    """
    rep = check_sources(
        turni_rows=[turni_row("Mario Rossi")],
        case_owners={"SF_DATABASE": ["Mario Rossi", "Lucia Serafini"]},
    )
    f = _finding(rep, "casi ma senza turno")
    assert f is not None and f.level == SEGNALA
    assert f.details == ["lucia serafini (da SF_DATABASE)"]
    assert "Ore previste" in f.hint
    assert rep.ok


def test_grafie_diverse_non_diventano_falsi_positivi():
    """Il roster scrive 'Eleonora Rosa Sissa', Salesforce 'Eleonora Sissa'.

    Senza la tabella alias questo controllo segnalerebbe quattro persone che
    hanno il loro turno — e il rumore fa ignorare i controlli.
    """
    rep = check_sources(
        turni_rows=[turni_row("Eleonora Rosa Sissa")],
        case_owners={"SF_DATABASE": ["Eleonora Sissa"]},
        aliases={"eleonora rosa sissa": "eleonora sissa"},
    )
    assert _finding(rep, "casi ma senza turno") is None


def test_senza_alias_la_grafia_diversa_si_vede():
    """Onesto in entrambe le direzioni: se la tabella alias non copre il caso,
    il controllo lo dice invece di tacere."""
    rep = check_sources(
        turni_rows=[turni_row("Eleonora Rosa Sissa")],
        case_owners={"SF_DATABASE": ["Eleonora Sissa"]},
    )
    assert _finding(rep, "casi ma senza turno") is not None
