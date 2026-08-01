"""Golden test: le sorgenti WFM ricostruiscono i fogli del W30.

È il test che conta più di tutti gli altri. Le due sorgenti e il loro risultato
sono entrambi nel repo, quindi la semantica della trasformazione non si indovina:
si verifica contro ciò che il processo manuale ha effettivamente prodotto.

Le sole differenze ammesse sono dichiarate qui sotto e valgono **solo sulle
righe che le giustificano** — non come conteggi per campo, che nasconderebbero
una riga sbagliata dentro cento righe attese.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from fasterreports.core.names import normalize_name
from fasterreports.core.wfmsource import (
    read_alias_map,
    read_backoffice,
    read_roster,
    week_bounds,
)
from fasterreports.core.xlsxsource import read_sheet

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "samples" / "omni-report" / "Omni Report W30.xlsm"
ROSTER = ROOT / "samples" / "omni-report" / "sorgenti" / "Turni_W30.xlsx"
BACKOFFICE = ROOT / "samples" / "omni-report" / "sorgenti" / "Back_Office_Time_Final.xlsx"

MONDAY_W30 = 46223  # 20/07/2026
TOL = 1e-9


def _need(*paths):
    for p in paths:
        if not p.is_file():
            pytest.skip(f"campione assente: {p.name}")


def _num(v):
    if v is None or str(v).strip() == "":
        return None
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        return str(v).strip()


def _blank(v) -> bool:
    return v is None or (isinstance(v, str) and not v.strip())


def _same(a, b) -> bool:
    if _blank(a) and _blank(b):
        return True
    if _blank(a) or _blank(b):
        return False
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) < TOL
    return str(a).strip() == str(b).strip()


@pytest.fixture(scope="module")
def week():
    return week_bounds(MONDAY_W30)


@pytest.fixture(scope="module")
def aliases():
    _need(WORKBOOK)
    return read_alias_map(WORKBOOK)


@pytest.fixture(scope="module")
def roster(week):
    _need(ROSTER)
    return read_roster(ROSTER, week=week)


@pytest.fixture(scope="module")
def backoffice(week, roster, aliases):
    _need(BACKOFFICE)
    return read_backoffice(
        BACKOFFICE, week=week, aliases=aliases,
        allowed_keys=set(roster.notes.target_agents),
    )


@pytest.fixture(scope="module")
def wb_turni():
    _need(WORKBOOK)
    sheet = read_sheet(WORKBOOK, "Turni")
    out = {}
    for r in sorted(sheet.rows):
        if r == 1:
            continue
        row = sheet.row(r)
        nome = str(row.get("A", "")).strip()
        if not nome:
            continue
        out[(nome, int(_num(row.get("E"))))] = {
            "Team/Skill": row.get("B"), "Contratto": row.get("C"),
            "Ore/gg": _num(row.get("D")), "Stato": row.get("F"),
            "Inizio turno": _num(row.get("G")), "Fine turno": _num(row.get("H")),
            "chiave": row.get("I"),
        }
    return out


@pytest.fixture(scope="module")
def wb_slot():
    _need(WORKBOOK)
    sheet = read_sheet(WORKBOOK, "Slot Only Cases")
    out = {}
    for r in sorted(sheet.rows):
        if r == 1:
            continue
        row = sheet.row(r)
        key = str(row.get("A", "")).strip()
        if not key:
            continue
        out[(key, int(_num(row.get("B"))))] = {
            "Slot inizio": _num(row.get("C")), "Slot fine": _num(row.get("D")),
            "Stato BO": row.get("E"),
        }
    return out


def _as_dict_turni(src):
    return {
        (r[0], r[4]): {
            "Team/Skill": r[1], "Contratto": r[2], "Ore/gg": r[3], "Stato": r[5],
            "Inizio turno": r[6], "Fine turno": r[7], "chiave": r[8],
        }
        for r in src.data
    }


def _as_dict_slot(src):
    return {
        (r[0], r[1]): {"Slot inizio": r[2], "Slot fine": r[3], "Stato BO": r[4]}
        for r in src.data
    }


# --- Turni ------------------------------------------------------------------

def test_turni_stesse_righe(roster, wb_turni):
    mine = _as_dict_turni(roster)
    assert len(mine) == len(wb_turni) == 252
    assert set(mine) == set(wb_turni)


def test_turni_36_agenti_7_giorni(roster):
    mine = _as_dict_turni(roster)
    agenti = {k[0] for k in mine}
    giorni = {k[1] for k in mine}
    assert len(agenti) == 36
    assert len(giorni) == 7
    assert len(mine) == 36 * 7


def test_turni_campi_identici_sulle_righe_lavora(roster, wb_turni):
    """Sulle righe LAVORA nulla può divergere: sono quelle che il motore legge."""
    mine = _as_dict_turni(roster)
    diffs = []
    for key, m in mine.items():
        if m["Stato"] != "LAVORA":
            continue
        w = wb_turni[key]
        for f in ("Team/Skill", "Ore/gg", "Stato", "Inizio turno", "Fine turno", "chiave"):
            if not _same(m[f], w[f]):
                diffs.append(f"{key} {f}: ricostruito={m[f]!r} workbook={w[f]!r}")
    assert not diffs, "differenze su righe LAVORA:\n  " + "\n  ".join(diffs[:10])


def test_turni_differenze_solo_su_righe_non_lavora(roster, wb_turni):
    """Le celle spazzatura del processo manuale stanno solo sulle righe FERIE-OFF.

    Il workbook ci mette `Ore/gg`=8 e `Inizio turno`=1447 su righe che nessuna
    formula legge (tutti i SUMIFS/MINIFS filtrano Stato="LAVORA"). La pipeline
    scrive celle vuote.
    """
    mine = _as_dict_turni(roster)
    for key, m in mine.items():
        if m["Stato"] == "LAVORA":
            continue
        assert m["Ore/gg"] is None
        assert m["Inizio turno"] is None
        w = wb_turni[key]
        # E il workbook, su quelle righe, ha davvero valori non vuoti.
        assert not _blank(w["Ore/gg"])


def test_turni_solo_skill_hpo(roster, wb_turni):
    mine = _as_dict_turni(roster)
    assert {m["Team/Skill"] for m in mine.values()} == {"HPO"}
    assert {str(w["Team/Skill"]).strip() for w in wb_turni.values()} == {"HPO"}


def test_turni_chiave_coincide_con_quella_del_workbook(roster, wb_turni):
    """`normalize_name` contro le 151 chiavi vere delle righe LAVORA."""
    mine = _as_dict_turni(roster)
    checked = 0
    for key, m in mine.items():
        if m["Stato"] != "LAVORA":
            continue
        assert m["chiave"] == str(wb_turni[key]["chiave"]).strip()
        assert m["chiave"] == normalize_name(key[0])
        checked += 1
    assert checked == 151


def test_agente_con_skill_marcata_escluso_per_default(roster):
    """Nora Ed Dahir ha `HPO   *`: fuori da Turni, ma tracciata."""
    mine = _as_dict_turni(roster)
    assert "Nora Ed Dahir" not in {k[0] for k in mine}
    assert "nora ed dahir" in roster.notes.target_agents
    assert any("*" in s for s in roster.notes.skills_with_marker)


def test_include_marked_aggiunge_solo_quell_agente(week, wb_turni):
    _need(ROSTER)
    src = read_roster(ROSTER, week=week, include_marked=True)
    mine = _as_dict_turni(src)
    extra = set(mine) - set(wb_turni)
    assert {k[0] for k in extra} == {"Nora Ed Dahir"}
    assert len(extra) == 7  # sette giorni


# --- Slot Only Cases --------------------------------------------------------

def test_slot_stesse_righe(backoffice, wb_slot):
    mine = _as_dict_slot(backoffice)
    assert len(mine) == len(wb_slot) == 259
    assert set(mine) == set(wb_slot)


def test_slot_37_agenti(backoffice):
    mine = _as_dict_slot(backoffice)
    assert len({k[0] for k in mine}) == 37


def test_slot_orari_identici_dove_c_e_uno_slot(backoffice, wb_slot):
    mine = _as_dict_slot(backoffice)
    diffs = []
    for key, m in mine.items():
        if m["Slot inizio"] is None:
            continue
        w = wb_slot[key]
        for f in ("Slot inizio", "Slot fine", "Stato BO"):
            if not _same(m[f], w[f]):
                diffs.append(f"{key} {f}: ricostruito={m[f]!r} workbook={w[f]!r}")
    assert not diffs, "differenze sugli slot veri:\n  " + "\n  ".join(diffs[:10])


def test_slot_no_bot_scritto_esplicitamente(backoffice, wb_slot):
    """Dove il processo manuale lascia vuoto, la pipeline scrive `NO BOT`.

    Il VBA ha un ramo `If status <> "NO BOT"` che oggi non scatta mai. L'esito
    numerico non cambia — senza orari la riga viene scartata comunque — ma
    l'intenzione diventa leggibile.
    """
    mine = _as_dict_slot(backoffice)
    nobot = [k for k, m in mine.items() if m["Stato BO"] == "NO BOT"]
    assert nobot
    for key in nobot:
        assert _blank(wb_slot[key]["Stato BO"])
        assert mine[key]["Slot inizio"] is None


def test_alias_applicati_solo_per_raggiungere_lo_spazio_del_roster(backoffice, aliases):
    """La tabella alias mescola due direzioni: se ne usa una sola.

    `alessandro passierello` -> `passariello` serve (grafia back-office ->
    roster). `nadia ariefieva` -> `nadiia ariefieva` NO: porterebbe fuori dallo
    spazio del roster e la riga sparirebbe.
    """
    applied = backoffice.notes.aliases_applied
    assert "alessandro passierello" in applied
    assert "asia chirrullo" in applied
    assert "nadia ariefieva" not in applied
    assert "eleonora rosa sissa" not in applied
    # E gli agenti "non tradotti" ci sono comunque, con la grafia del roster.
    keys = {k[0] for k in _as_dict_slot(backoffice)}
    assert "nadia ariefieva" in keys
    assert "alessandro passariello" in keys


def test_agenti_fuori_target_esclusi(backoffice):
    """Il foglio back-office contiene anche agenti di altri team."""
    excluded = set(backoffice.notes.keys_not_allowed)
    assert "lucia serafini" in excluded  # HPS nel roster, non HPO
    assert excluded


# --- coerenza fra le due fonti ---------------------------------------------

def test_insiemi_agenti_differiscono_solo_per_l_agente_marcato(roster, backoffice):
    turni = {normalize_name(r[0]) for r in roster.data}
    slot = {r[0] for r in backoffice.data}
    assert slot - turni == {"nora ed dahir"}
    assert turni - slot == set()
