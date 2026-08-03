"""Parser delle celle-turno.

Ogni forma qui sotto è stata **osservata** nel roster reale
(`tools/audit_shift_forms.py` sul W30 conta 26 schemi distinti). I casi che
devono fallire sono altrettanto importanti: un turno interpretato a caso diventa
ore previste sbagliate, senza errore.
"""

from __future__ import annotations

import pytest

from fasterreports.core.shifts import (
    KIND_ASSENTE,
    KIND_FLESSIBILITA,
    KIND_LAVORA,
    KIND_RIPOSO,
    KIND_TRAINING,
    MARKER_LEADING_SEP,
    MARKER_PREFIX_O,
    MARKER_PREFIX_S,
    MARKER_TRAILING_O,
    STATO_FERIE_OFF,
    STATO_LAVORA,
    ShiftParseError,
    parse_shift,
    parse_slot,
)

H = 1 / 24.0


def frac(hh: int, mm: int = 0) -> float:
    return (hh * 3600 + mm * 60) / 86400.0


# --- forme osservate nel file reale ----------------------------------------

@pytest.mark.parametrize(
    "raw,kind",
    [
        ("0900_1331_1431_1800", KIND_LAVORA),          # 3761 celle
        ("0830_1301_1401_1730 ", KIND_LAVORA),         # 3352 (spazio finale)
        ("0900_1300_1400_1800 O", KIND_LAVORA),        # 1530 (marcatore O)
        ("0800_1200_    _     ", KIND_LAVORA),         # 294 (senza pausa)
        ("1400_1800_    _", KIND_LAVORA),              # 202
        ("s0900_1230_1300_1700", KIND_LAVORA),         # 144 (prefisso s)
        ("s0900_1330_1430_2000O", KIND_LAVORA),        # 69 (prefisso + O attaccata)
        ("_1000_1239_1339_1800", KIND_LAVORA),         # 57 (underscore iniziale)
        ("1600_2100_    _     O", KIND_LAVORA),        # 43
        ("_1000_1300_    _    ", KIND_LAVORA),         # 24
        ("_1430_1800_    _", KIND_LAVORA),             # 21
        ("_0900_1300_1330_1530O", KIND_LAVORA),        # 8
        ("_0930_0900_    _    O", KIND_LAVORA),        # 4
        ("s0900_1400_    _    O", KIND_LAVORA),        # 4
        ("s0800_1200_    _    ", KIND_LAVORA),         # 3
        ("o1400 1430", KIND_LAVORA),                   # 1 (spazi invece di _)
        ("o1300 1700 1800 2200O", KIND_LAVORA),        # 1
        ("OFF", KIND_RIPOSO),                          # 5481
        ("Off", KIND_RIPOSO),
        ("off", KIND_RIPOSO),
        ("OFF            O", KIND_RIPOSO),             # 79
        ("Training", KIND_TRAINING),                   # 28
        ("Flessibilità", KIND_FLESSIBILITA),           # 3
        ("Flessibilità        ", KIND_FLESSIBILITA),   # 10
        ("    _    _    _     ", KIND_ASSENTE),        # 10
        ("    _    _    _     O", KIND_ASSENTE),       # 101
        ("_    _    _    _    ", KIND_ASSENTE),        # 3
        ("_    _    _    _", KIND_ASSENTE),            # 2
        ("", KIND_ASSENTE),
        (None, KIND_ASSENTE),
    ],
)
def test_forme_osservate(raw, kind):
    assert parse_shift(raw).kind == kind


# --- aritmetica, verificata contro il workbook -----------------------------

def test_turno_con_pausa():
    """Ahmed Afifi, 25/07: il workbook dice Ore/gg = 7."""
    s = parse_shift("0900_1300_1330_1630")
    assert s.start == pytest.approx(frac(9))
    assert s.break_start == pytest.approx(frac(13))
    assert s.break_end == pytest.approx(frac(13, 30))
    assert s.end == pytest.approx(frac(16, 30))
    assert s.net_hours == pytest.approx(7.0)   # 7,5h lorde − 0,5h pausa


def test_turno_senza_pausa():
    """Ahmed Afifi, 24/07: il workbook dice Ore/gg = 5."""
    s = parse_shift("0900_1400_    _     ")
    assert s.net_hours == pytest.approx(5.0)
    assert s.break_start is None and s.break_end is None


def test_turno_08_1630_con_pausa():
    """Ahmed Afifi, 26/07: il workbook dice Ore/gg = 8."""
    assert parse_shift("0800_1300_1330_1630").net_hours == pytest.approx(8.0)


def test_arrotondamento_a_due_decimali():
    """Viktoriia Lavrinets, 20/07: 08:00-13:11 = 5,1833h, il workbook ha 5.18."""
    assert parse_shift("0800_1311_    _    ").net_hours == pytest.approx(5.18)


def test_turno_a_cavallo_della_mezzanotte():
    # Non osservato nel W30, ma il conto non deve dare ore negative.
    s = parse_shift("2200_0200_    _    ")
    assert s.net_hours == pytest.approx(4.0)


def test_stato_scritto_nel_foglio():
    assert parse_shift("0900_1300_1330_1630").stato == STATO_LAVORA
    assert parse_shift("Off").stato == STATO_FERIE_OFF
    # Training e Flessibilità non compaiono nel W30: non si indovina.
    assert parse_shift("Training").stato is None
    assert parse_shift("Flessibilità").stato is None


def test_riposo_non_ha_ore():
    assert parse_shift("OFF").net_hours is None


# --- marcatori: conservati, non buttati ------------------------------------

@pytest.mark.parametrize(
    "raw,marker",
    [
        ("0900_1300_1400_1800 O", MARKER_TRAILING_O),
        ("s0900_1330_1430_2000O", MARKER_TRAILING_O),
        ("OFF            O", MARKER_TRAILING_O),
        ("s0900_1230_1300_1700", MARKER_PREFIX_S),
        ("o1400 1430", MARKER_PREFIX_O),
        ("_1000_1239_1339_1800", MARKER_LEADING_SEP),
    ],
)
def test_marcatori_conservati(raw, marker):
    assert marker in parse_shift(raw).markers


def test_off_non_viene_mutilato():
    """'OFF' finisce per F: nessuna O da staccare."""
    s = parse_shift("OFF")
    assert s.kind == KIND_RIPOSO
    assert s.markers == ()


def test_parola_che_finisce_per_o_non_viene_mutilata():
    """Una parola sconosciuta che finisce per O deve fallire, non perdere l'ultima lettera."""
    with pytest.raises(ShiftParseError):
        parse_shift("RIPOSO")


# --- fail loud -------------------------------------------------------------

@pytest.mark.parametrize(
    "raw",
    ["JOFF", "Ferie", "0900-1300", "0900_1300_1330", "abc_def", "9:00_13:00", "090_1300"],
)
def test_forme_ignote_bloccano(raw):
    with pytest.raises(ShiftParseError):
        parse_shift(raw)


def test_messaggio_di_errore_azionabile():
    with pytest.raises(ShiftParseError) as e:
        parse_shift("Ferie", where="G42 (Mario Rossi, 2026-07-20)")
    msg = str(e.value)
    assert "G42" in msg              # dove
    assert "Mario Rossi" in msg      # chi
    assert "'Ferie'" in msg          # cosa
    assert "Forme ammesse" in msg    # come si rimedia
    assert "ore previste sbagliate" in msg  # perché non si indovina


def test_orario_impossibile_blocca():
    with pytest.raises(ShiftParseError) as e:
        parse_shift("2599_2600_    _    ")
    assert "non è un orario valido" in str(e.value)


def test_tre_orari_bloccano():
    with pytest.raises(ShiftParseError) as e:
        parse_shift("0900_1300_1330")
    assert "attesi 2" in str(e.value)


# --- slot back office ------------------------------------------------------

def test_slot_orario():
    s = parse_slot("1000_1030")
    assert s.start == pytest.approx(frac(10))
    assert s.end == pytest.approx(frac(10, 30))
    assert s.stato_bo == "BOT"


def test_slot_no_bot():
    s = parse_slot("NO BOT")
    assert s.stato_bo == "NO BOT"
    assert s.start is None


def test_slot_request_riportato_come_nel_processo_manuale():
    """`REQUEST` = cambio turno pendente. Non e' uno slot, ma va scritto.

    Misurato sul W30: 7 celle `REQUEST` nel foglio del workbook. Scriverlo
    rende il foglio uguale a quello fatto a mano, e lascia visibile che quel
    giorno una richiesta era pendente.
    """
    slot = parse_slot("REQUEST")
    assert slot.stato_bo == "REQUEST"
    assert slot.start is None and slot.end is None


@pytest.mark.parametrize("raw", ["0", "10", "93", "71.48", "-9.19", "", None])
def test_slot_numerico_non_diventa_orario(raw):
    """Le sezioni di calcolo del foglio contengono conteggi e percentuali."""
    assert parse_slot(raw).stato_bo is None


def test_slot_etichetta_blocca():
    """Le righe-etichetta vanno saltate dal lettore, non interpretate qui."""
    with pytest.raises(ShiftParseError):
        parse_slot("Agents in Only Cases per Interval")
