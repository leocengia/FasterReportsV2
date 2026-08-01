"""`normalize_name` contro la formula del workbook.

La formula, verbatim da `Helper Turni!P2`:

    SUBSTITUTE(...(LOWER(TRIM($A2)),"'",""),"’",""),"à","a"),"á","a"),"è","e"),
    "é","e"),"ì","i"),"í","i"),"ò","o"),"ó","o"),"ù","u"),"ú","u")

Questi test la replicano caso per caso, inclusi gli accenti acuti che il
`NormKey` del VBA **non** gestisce.
"""

from __future__ import annotations

import pytest

from fasterreports.core.names import (
    VBA_MISSING_ACCENTS,
    excel_trim,
    full_name,
    has_marker,
    normalize_name,
    normalize_skill,
)


@pytest.mark.parametrize(
    "src,expected",
    [
        ("Ahmed Afifi", "ahmed afifi"),
        ("AHMED AFIFI", "ahmed afifi"),
        ("  Ahmed  Afifi  ", "ahmed afifi"),   # TRIM collassa gli spazi interni
        ("D'Angelo", "dangelo"),               # apostrofo diritto
        ("D’Angelo", "dangelo"),               # apostrofo tipografico
        ("Niccolò Verdà", "niccolo verda"),    # accenti gravi
        ("Cristóbal Martí", "cristobal marti"),  # accenti acuti
        ("Nadiia Ariefieva", "nadiia ariefieva"),
        ("", ""),
        ("   ", ""),
    ],
)
def test_normalize_name(src, expected):
    assert normalize_name(src) == expected


def test_none():
    assert normalize_name(None) == ""


@pytest.mark.parametrize("accent", VBA_MISSING_ACCENTS)
def test_accenti_acuti_che_il_vba_sbaglia(accent):
    """Documenta la divergenza, non la corregge nel VBA.

    Le formule mappano á í ó ú sulla vocale semplice, `NormKey` no. Finché il
    VBA resta com'è, un agente con questi accenti viene trovato dalle ore
    previste ma non dalle regole di malpractice.
    """
    nome = f"M{accent}rio Rossi"
    got = normalize_name(nome)
    assert accent not in got, f"{accent} non normalizzato: {got!r}"


def test_excel_trim_collassa_gli_spazi_interni():
    # Il TRIM di Excel fa questo; il Trim del VBA no (tocca solo i bordi).
    assert excel_trim("  a   b  ") == "a b"
    assert excel_trim("a\t b") == "a b"


def test_full_name():
    assert full_name("Ahmed", "Afifi") == "Ahmed Afifi"
    assert full_name("  Ahmed ", " Afifi ") == "Ahmed Afifi"
    assert full_name("Ahmed", None) == "Ahmed"
    assert full_name(None, None) == ""


def test_full_name_e_normalize_insieme():
    # È la catena che produce `Turni!A` e `Turni!I`.
    assert normalize_name(full_name("Ahmed", "Afifi")) == "ahmed afifi"


# --- skill ------------------------------------------------------------------

@pytest.mark.parametrize(
    "src,expected",
    [
        ("HPO", "hpo"),
        ("HPO                *", "hpo"),   # caso reale: Nora Ed Dahir
        ("hpo", "hpo"),
        ("HPO ", "hpo"),
        ("VRBO               *", "vrbo"),
        ("RELO               *", "relo"),
    ],
)
def test_normalize_skill(src, expected):
    assert normalize_skill(src) == expected


def test_hps_non_diventa_hpo():
    """`HPS` esiste nel blocco 2 del roster e somiglia a `HPO`.

    Se la normalizzazione li confondesse, 18 agenti di un altro team
    entrerebbero nel report.
    """
    assert normalize_skill("HPS") != normalize_skill("HPO")


def test_has_marker():
    assert has_marker("HPO                *")
    assert not has_marker("HPO")
    assert not has_marker(None)
