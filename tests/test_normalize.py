from __future__ import annotations

import pytest

from fasterreports.core.normalize import NormalizeRules, normalize, normalize_raw

RULES = NormalizeRules()


@pytest.mark.parametrize(
    "src,expected",
    [
        ("Agent Email", "agent email"),
        ("agent_email", "agent email"),
        ("  Agent   Email  ", "agent email"),
        ("AGENT_EMAIL", "agent email"),
        ("Total Time in seconds", "total time in seconds"),
        ("Total Time (s)", "total time s"),
        ("Productive Aux Flag (Yes / No)", "productive aux flag yes no"),
        ("Wrap-up Time in seconds", "wrap up time in seconds"),
        ("wrap_up_time_seconds", "wrap up time seconds"),
        ("Co-Browse Usage %", "co browse usage"),
        ("Case Origin (group)", "case origin group"),
        ("case_origin_group", "case origin group"),
        ("PSAT Score | Whole", "psat score whole"),
        ("﻿Date Viewpoint", "date viewpoint"),
    ],
)
def test_normalizza(src, expected):
    assert normalize(src, RULES) == expected


def test_punteggiatura_diventa_spazio_non_niente():
    # "Total Time (s)" -> "total time s": se la parentesi sparisse senza lasciare
    # spazio si otterrebbe "total time s" comunque, ma "A(B)" mostra la differenza.
    assert normalize("A(B)", RULES) == "a b"


def test_case_origin_e_case_origin_group_non_collassano():
    # Entrambe esistono in SF_DATABASE (col. T e U): se normalizzassero uguale,
    # ogni run sarebbe ambiguo su un file valido.
    assert normalize("Case Origin", RULES) != normalize("Case Origin (group)", RULES)


def test_agent_name_e_agent_name_underscore_collassano():
    # Collisione reale in PSAT_DATASET (col. I e BO). E' il motivo per cui il
    # matcher tenta il confronto grezzo prima di quello normalizzato.
    assert normalize("Agent Name", RULES) == normalize("agent_name", RULES)


def test_psat_score_collassa():
    # Collisione reale in PSAT_DATASET: 'psat_score' (DP, valori 0/1) e
    # 'PSAT Score' (EK, scala diversa).
    assert normalize("psat_score", RULES) == normalize("PSAT Score", RULES)


def test_normalize_raw_tollera_solo_bordi_e_bom():
    assert normalize_raw("  Agent Email  ") == "Agent Email"
    assert normalize_raw("﻿Agent Email") == "Agent Email"
    assert normalize_raw("agent_email") == "agent_email"  # resta distinto


def test_regole_disattivabili():
    r = NormalizeRules(lower=False, underscore_eq_space=False, strip_punct=())
    assert normalize("Agent_Email", r) == "Agent_Email"
