"""Normalizzazione dei nomi di intestazione.

Serve a far combaciare `Wrap-up Time in seconds` con `wrap_up_time_seconds`.
Va usata SOLO come secondo tentativo: la normalizzazione perde informazione e
puo' far collassare colonne diverse sulla stessa stringa. Esempio reale
misurato in PSAT_DATASET: `Agent Name` (col. I) e `agent_name` (col. BO)
normalizzano entrambe su "agent name", ma sono colonne diverse. Per questo il
matcher tenta sempre prima il confronto sul nome grezzo.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

_WS = re.compile(r"\s+")


@dataclass(frozen=True)
class NormalizeRules:
    lower: bool = True
    strip: bool = True
    collapse_spaces: bool = True
    underscore_eq_space: bool = True
    strip_punct: tuple[str, ...] = field(default=("(", ")", "%", "–", "-", "/", "|", ".", ",", ":"))

    @classmethod
    def from_dict(cls, d: dict | None) -> NormalizeRules:
        d = d or {}
        return cls(
            lower=bool(d.get("lower", True)),
            strip=bool(d.get("strip", True)),
            collapse_spaces=bool(d.get("collapse_spaces", True)),
            underscore_eq_space=bool(d.get("underscore_eq_space", True)),
            strip_punct=tuple(d.get("strip_punct", cls.strip_punct)),
        )


def normalize(name: str, rules: NormalizeRules) -> str:
    s = str(name)
    # Il BOM di un CSV UTF-8 finisce nella prima intestazione: va via sempre.
    s = s.lstrip("﻿")
    if rules.lower:
        s = s.lower()
    if rules.underscore_eq_space:
        s = s.replace("_", " ")
    for p in rules.strip_punct:
        # La punteggiatura diventa spazio, non stringa vuota: cosi'
        # "Total Time (s)" -> "total time s" e non "total time s)" ne' "totaltimes".
        s = s.replace(p, " ")
    if rules.collapse_spaces:
        s = _WS.sub(" ", s)
    if rules.strip:
        s = s.strip()
    return s


def normalize_raw(name: str) -> str:
    """Confronto 'esatto': tollera solo BOM e spazi ai bordi.

    Un export che aggiunge uno spazio finale all'header non deve rompere il
    match esatto — non e' una rinomina, e' rumore. Tutto il resto (maiuscole,
    underscore, punteggiatura) resta significativo.
    """
    return str(name).lstrip("﻿").strip()
