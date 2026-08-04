"""Normalizzazione dei nomi agente — una regola, un posto.

Il workbook oggi la implementa in tre punti diversi, e due divergono:

  formule   `Helper Turni!P`, `Email Agenti!C`, `Anagrafica!B`
            LOWER(TRIM(x)) + rimozione di ' e ’ + 10 accenti (à á è é ì í ò ó ù ú)
  VBA       `CreaMalpractice.NormKey`
            LCase(Trim(x)) + rimozione di ' e ’ + SOLO 6 accenti (à è é ì ò ù)
            -> mancano á í ó ú
  a mano    chi produce `Turni!I` e `Slot Only Cases!A`, con regola ignota

`normalize_name` replica **la versione delle formule**, che è quella su cui si
basano i join di `Helper Turni`. È la regola che la pipeline usa per scrivere
`Turni!I` e `Slot Only Cases!A`.

Il `NormKey` del VBA NON viene toccato: allinearlo è una modifica al motore e va
fatta con il golden test in piedi, altrimenti non si distingue una correzione da
una regressione. Finché resta com'è, un agente con á/í/ó/ú viene trovato dalle
ore previste ma non dalle regole di malpractice.
"""

from __future__ import annotations

import re

# Nell'ordine delle SUBSTITUTE annidate della formula. L'ordine non conta
# (sono sostituzioni indipendenti di un carattere), ma tenerlo uguale rende il
# confronto con la formula immediato.
_APOSTROPHES = ("'", "’")  # ' e ’
_ACCENTS = (
    ("à", "a"),  # à
    ("á", "a"),  # á
    ("è", "e"),  # è
    ("é", "e"),  # é
    ("ì", "i"),  # ì
    ("í", "i"),  # í
    ("ò", "o"),  # ò
    ("ó", "o"),  # ó
    ("ù", "u"),  # ù
    ("ú", "u"),  # ú
)

# Accenti che il VBA NON gestisce: serve al test che documenta la divergenza.
VBA_MISSING_ACCENTS = ("á", "í", "ó", "ú")

_WS = re.compile(r"\s+")


def excel_trim(value: str) -> str:
    """TRIM di Excel: bordi *e* spazi interni multipli collassati a uno.

    Diverso dal `Trim` del VBA, che tocca solo i bordi — per quello il VBA
    chiama poi `WorksheetFunction.Trim`.
    """
    return _WS.sub(" ", str(value)).strip()


def normalize_name(value: str | None) -> str:
    """Nome agente -> chiave di join. Vuoto se il nome è vuoto."""
    if value is None:
        return ""
    s = excel_trim(value).lower()
    if not s:
        return ""
    for ap in _APOSTROPHES:
        s = s.replace(ap, "")
    for src, dst in _ACCENTS:
        s = s.replace(src, dst)
    return s


def full_name(first: str | None, last: str | None) -> str:
    """`Name` + `Surname` come li concatena il foglio `Turni`.

    Ricavato dal confronto sorgente/risultato sul W30: `Name`='Ahmed',
    `Surname`='Afifi' -> `Turni!A`='Ahmed Afifi'.
    """
    parts = [excel_trim(p) for p in (first, last) if p is not None and str(p).strip()]
    return " ".join(parts)


def normalize_skill(value: str | None) -> str:
    """Skill normalizzata per il *confronto*, non per la scrittura.

    Serve al controllo di coerenza che intercetta `HPO␣␣␣*` e `hpo`: entrambe
    normalizzano su `hpo` ma non sono `HPO` esatto, e il FILTER di
    `Helper Turni` usa l'uguaglianza esatta. Vedi docs/audit-workbook-W30.md.

    L'asterisco è un marcatore sistematico nel roster (`HPO␣*`, `VRBO␣*`,
    `RELO␣*`), non un errore di battitura: va rimosso per il confronto ma
    segnalato, mai ignorato in silenzio.
    """
    if value is None:
        return ""
    return excel_trim(str(value).replace("*", " ")).lower()


def has_marker(value: str | None) -> bool:
    """True se la skill porta l'asterisco."""
    return bool(value) and "*" in str(value)
