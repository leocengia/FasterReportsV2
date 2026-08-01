"""Parser delle celle-turno del roster WFM.

Una cella rappresenta il turno di un agente in un giorno. La forma di base è

    inizio turno _ inizio pausa _ fine pausa _ fine turno      es. 0900_1300_1330_1630

ma nel file reale le forme sono **26**, enumerate una per una prima di scrivere
questo modulo (`tools/audit_shift_forms.py`). Alcune non sono affatto turni:

    Training        28 celle
    Flessibilità    13 celle

Se si trattassero come "turno non riconosciuto" si perderebbero; se si
trattassero come riposo si falserebbero le ore previste. Sono stati a sé.

Ci sono poi **marcatori** che accompagnano i turni e il cui significato non è
documentato: un prefisso `s` (144 celle), un prefisso `o` (2), un underscore
iniziale (57+), e una `O` finale (1500+). Il parser li **conserva** in
`markers` invece di scartarli: il preflight ne riporta il conteggio, così se un
giorno si scopre che `O` vuol dire straordinario il dato non è già stato buttato.

Regola di fondo, come nel resto del progetto: le forme osservate e capite si
assorbono, qualunque altra **blocca** con il valore nel messaggio. Un turno
interpretato male diventa ore previste sbagliate, che è esattamente l'errore
silenzioso che stiamo eliminando.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

# Stati che il foglio `Turni` conosce (colonna F). Ricavati dal W30: nel
# risultato compaiono solo LAVORA e FERIE-OFF.
STATO_LAVORA = "LAVORA"
STATO_FERIE_OFF = "FERIE-OFF"

KIND_LAVORA = "lavora"
KIND_RIPOSO = "riposo"
KIND_ASSENTE = "assente"  # cella con soli separatori: non schedulato
KIND_TRAINING = "training"
KIND_FLESSIBILITA = "flessibilita"

# Parole intere che non sono turni. Confrontate senza accenti né maiuscole.
_WORD_KINDS = {
    "off": KIND_RIPOSO,
    "training": KIND_TRAINING,
    "flessibilita": KIND_FLESSIBILITA,
}

# Marcatori di significato ignoto, conservati e contati.
MARKER_TRAILING_O = "O finale"
MARKER_PREFIX_S = "prefisso s"
MARKER_PREFIX_O = "prefisso o"
MARKER_LEADING_SEP = "separatore iniziale"

_HHMM = re.compile(r"^\d{4}$")
_SPLIT = re.compile(r"[_\s]+")


class ShiftParseError(ValueError):
    """Forma della cella non riconosciuta: si blocca, non si indovina."""

    def __init__(self, raw: str, reason: str, *, where: str = "") -> None:
        self.raw = raw
        self.reason = reason
        self.where = where
        loc = f" ({where})" if where else ""
        super().__init__(
            f"Cella-turno non interpretabile{loc}: {raw!r}\n"
            f"  Motivo: {reason}\n"
            f"  Forme ammesse: 'HHMM_HHMM_HHMM_HHMM' (turno con pausa), "
            f"'HHMM_HHMM' (turno senza pausa), 'Off', 'Training', "
            f"'Flessibilita', cella vuota.\n"
            f"  Se questa forma è legittima va aggiunta a core/shifts.py con il "
            f"suo significato: interpretarla a caso produrrebbe ore previste sbagliate."
        )


@dataclass(frozen=True)
class Shift:
    kind: str
    start: float | None = None
    break_start: float | None = None
    break_end: float | None = None
    end: float | None = None
    markers: tuple[str, ...] = ()
    raw: str = ""

    @property
    def is_lavora(self) -> bool:
        return self.kind == KIND_LAVORA

    @property
    def stato(self) -> str | None:
        """Il valore da scrivere in `Turni!F`.

        `Training` e `Flessibilità` non compaiono nel W30, quindi non si sa
        come li tratti il processo manuale: restituiscono None e il preflight li
        segnala. Vedi la domanda aperta in docs/audit-workbook-W30.md.
        """
        if self.kind == KIND_LAVORA:
            return STATO_LAVORA
        if self.kind == KIND_RIPOSO:
            return STATO_FERIE_OFF
        return None

    @property
    def net_hours(self) -> float | None:
        """Ore lavorate al netto della pausa.

        Formula ricavata dal confronto col W30 (Ahmed Afifi):
          0900_1300_1330_1630 -> 7,5h lorde − 0,5h pausa = 7h  (workbook: 7)
          0900_1400_    _     -> 5h, nessuna pausa            (workbook: 5)

        Arrotondamento a 2 decimali: il workbook ha `5.18` dove il conto esatto
        dà 5,18333… (Viktoriia Lavrinets, 20/07, turno 08:00–13:11). È l'unico
        valore non "tondo" della settimana, quindi 2 decimali è **inferito da un
        solo campione** e non si distingue da un troncamento: con più settimane
        va riverificato. Il golden test lo intercetterebbe.
        """
        if self.start is None or self.end is None:
            return None
        gross = self.end - self.start
        if gross < 0:  # turno a cavallo della mezzanotte
            gross += 1.0
        pause = 0.0
        if self.break_start is not None and self.break_end is not None:
            pause = self.break_end - self.break_start
            if pause < 0:
                pause += 1.0
        return round((gross - pause) * 24.0, 2)


def _fold(text: str) -> str:
    """Minuscolo senza accenti: 'Flessibilità' -> 'flessibilita'."""
    nfkd = unicodedata.normalize("NFKD", str(text).lower())
    return "".join(c for c in nfkd if not unicodedata.combining(c)).strip()


def _to_fraction(hhmm: str, raw: str, where: str) -> float:
    h, m = int(hhmm[:2]), int(hhmm[2:])
    # 2400 non compare nel file ma è una scrittura plausibile per la mezzanotte.
    if h == 24 and m == 0:
        return 1.0
    if not (0 <= h <= 23 and 0 <= m <= 59):
        raise ShiftParseError(raw, f"{hhmm!r} non è un orario valido (HHMM)", where=where)
    return (h * 3600 + m * 60) / 86400.0


def parse_shift(value, *, where: str = "") -> Shift:
    """Interpreta una cella-turno. Solleva `ShiftParseError` se non la riconosce."""
    if value is None:
        return Shift(kind=KIND_ASSENTE, raw="")
    raw = str(value)
    s = raw.strip()
    if not s:
        return Shift(kind=KIND_ASSENTE, raw=raw)

    markers: list[str] = []

    # --- marcatori, staccati prima di guardare la struttura ------------------
    # 'O' finale (1738 celle), con o senza spazio davanti: '..._1800 O',
    # 's..._2000O', 'OFF            O', '_    _    _     O'.
    # Va staccata prima del confronto con le parole, altrimenti 'OFF O' non
    # verrebbe riconosciuto come riposo. Ma non si stacca alla cieca: una parola
    # che finisse per O verrebbe mutilata. Quindi solo se ciò che resta è vuoto,
    # finisce con una cifra, o è una parola nota.
    m = re.match(r"^(.*?)\s*O$", s, re.S)
    if m:
        head = m.group(1)
        core = head.replace("_", "").replace(" ", "")
        if core == "" or core[-1:].isdigit() or _fold(core) in _WORD_KINDS:
            s = head.strip()
            markers.append(MARKER_TRAILING_O)

    # prefisso 's' / 'o' minuscolo davanti alle cifre
    if re.match(r"^[so]\d{4}", s):
        markers.append(MARKER_PREFIX_S if s[0] == "s" else MARKER_PREFIX_O)
        s = s[1:]

    # separatore iniziale: '_1000_1239_...' e '_    _    _    '
    if s.startswith("_"):
        markers.append(MARKER_LEADING_SEP)
        s = s[1:]

    # --- parole intere ------------------------------------------------------
    word = _fold(s).replace("_", "").replace(" ", "")
    if word in _WORD_KINDS:
        return Shift(kind=_WORD_KINDS[word], markers=tuple(markers), raw=raw)

    # --- struttura a campi --------------------------------------------------
    tokens = [t for t in _SPLIT.split(s) if t]
    if not tokens:
        # '    _    _    _     ' — solo separatori: non schedulato.
        return Shift(kind=KIND_ASSENTE, markers=tuple(markers), raw=raw)

    bad = [t for t in tokens if not _HHMM.match(t)]
    if bad:
        raise ShiftParseError(
            raw,
            f"campi non riconosciuti: {', '.join(repr(b) for b in bad)}",
            where=where,
        )

    times = [_to_fraction(t, raw, where) for t in tokens]
    if len(times) == 2:
        return Shift(
            kind=KIND_LAVORA, start=times[0], end=times[1],
            markers=tuple(markers), raw=raw,
        )
    if len(times) == 4:
        return Shift(
            kind=KIND_LAVORA, start=times[0], break_start=times[1],
            break_end=times[2], end=times[3], markers=tuple(markers), raw=raw,
        )
    raise ShiftParseError(
        raw,
        f"{len(times)} orari trovati, attesi 2 (senza pausa) o 4 (con pausa)",
        where=where,
    )


# ---------------------------------------------------------------------------
# Slot back-office: forma piu' semplice, due orari o uno stato.
# ---------------------------------------------------------------------------

SLOT_BOT = "BOT"
SLOT_NO_BOT = "NO BOT"
SLOT_REQUEST = "REQUEST"

KIND_SLOT = "slot"
KIND_SLOT_NO_BOT = "no_bot"
KIND_SLOT_REQUEST = "request"
KIND_SLOT_VUOTO = "vuoto"


@dataclass(frozen=True)
class Slot:
    kind: str
    start: float | None = None
    end: float | None = None
    raw: str = ""

    @property
    def stato_bo(self) -> str | None:
        """Il valore da scrivere in `Slot Only Cases!E`.

        Il W30 contiene `BOT` o vuoto, mai `NO BOT` — chi incolla lascia la
        cella vuota. Il VBA però ha un ramo esplicito
        `If status <> "NO BOT" ...`: la pipeline scrive `NO BOT`, così quel ramo
        fa quello per cui è stato scritto. L'esito numerico non cambia (senza
        orari la riga viene scartata comunque), ma l'intenzione diventa
        leggibile. Vedi docs/audit-workbook-W30.md.
        """
        if self.kind == KIND_SLOT:
            return SLOT_BOT
        if self.kind == KIND_SLOT_NO_BOT:
            return SLOT_NO_BOT
        return None


def parse_slot(value, *, where: str = "") -> Slot:
    """Interpreta una cella del foglio back-office.

    Le celle numeriche (conteggi e percentuali delle sezioni di calcolo) non
    sono slot: restituiscono `vuoto`. Il chiamante non deve passarle affatto —
    filtra le righe di totale prima — ma se ci arrivano non devono diventare
    orari per errore.
    """
    if value is None:
        return Slot(kind=KIND_SLOT_VUOTO, raw="")
    raw = str(value)
    s = raw.strip()
    if not s:
        return Slot(kind=KIND_SLOT_VUOTO, raw=raw)

    folded = " ".join(_fold(s).split())
    if folded == "no bot":
        return Slot(kind=KIND_SLOT_NO_BOT, raw=raw)
    if folded == "request":
        return Slot(kind=KIND_SLOT_REQUEST, raw=raw)

    tokens = [t for t in _SPLIT.split(s) if t]
    if len(tokens) == 2 and all(_HHMM.match(t) for t in tokens):
        return Slot(
            kind=KIND_SLOT,
            start=_to_fraction(tokens[0], raw, where),
            end=_to_fraction(tokens[1], raw, where),
            raw=raw,
        )

    # Numero puro: sezione di calcolo, non uno slot.
    try:
        float(s.replace(",", "."))
        return Slot(kind=KIND_SLOT_VUOTO, raw=raw)
    except ValueError:
        pass

    raise ShiftParseError(
        raw, "atteso 'HHMM_HHMM', 'NO BOT', 'REQUEST' o cella vuota", where=where
    )
