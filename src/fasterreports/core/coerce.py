"""Conversione dei valori testuali del CSV nei tipi che Excel si aspetta.

Perche' e' esplicita e non delegata: una colonna AHT che arriva come testo
produce medie sbagliate senza errore. La coercizione qui e' anche un controllo
di mappatura — se il 90% dei valori di quella che dovrebbe essere `Case AHT
(mins)` non e' numerico, la colonna agganciata e' quella sbagliata.

Locale: gli export arrivano indifferentemente con `1,5` o `1.5`. Si accettano
entrambi, ma NON si indovina sui separatori di migliaia ambigui (`1.234` resta
1.234, non 1234): meglio un numero riconoscibilmente strano che uno plausibile
e falso.
"""

from __future__ import annotations

import re
from datetime import date, datetime, time, timedelta

# Epoch del formato data di Excel (sistema 1900), con il salto del 1900 bisestile
# inesistente: 1899-12-30 e' l'origine effettiva per le date >= 1900-03-01.
EXCEL_EPOCH = datetime(1899, 12, 30)

_NUM_RE = re.compile(r"^[+-]?(\d+([.,]\d*)?|[.,]\d+)([eE][+-]?\d+)?$")
_PCT_RE = re.compile(r"^([+-]?[\d.,]+)\s*%$")

_DATETIME_FORMATS = (
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d %H:%M",
    "%Y-%m-%dT%H:%M:%S",
    "%Y-%m-%dT%H:%M:%SZ",
    "%Y-%m-%d",
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y %H:%M",
    "%d/%m/%Y",
    "%m/%d/%Y %H:%M:%S",
    "%m/%d/%Y %H:%M",
    "%m/%d/%Y",
    "%d-%m-%Y %H:%M:%S",
    "%d-%m-%Y",
    "%Y/%m/%d %H:%M:%S",
    "%Y/%m/%d",
    "%d/%m/%Y %I:%M:%S %p",
    "%m/%d/%Y %I:%M:%S %p",
)


class Uncoercible(Exception):
    """Il valore non e' convertibile nel dtype richiesto."""


def is_blank(value) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def to_str(value) -> str | None:
    if is_blank(value):
        return None
    return str(value).strip()


def to_float(value) -> float | None:
    if is_blank(value):
        return None
    if isinstance(value, (int, float)):
        return float(value)

    s = str(value).strip()
    # Percentuali: "12,5%" -> 0.125. Le colonne Co-Browse Usage % e
    # Misrouted Cases arrivano cosi' o come frazione, a seconda dell'export.
    m = _PCT_RE.match(s)
    scale = 1.0
    if m:
        s, scale = m.group(1), 0.01

    # Nessuno spazio interno viene rimosso: togliendoli, "1 2" diventerebbe 12.
    # Se un export iniziasse a usare lo spazio come separatore delle migliaia,
    # meglio che il preflight si fermi e che il caso venga gestito per bene.
    if not _NUM_RE.match(s):
        raise Uncoercible(str(value))
    # Un solo separatore decimale possibile: virgola o punto, non entrambi.
    if s.count(",") and s.count("."):
        raise Uncoercible(str(value))
    s = s.replace(",", ".")
    try:
        return float(s) * scale
    except ValueError:
        raise Uncoercible(str(value)) from None


def to_int(value) -> int | None:
    f = to_float(value)
    if f is None:
        return None
    if f != int(f):
        raise Uncoercible(str(value))
    return int(f)


def to_datetime(value) -> datetime | None:
    """Restituisce un datetime; xlwings lo scrive come data Excel."""
    if is_blank(value):
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime(value.year, value.month, value.day)

    # Un numero e' un seriale Excel: gli export che passano per Excel li hanno.
    if isinstance(value, (int, float)):
        return _from_excel_serial(float(value))

    s = str(value).strip()
    for fmt in _DATETIME_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    try:  # ISO con offset/microsecondi
        return datetime.fromisoformat(s.replace("Z", "+00:00"))
    except ValueError:
        pass
    if _NUM_RE.match(s.replace(",", ".")):
        return _from_excel_serial(float(s.replace(",", ".")))
    raise Uncoercible(str(value))


def _from_excel_serial(serial: float) -> datetime:
    if serial < 0 or serial > 2_958_466:  # ~anno 9999
        raise Uncoercible(str(serial))
    return EXCEL_EPOCH + timedelta(days=serial)


def to_excel_fraction(t: time | datetime) -> float:
    """Ora del giorno come frazione, la forma con cui Excel la tiene."""
    if isinstance(t, datetime):
        t = t.time()
    return (t.hour * 3600 + t.minute * 60 + t.second + t.microsecond / 1e6) / 86400.0


COERCERS = {
    "str": to_str,
    "float": to_float,
    "int": to_int,
    "datetime": to_datetime,
}


def coerce(value, dtype: str):
    try:
        fn = COERCERS[dtype]
    except KeyError:
        raise Uncoercible(f"dtype non supportato: {dtype}") from None
    return fn(value)
