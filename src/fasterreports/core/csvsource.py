"""Lettura dei CSV sorgente.

Nessuna assunzione sull'ordine delle colonne: si restituiscono gli header cosi'
come sono e le righe come liste di stringhe. A decidere cosa va dove ci pensano
matcher e transform.

Perche' `csv` di libreria e non pandas: l'unico consumatore a valle e' xlwings,
che vuole una lista 2D di valori Python. Un DataFrame in mezzo aggiungerebbe una
dipendenza pesante, una inferenza di tipo che poi va comunque disfatta (la
coercizione qui e' esplicita e fail-loud) e un secondo giro di copie su 130k
righe. Se in futuro servisse pandas per l'analisi, si aggiunge a valle.
"""

from __future__ import annotations

import csv
import io
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator

from .errors import SourceError

# In ordine di tentativo. utf-8-sig prima di utf-8: mangia il BOM se c'e'.
ENCODINGS = ("utf-8-sig", "utf-8", "cp1252", "latin-1")
DELIMITERS = (",", ";", "\t", "|")

# Excel non accetta piu' di 1_048_576 righe: oltre, il problema e' a monte.
EXCEL_MAX_ROWS = 1_048_576


@dataclass
class CsvSource:
    path: Path
    headers: list[str]
    encoding: str
    delimiter: str
    _rows: Iterator[list[str]]

    def rows(self) -> Iterator[list[str]]:
        """Righe dati, in streaming. Consumabile una volta sola."""
        return self._rows

    @property
    def n_cols(self) -> int:
        return len(self.headers)


def sniff_encoding(path: Path) -> str:
    head = path.read_bytes()[:64_000]
    for enc in ENCODINGS:
        try:
            head.decode(enc)
        except UnicodeDecodeError:
            continue
        return enc
    # latin-1 decodifica qualunque byte, quindi qui non si arriva mai.
    raise SourceError(f"{path.name}: encoding non riconosciuto.")


def sniff_delimiter(first_line: str) -> str:
    """Il separatore e' quello che compare piu' volte fuori dalle virgolette.

    csv.Sniffer sbaglia sui campi di testo libero (i verbatim PSAT contengono
    virgole e punti e virgola), quindi si conta a mano ignorando le quote.
    """
    best, best_count = ",", -1
    for d in DELIMITERS:
        count = _count_unquoted(first_line, d)
        if count > best_count:
            best, best_count = d, count
    if best_count <= 0:
        raise SourceError(
            "Impossibile riconoscere il separatore della prima riga.\n"
            f"  Separatori provati: {', '.join(repr(d) for d in DELIMITERS)}\n"
            f"  Prima riga: {first_line[:200]!r}"
        )
    return best


def _count_unquoted(line: str, delim: str) -> int:
    count, in_quotes = 0, False
    for ch in line:
        if ch == '"':
            in_quotes = not in_quotes
        elif ch == delim and not in_quotes:
            count += 1
    return count


def read_csv(path: str | Path, *, encoding: str | None = None, delimiter: str | None = None) -> CsvSource:
    path = Path(path)
    if not path.is_file():
        raise SourceError(f"File sorgente non trovato: {path}")
    if path.stat().st_size == 0:
        raise SourceError(f"File sorgente vuoto: {path}")

    enc = encoding or sniff_encoding(path)
    fh = path.open("r", encoding=enc, newline="")
    first_line = fh.readline()
    if not first_line.strip():
        raise SourceError(f"{path.name}: la prima riga (intestazioni) e' vuota.")
    delim = delimiter or sniff_delimiter(first_line)

    fh.seek(0)
    reader = csv.reader(fh, delimiter=delim)
    try:
        headers = next(reader)
    except StopIteration:
        fh.close()
        raise SourceError(f"{path.name}: nessuna riga di intestazione.") from None

    headers = [h.strip() for h in headers]
    if headers and headers[0].startswith("﻿"):
        headers[0] = headers[0].lstrip("﻿")

    if not any(headers):
        fh.close()
        raise SourceError(f"{path.name}: la riga di intestazione non contiene nomi.")

    def iter_rows() -> Iterator[list[str]]:
        try:
            for row in reader:
                # Righe completamente vuote: coda del file, non dati.
                if not row or not any(c.strip() for c in row):
                    continue
                yield row
        finally:
            fh.close()

    return CsvSource(
        path=path,
        headers=headers,
        encoding=enc,
        delimiter=delim,
        _rows=iter_rows(),
    )


def read_csv_text(text: str, *, name: str = "<memoria>", delimiter: str | None = None) -> CsvSource:
    """Variante da stringa: serve ai test, evita file temporanei."""
    first_line = text.split("\n", 1)[0]
    if not first_line.strip():
        raise SourceError(f"{name}: la prima riga (intestazioni) e' vuota.")
    delim = delimiter or sniff_delimiter(first_line)
    reader = csv.reader(io.StringIO(text), delimiter=delim)
    headers = [h.strip().lstrip("﻿") for h in next(reader)]

    def iter_rows() -> Iterator[list[str]]:
        for row in reader:
            if not row or not any(c.strip() for c in row):
                continue
            yield row

    return CsvSource(
        path=Path(name),
        headers=headers,
        encoding="utf-8",
        delimiter=delim,
        _rows=iter_rows(),
    )
