"""Costruzione del blocco canonico da scrivere nel foglio.

Output: una lista di righe, ognuna allineata alle colonne del foglio da
`data_start_col` fino all'ultima colonna input del dataset. Le celle non
mappate sono None (xlwings le lascia vuote). Le colonne `derived` non vengono
toccate: se le calcola il template, sono formule; se le calcola Python, ci
pensa `add_derived`.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from . import coerce as C
from .contract import Dataset, index_to_col
from .errors import CoercionError, EmptyDatasetError
from .matcher import DatasetMapping


@dataclass
class ColumnStats:
    canonical: str
    target_col: str
    total: int = 0
    blank: int = 0
    bad: int = 0
    examples: list[str] = field(default_factory=list)
    # Date che si leggono in due modi (`8/10/2026`) senza che il contratto dica
    # quale. Il valore c'e' e sembra buono, quindi non e' un `bad`: e' peggio,
    # perche' non lo vedrebbe nessuno. Il preflight lo SEGNALA.
    ambigue: int = 0
    esempi_ambigui: list[str] = field(default_factory=list)

    @property
    def considered(self) -> int:
        """Valori non vuoti: il denominatore giusto per la quota di scarti."""
        return self.total - self.blank

    @property
    def bad_ratio(self) -> float:
        return self.bad / self.considered if self.considered else 0.0


@dataclass
class Block:
    dataset: str
    start_col: str
    end_col: str
    rows: list[list]
    stats: dict[str, ColumnStats]
    # Righe della sorgente che non erano dati: la coda dei totali di un report
    # formattato, o righe senza chiave. Vanno RIPORTATE, non solo scartate — un
    # export che perde meta' delle righe perche' una colonna si e' spostata
    # produrrebbe lo stesso silenzio di sempre. Il preflight le stampa.
    dropped: list[str] = field(default_factory=list)
    # Le righe SOPRA l'intestazione nella sorgente (titolo, `As of <quando>`, i
    # filtri del report). Vuoto per un export normale. Le porta il writer nel
    # foglio, cosi' quelle righe dicono la settimana di questo giro e non quella
    # in cui e' stato costruito il template. Non le tocca `build_block`: non sono
    # dati, non hanno colonne canoniche, non si coercizzano.
    preamble: list[list] = field(default_factory=list)

    @property
    def n_rows(self) -> int:
        return len(self.rows)

    @property
    def n_cols(self) -> int:
        return len(self.rows[0]) if self.rows else 0


def build_block(
    dataset: Dataset,
    mapping: DatasetMapping,
    rows,
    *,
    max_uncoercible_ratio: float = 0.02,
    min_rows: int = 1,
) -> Block:
    """Trasforma le righe sorgente nel blocco allineato alle colonne target."""
    start = dataset.start_index
    end = dataset.last_input_index
    width = end - start + 1

    # Per ogni colonna target: da quale indice del CSV prendere, con che dtype e
    # con che formato di data (vuoto = si prova la lista dei formati noti).
    plan: list[tuple[int, int, str, str, str]] = []
    by_canonical = mapping.by_canonical()
    for fld in dataset.input_fields:
        res = by_canonical[fld.canonical]
        plan.append(
            (fld.target_index - start, res.source_index, fld.dtype, fld.canonical,
             fld.date_format)
        )

    stats = {
        canonical: ColumnStats(canonical=canonical, target_col=index_to_col(start + off))
        for off, _, _, canonical, _ in plan
    }

    # La colonna sorgente che rende una riga un RECORD. Se il contratto la
    # dichiara, una riga senza quel valore non e' un dato: e' la coda dei totali
    # (`Total | Sum | 72,13`, `Count | 178`) o una riga di separazione. E i
    # `stop_values` chiudono la tabella: dalla prima riga che ne contiene uno non
    # c'e' piu' niente di utile sotto.
    key_idx: int | None = None
    if dataset.key_field:
        key_idx = by_canonical[dataset.key_field].source_index
    stop = {s.casefold() for s in dataset.stop_values}

    out: list[list] = []
    dropped: list[str] = []
    for row in rows:
        if key_idx is not None:
            chiave = row[key_idx] if key_idx < len(row) else None
            testo = "" if chiave is None else str(chiave).strip()
            if testo.casefold() in stop:
                dropped.append(f"riga di chiusura ({dataset.key_field}={testo!r})")
                break
            if not testo:
                dropped.append(f"riga senza {dataset.key_field}")
                continue
        line: list = [None] * width
        for off, src_idx, dtype, canonical, fmt in plan:
            st = stats[canonical]
            st.total += 1
            raw = row[src_idx] if src_idx < len(row) else None
            if C.is_blank(raw):
                st.blank += 1
                continue
            # Solo se il contratto NON dichiara il formato: con `date_format` la
            # lettura e' decisa, non c'e' ambiguita' da segnalare.
            if dtype == "datetime" and not fmt and C.data_ambigua(raw):
                st.ambigue += 1
                if len(st.esempi_ambigui) < 4:
                    st.esempi_ambigui.append(str(raw)[:40])
            try:
                line[off] = C.coerce(raw, dtype, fmt)
            except C.Uncoercible:
                st.bad += 1
                if len(st.examples) < 8:
                    st.examples.append(str(raw)[:80])
                line[off] = None
        out.append(line)

    if len(out) < min_rows:
        raise EmptyDatasetError(
            f"Dataset {dataset.name}: {len(out)} righe dati, minimo richiesto {min_rows}.\n"
            f"  Il CSV ha intestazioni valide ma nessun dato utile: export incompleto?"
        )

    for st in stats.values():
        if st.bad_ratio > max_uncoercible_ratio:
            fld = next(f for f in dataset.input_fields if f.canonical == st.canonical)
            raise CoercionError(
                dataset.name, st.canonical, fld.dtype, st.bad_ratio,
                max_uncoercible_ratio, st.examples,
            )

    return Block(
        dataset=dataset.name,
        start_col=index_to_col(start),
        end_col=index_to_col(end),
        rows=out,
        stats=stats,
        dropped=dropped,
    )


def add_derived(dataset: Dataset, block: Block, offset_hours: float) -> Block:
    """Calcola in Python le colonne derivate (modalita' derived.mode=python).

    Replica esattamente le formule del template:
      P = IF(I="","", INT(I + offset/24))
      Q = IF(I="","", (I + offset/24) - INT(I + offset/24))

    Da usare solo come ottimizzazione: la modalita' `formula` e' la fedele,
    perche' la verita' resta una sola, nel workbook.
    """
    derived = dataset.derived_fields
    if not derived:
        return block

    src_names = {f.source for f in derived if f.source}
    if len(src_names) != 1:
        raise ValueError(
            f"Dataset {dataset.name}: i campi derivati devono discendere da un solo "
            f"campo input, trovati {sorted(n for n in src_names if n)}."
        )
    src_name = next(iter(src_names))

    start = dataset.start_index
    src_field = next(f for f in dataset.input_fields if f.canonical == src_name)
    src_off = src_field.target_index - start

    # Il blocco va allargato fino all'ultima colonna derivata.
    new_end = max(block.n_cols + start - 1, max(f.target_index for f in derived))
    width = new_end - start + 1
    shift = offset_hours / 24.0

    rows: list[list] = []
    for line in block.rows:
        line = list(line) + [None] * (width - len(line))
        base = line[src_off]
        if base is not None:
            serial = _to_serial(base) + shift
            day = int(serial // 1)
            for f in derived:
                off = f.target_index - start
                line[off] = day if f.canonical == "Data Milano" else serial - day
        rows.append(line)

    return Block(
        dataset=block.dataset,
        start_col=block.start_col,
        end_col=index_to_col(new_end),
        rows=rows,
        stats=block.stats,
        dropped=block.dropped,
        preamble=block.preamble,
    )


def _to_serial(value) -> float:
    from datetime import datetime

    if isinstance(value, datetime):
        return (value - C.EXCEL_EPOCH).total_seconds() / 86400.0
    return float(value)
