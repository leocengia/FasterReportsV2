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

    # Per ogni colonna target: da quale indice del CSV prendere e con che dtype.
    plan: list[tuple[int, int, str, str]] = []  # (offset_blocco, idx_sorgente, dtype, canonical)
    by_canonical = mapping.by_canonical()
    for fld in dataset.input_fields:
        res = by_canonical[fld.canonical]
        plan.append((fld.target_index - start, res.source_index, fld.dtype, fld.canonical))

    stats = {
        canonical: ColumnStats(canonical=canonical, target_col=index_to_col(start + off))
        for off, _, _, canonical in plan
    }

    out: list[list] = []
    for row in rows:
        line: list = [None] * width
        for off, src_idx, dtype, canonical in plan:
            st = stats[canonical]
            st.total += 1
            raw = row[src_idx] if src_idx < len(row) else None
            if C.is_blank(raw):
                st.blank += 1
                continue
            try:
                line[off] = C.coerce(raw, dtype)
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
    )


def _to_serial(value) -> float:
    from datetime import datetime

    if isinstance(value, datetime):
        return (value - C.EXCEL_EPOCH).total_seconds() / 86400.0
    return float(value)
