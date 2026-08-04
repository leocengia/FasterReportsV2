"""Scrittura dei dataset nel workbook, via xlwings.

STATO: scritto secondo la struttura misurata del workbook W30, MAI ESEGUITO —
questo ambiente non ha Excel. Vedi docs/architettura.md §8 per la checklist
della prima esecuzione reale.

Perche' xlwings e non openpyxl: openpyxl in scrittura riscrive le formule ad
array dinamico e perde i valori in cache, cioe' corrompe il motore. Sul file WOW
da 60 MB va anche in out-of-memory (contesto WOW §5, verificato). openpyxl qui
non viene importato affatto.

Tre cose che il workbook reale impone e che il piano non menzionava:

1. `SF_DATABASE` ospita la tabella (ListObject) `AHT_Data`. Scrivere valori
   dentro l'intervallo di una tabella non la ridimensiona: se la settimana nuova
   ha un numero di righe diverso, la tabella resta della misura vecchia. Nel W30
   consegnato e' proprio cosi' (tabella a riga 3389, dati a 3568) e il foglio
   `Profilo Colonne SF`, che legge `ROWS(AHT_Data[])`, conta 179 righe in meno.
   Quindi: dopo la scrittura, `resize`.
2. Le formule `P`/`Q` di `AT_DATASET` sono pre-riempite fino a riga 130000. Se
   l'export ne porta di piu', le righe oltre restano senza Data/Ora Milano e il
   VBA le salta senza dirlo. Quindi: estenderle, o fermarsi.
3. Il VBA trova l'ultima riga con `End(xlUp)` su `AT_DATASET!B` e
   `SF_DATABASE!BB`. Righe vecchie rimaste sotto i dati nuovi verrebbero
   incluse nei conteggi. Quindi: pulire prima di scrivere, sempre.

`Turni` e `Slot Only Cases` passano dallo stesso percorso ma sono i casi
**semplici**: zero formule, nessun ListObject, nessuna colonna derivata. Per
loro `_clear_data` pulisce un intervallo contiguo (`A:I` e `A:E`) fino
all'ultima riga usata del foglio — che serve, perché nel W30 `Turni` è
dimensionato fino a riga 1141 mentre i dati sono 252: c'è molto spazio in cui
possono annidarsi residui di settimane precedenti, e il VBA li leggerebbe.
"""

from __future__ import annotations

from dataclasses import dataclass

from ..core.contract import Contract, Dataset, col_to_index, index_to_col
from ..core.errors import PipelineError
from ..core.transform import Block


@dataclass
class WriteResult:
    dataset: str
    rows_written: int
    rows_cleared: int
    range_written: str
    table_resized: str | None = None
    formulas_extended: str | None = None
    warnings: tuple[str, ...] = ()


def write_block(book, contract: Contract, dataset: Dataset, block: Block) -> WriteResult:
    """Pulisce i dati vecchi e scrive il blocco nel foglio del dataset."""
    try:
        sht = book.sheets[dataset.sheet]
    except Exception:
        raise PipelineError(
            f"Il template non contiene il foglio {dataset.sheet!r}. "
            f"Fogli presenti: {', '.join(s.name for s in book.sheets)}"
        ) from None

    warnings: list[str] = []
    first_row = dataset.header_row + 1
    start_idx = dataset.start_index

    rows_cleared = _clear_data(sht, dataset)

    n = block.n_rows
    if n:
        addr = (
            f"{block.start_col}{first_row}:"
            f"{block.end_col}{first_row + n - 1}"
        )
        sht.range(addr).value = block.rows
    else:
        addr = "(nessuna riga)"

    last_row = first_row + n - 1

    # --- 1. ListObject: riallinea la tabella ai dati -------------------------
    resized = None
    if dataset.list_object and n:
        resized = _resize_table(book, sht, dataset, last_row, warnings)

    # --- 2. formule derivate: estendile se i dati vanno oltre ----------------
    extended = None
    if dataset.derived_fields and dataset.max_template_row and n:
        if last_row > dataset.max_template_row:
            extended = _extend_derived_formulas(sht, dataset, last_row, warnings)

    return WriteResult(
        dataset=dataset.name,
        rows_written=n,
        rows_cleared=rows_cleared,
        range_written=addr,
        table_resized=resized,
        formulas_extended=extended,
        warnings=tuple(warnings),
    )


def _clear_data(sht, dataset: Dataset) -> int:
    """Svuota i dati sotto l'intestazione, senza toccare le formule derivate.

    Si pulisce fino all'ultima riga usata del foglio, non fino all'ultima riga
    scritta la volta prima: un export piu' corto lascerebbe in fondo righe della
    settimana precedente, e il VBA le conterebbe (legge End(xlUp)).
    """
    first_row = dataset.header_row + 1
    last_used = sht.used_range.last_cell.row
    if last_used < first_row:
        return 0

    derived_cols = {f.target_index for f in dataset.derived_fields}
    start = dataset.start_index
    end = col_to_index(dataset.data_end_col) if dataset.data_end_col else dataset.last_input_index

    # Blocchi contigui di colonne non derivate: cancellare le formule P/Q
    # significherebbe dover ricostruirle, e la fedelta' al template e' il punto.
    run: list[int] = []
    for idx in range(start, end + 1):
        if idx in derived_cols:
            _clear_run(sht, run, first_row, last_used)
            run = []
        else:
            run.append(idx)
    _clear_run(sht, run, first_row, last_used)

    return last_used - first_row + 1


def _clear_run(sht, run: list[int], first_row: int, last_row: int) -> None:
    if not run:
        return
    a, b = index_to_col(run[0]), index_to_col(run[-1])
    sht.range(f"{a}{first_row}:{b}{last_row}").clear_contents()


def _resize_table(book, sht, dataset: Dataset, last_row: int, warnings: list[str]) -> str | None:
    """Ridimensiona il ListObject all'estensione reale dei dati."""
    end_col = dataset.data_end_col or index_to_col(dataset.last_input_index)
    addr = f"${dataset.data_start_col}${dataset.header_row}:${end_col}${last_row}"
    try:
        # xlwings non espone i ListObject: si passa per l'API COM.
        lo = sht.api.ListObjects(dataset.list_object)
        lo.Resize(sht.api.Range(addr))
        return addr
    except Exception as exc:  # pragma: no cover — serve Excel
        warnings.append(
            f"{dataset.sheet}: impossibile ridimensionare la tabella "
            f"{dataset.list_object!r} a {addr} ({exc}). "
            f"I fogli che leggono ROWS({dataset.list_object}[]) — es. "
            f"'Profilo Colonne SF' — resteranno sul conteggio vecchio."
        )
        return None


def _extend_derived_formulas(sht, dataset: Dataset, last_row: int, warnings: list[str]) -> str | None:
    """Trascina in basso le formule delle colonne derivate.

    Si copia la formula della prima riga dati invece di scriverne una da
    stringa: cosi' la verita' resta nel template, non duplicata qui.
    """
    first_row = dataset.header_row + 1
    src_start = index_to_col(min(f.target_index for f in dataset.derived_fields))
    src_end = index_to_col(max(f.target_index for f in dataset.derived_fields))
    template_end = dataset.max_template_row or first_row
    try:
        src = sht.range(f"{src_start}{first_row}:{src_end}{first_row}")
        dst = sht.range(f"{src_start}{template_end + 1}:{src_end}{last_row}")
        src.copy(dst)
        addr = f"{src_start}{template_end + 1}:{src_end}{last_row}"
        warnings.append(
            f"{dataset.sheet}: i dati arrivano a riga {last_row}, oltre le "
            f"{template_end} righe di formule del template. Formule {src_start}/{src_end} "
            f"estese fino a {last_row} — da verificare a occhio la prima volta."
        )
        return addr
    except Exception as exc:  # pragma: no cover — serve Excel
        raise PipelineError(
            f"{dataset.sheet}: i dati arrivano a riga {last_row} ma le formule "
            f"{src_start}/{src_end} del template si fermano a {template_end}, e "
            f"l'estensione automatica e' fallita ({exc}).\n"
            f"  Senza Data/Ora Milano il VBA scarta silenziosamente quelle righe: "
            f"meglio fermarsi qui.\n"
            f"  Rimedio: estendi le formule nel template e aggiorna "
            f"max_template_row in columns.yml."
        ) from None
