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

In fondo al file c'è una scrittura di natura diversa (`write_aht_history`): non
riversa un file di input in un foglio, ma un risultato calcolato dai dati di
`SF_DATABASE` già letti. Vedi il commento che la introduce per il perché non è un
`Dataset` del contratto.
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
    # L'intervallo in cui e' stato riscritto il preambolo del download (solo per i
    # dataset con `header_row > 1`).
    preamble_written: str | None = None
    warnings: tuple[str, ...] = ()


def write_block(book, contract: Contract, dataset: Dataset, block: Block) -> WriteResult:
    """Pulisce i dati vecchi e scrive il blocco nel foglio del dataset."""
    try:
        sht = book.sheets[dataset.sheet]
    except Exception:
        if dataset.optional:
            # Una sezione opzionale che il template non ha: non e' un guasto, e'
            # un template piu' vecchio. Il blocco isolato (`dependent_sheets`)
            # semplicemente non c'e', e il resto dell'Omni Report e' intatto.
            # Fermarsi qui butterebbe un build valido per una sezione secondaria.
            return WriteResult(
                dataset=dataset.name,
                rows_written=0,
                rows_cleared=0,
                range_written="(foglio assente nel template)",
                warnings=(
                    f"Il template non contiene il foglio {dataset.sheet!r}: i dati di "
                    f"{dataset.name} non sono stati scritti da nessuna parte, e "
                    f"{', '.join(dataset.dependent_sheets) or 'i fogli che lo leggono'} "
                    f"non compariranno nel report. Il resto del workbook e' valido.",
                ),
            )
        raise PipelineError(
            f"Il template non contiene il foglio {dataset.sheet!r}. "
            f"Fogli presenti: {', '.join(s.name for s in book.sheets)}"
        ) from None

    warnings: list[str] = []
    first_row = dataset.header_row + 1
    start_idx = dataset.start_index

    rows_cleared = _clear_data(sht, dataset)
    preambolo = _write_preamble(sht, dataset, block, warnings)

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
        preamble_written=preambolo,
        warnings=tuple(warnings),
    )


def _write_preamble(sht, dataset: Dataset, block: Block, warnings: list[str]) -> str | None:
    """Ricopia nel foglio le righe che nel download stavano sopra l'intestazione.

    Perche' non lasciarle stare. Quelle righe sono il titolo del report, la riga
    `As of <quando>` e il blocco `Filtered By` — cioe' l'unico posto in cui il
    foglio dichiara **quale intervallo e' stato chiesto** al report. Se non le si
    riscrive, restano quelle del giorno in cui e' stato costruito il template: a
    dicembre il foglio direbbe ancora `As of 2026-08-13` e `Date/Time Closed
    greater or equal 8/3/2026`. Una data sbagliata che sembra giusta e' il difetto
    che questo progetto insegue da mesi.

    Le righe del preambolo sono allineate alla stessa griglia di colonne dei dati
    (quella che `read_table` ha ricavato dall'intestazione del sorgente): la prima
    cella di ogni riga di preambolo finisce in `data_start_col`. E' l'unica
    corrispondenza sensata — il preambolo non ha colonne proprie — e per il report
    SF e' anche quella letterale, perche' li' sia il titolo sia `Full Name` stanno
    in colonna B.

    NON si tocca la riga delle intestazioni. Quella e' del template, ed e' giusto:
    'Duplicates Helper' legge `DUP_DATASET` per POSIZIONE, quindi la riga 14 del
    template *e'* il contratto con quel foglio — e un rename nell'export lo becca
    il matcher, prima, nel preflight.
    """
    if dataset.header_row <= 1 or not block.preamble:
        return None

    capienza = dataset.header_row - 1
    righe = block.preamble[:capienza]
    if len(block.preamble) > capienza:
        warnings.append(
            f"{dataset.sheet}: il download ha {len(block.preamble)} righe sopra "
            f"l'intestazione, il foglio ne tiene {capienza}. Scritte le prime "
            f"{capienza}, le altre no.\n"
            f"  Vuol dire che il report ha piu' filtri di prima. I DATI sono a "
            f"posto — la pipeline li scrive sempre dalla riga "
            f"{dataset.header_row + 1} — ma il preambolo del foglio ora e' "
            f"incompleto: conviene allungarlo nel template (e aggiornare "
            f"header_row, con tutto quello che ne segue)."
        )

    fine_col = dataset.data_end_col or index_to_col(dataset.last_input_index)
    sht.range(f"{dataset.data_start_col}1:{fine_col}{capienza}").clear_contents()

    # Larghezza uniforme: xlwings vuole righe tutte della stessa lunghezza, e il
    # preambolo del sorgente ha righe corte (una cella) accanto a righe vuote.
    larghezza = col_to_index(fine_col) - dataset.start_index + 1
    normalizzate = [
        list(r[:larghezza]) + [None] * max(0, larghezza - len(r)) for r in righe
    ]
    if not normalizzate:
        return None
    sht.range(
        f"{dataset.data_start_col}1:{fine_col}{len(normalizzate)}"
    ).value = normalizzate
    return f"{dataset.data_start_col}1:{fine_col}{len(normalizzate)}"


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


# ---------------------------------------------------------------------------
# Il foglio derivato: non viene da un file di input, si calcola dai dati di
# SF_DATABASE che sono gia' stati letti e validati.
#
# Per questo NON e' un `Dataset` del contratto: un dataset, qui, e' una sorgente
# che si importa e si valida. Questo e' un risultato, e il contratto non ha
# niente da controllare su di lui.
#
# Fino al 2026-08-20 accanto c'era `append_helper_casetype`, che aggiungeva a
# 'Helper CaseType' le coppie (canale, case type) viste nei dati e non in lista.
# E' stata tolta: quella lista E' la lista curata dell'utente, e allargarla da
# se' faceva entrare nelle heat map di 'AHT Trend WoW' case type che nessuno
# aveva chiesto. Ora i fuori-lista li ELENCA il preflight, con volume e AHT, e la
# decisione resta a chi cura la lista.
# ---------------------------------------------------------------------------

FOGLIO_STORICO = "AHT History"


def write_aht_history(book, righe_foglio: list[list], warnings: list[str] | None = None) -> WriteResult:
    """Riscrive 'AHT History' da zero con lo storico completo.

    Da zero e non in append: lo storico vero e' il CSV fuori dal workbook, e il
    foglio ne e' solo una copia. Ricostruirlo ogni volta rende impossibile che
    le due cose divergano — che e' il guasto tipico di un foglio che si aggiorna
    a mano.

    I formati numerici NON sono scritti qui: si leggono dalla riga 2 del
    template e si propagano in basso. La colonna A tiene un numero (33) ma si
    legge `W33` grazie al formato `"W"0`, e la colonna E ha due decimali. Se
    quei formati fossero scritti in questo file diventerebbero una seconda
    verita' accanto al template, che e' esattamente il genere di duplicazione
    che questo progetto evita. Il rovescio della medaglia, dichiarato: se
    qualcuno svuota completamente il foglio nel template, la riga 2 non ha piu'
    un formato da copiare e i numeri appaiono grezzi.
    """
    warnings = warnings if warnings is not None else []
    try:
        sht = book.sheets[FOGLIO_STORICO]
    except Exception:
        raise PipelineError(
            f"Il template non contiene il foglio {FOGLIO_STORICO!r}, che serve al "
            f"trend settimanale.\n"
            f"  Fogli presenti: {', '.join(s.name for s in book.sheets)}\n"
            f"  Senza quel foglio 'AHT Trend WoW' non ha da dove leggere: va aggiunto "
            f"al template (intestazioni {', '.join(righe_foglio[0])})."
        ) from None

    n_col = len(righe_foglio[0])
    ultima_col = index_to_col(n_col)

    # Pulizia fino all'ultima riga usata: uno storico piu' corto di prima (puo'
    # capitare correggendo il CSV a mano) lascerebbe in fondo righe vecchie, e le
    # formule del trend leggono fino a riga 100000 — le vedrebbero.
    last_used = max(sht.used_range.last_cell.row, len(righe_foglio))
    sht.range(f"A1:{ultima_col}{last_used}").clear_contents()

    sht.range(f"A1:{ultima_col}{len(righe_foglio)}").value = righe_foglio
    ultima_riga = len(righe_foglio)

    _propaga_formati(sht, ultima_col, ultima_riga, warnings)

    return WriteResult(
        dataset=FOGLIO_STORICO,
        rows_written=ultima_riga - 1,  # senza l'intestazione
        rows_cleared=max(0, last_used - 1),
        range_written=f"A2:{ultima_col}{ultima_riga}",
        warnings=tuple(warnings),
    )


def _propaga_formati(sht, ultima_col: str, ultima_riga: int, warnings: list[str]) -> None:
    """Estende alle righe nuove i formati numerici della riga 2 del template."""
    if ultima_riga <= 2:
        return
    for idx in range(1, col_to_index(ultima_col) + 1):
        col = index_to_col(idx)
        try:
            fmt = sht.range(f"{col}2").number_format
            if fmt:
                sht.range(f"{col}3:{col}{ultima_riga}").number_format = fmt
        except Exception as exc:  # pragma: no cover — serve Excel
            warnings.append(
                f"{FOGLIO_STORICO}: non ho potuto propagare il formato della colonna "
                f"{col} alle righe nuove ({exc}). I numeri sono corretti, la loro "
                f"resa a schermo puo' essere grezza (es. 33 invece di W33)."
            )
            return


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
