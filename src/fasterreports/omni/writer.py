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

In fondo al file ci sono due scritture di natura diversa (`write_aht_history`,
`append_helper_casetype`): non riversano un file di input in un foglio, ma un
risultato calcolato dai dati di `SF_DATABASE` già letti. Vedi il commento che
le introduce per il perché non sono `Dataset` del contratto.
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


# ---------------------------------------------------------------------------
# I due fogli derivati: non vengono da un file di input, si calcolano dai dati
# di SF_DATABASE che sono gia' stati letti e validati.
#
# Per questo NON sono `Dataset` del contratto: un dataset, qui, e' una sorgente
# che si importa e si valida. Questi due sono un risultato, e il contratto non
# ha niente da controllare su di loro.
# ---------------------------------------------------------------------------

FOGLIO_STORICO = "AHT History"
FOGLIO_HELPER_CASETYPE = "Helper CaseType"
# Colonna di 'Helper CaseType' che porta una formula su ogni riga utilizzabile.
# Serve per misurare fin dove il template ha preparato le formule: scrivere una
# coppia (canale, case type) oltre quel punto la lascerebbe senza AHT, senza
# volume e senza categoria — invisibile, che e' il difetto che questo passaggio
# esiste per togliere.
COL_FORMULA_HELPER = "D"


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


def append_helper_casetype(book, nuove: list[tuple[str, str]]) -> WriteResult:
    """Aggiunge in fondo a 'Helper CaseType' le coppie (canale, case type) nuove.

    Scrive SOLO le colonne A e B, e SOLO sotto le righe gia' occupate: tutto il
    resto del foglio sono formule del template, e 'CaseType Deepdive' punta alle
    righe per posizione. Muovere una riga esistente vorrebbe dire spostare i
    case type sotto le etichette sbagliate nel deepdive, che l'utente ha
    formattato a mano.

    Si ferma dove finiscono le formule del template. Una coppia scritta oltre
    quel punto avrebbe nome e canale ma nessun numero accanto: sarebbe presente
    e invisibile insieme — cioe' lo stesso difetto che questo passaggio serve a
    eliminare. Meglio dirlo e non scriverla.
    """
    warnings: list[str] = []
    try:
        sht = book.sheets[FOGLIO_HELPER_CASETYPE]
    except Exception:
        # Non e' fatale: e' un foglio di supporto a un report secondario, e il
        # resto dell'Omni Report e' valido. Ma va detto.
        return WriteResult(
            dataset=FOGLIO_HELPER_CASETYPE,
            rows_written=0,
            rows_cleared=0,
            range_written="(foglio assente)",
            warnings=(
                f"Il template non contiene il foglio {FOGLIO_HELPER_CASETYPE!r}: "
                f"'CaseType Deepdive' non verra' aggiornato con i case type nuovi.",
            ),
        )

    if not nuove:
        return WriteResult(
            dataset=FOGLIO_HELPER_CASETYPE,
            rows_written=0,
            rows_cleared=0,
            range_written="(nessun case type nuovo)",
        )

    capienza = _capienza_helper(sht)
    prima_libera = _prima_riga_libera(sht)

    spazio = capienza - prima_libera + 1
    da_scrivere = nuove[:spazio] if spazio > 0 else []
    if len(da_scrivere) < len(nuove):
        escluse = nuove[len(da_scrivere):]
        warnings.append(
            f"{FOGLIO_HELPER_CASETYPE}: le formule del template arrivano a riga "
            f"{capienza} e lo spazio libero e' finito. {len(escluse)} case type NON "
            f"sono stati aggiunti: "
            + ", ".join(f"{c}|{t}" for c, t in escluse[:6])
            + (f" (+{len(escluse) - 6})" if len(escluse) > 6 else "")
            + f". Rimedio: trascina le formule di {FOGLIO_HELPER_CASETYPE} piu' in "
            f"basso (colonne C..AG) e riallarga i loro intervalli di ranking."
        )
    if not da_scrivere:
        return WriteResult(
            dataset=FOGLIO_HELPER_CASETYPE,
            rows_written=0,
            rows_cleared=0,
            range_written=f"(nessuno spazio libero entro riga {capienza})",
            warnings=tuple(warnings),
        )

    ultima = prima_libera + len(da_scrivere) - 1
    sht.range(f"A{prima_libera}:B{ultima}").value = [[c, t] for c, t in da_scrivere]

    return WriteResult(
        dataset=FOGLIO_HELPER_CASETYPE,
        rows_written=len(da_scrivere),
        rows_cleared=0,
        range_written=f"A{prima_libera}:B{ultima}",
        warnings=tuple(warnings),
    )


def leggi_coppie_helper(book) -> list[tuple[str, str]]:
    """Le coppie (canale, case type) gia' elencate in 'Helper CaseType'.

    La lista curata vive nel template, non in una config: e' l'utente che
    decide quali case type vuole vedere e in che ordine, e leggerla da li' vuol
    dire che riordinarla a mano continua a funzionare.
    """
    try:
        sht = book.sheets[FOGLIO_HELPER_CASETYPE]
    except Exception:
        return []
    ultima = _prima_riga_libera(sht) - 1
    if ultima < 2:
        return []
    valori = sht.range(f"A2:B{ultima}").value
    if ultima == 2:  # xlwings appiattisce una riga sola
        valori = [valori]
    out = []
    for riga in valori:
        if not isinstance(riga, (list, tuple)) or len(riga) < 2:
            continue
        canale, ct = riga[0], riga[1]
        if canale and ct and str(canale).strip() and str(ct).strip():
            out.append((str(canale).strip(), str(ct).strip()))
    return out


def _prima_riga_libera(sht) -> int:
    """La prima riga in cui la colonna B (case type) e' vuota."""
    ultima_b = sht.range(f"B{sht.cells.last_cell.row}").end("up").row
    return max(2, ultima_b + 1)


def _capienza_helper(sht) -> int:
    """Fin dove il template ha preparato le formule di 'Helper CaseType'.

    Si misura sulla colonna D, che porta una formula su ogni riga utilizzabile:
    l'ultima cella non vuota di quella colonna e' l'ultima riga in cui una
    coppia nuova viene effettivamente calcolata.
    """
    return sht.range(f"{COL_FORMULA_HELPER}{sht.cells.last_cell.row}").end("up").row


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
