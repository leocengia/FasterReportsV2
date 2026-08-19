"""Flusso end-to-end: CSV -> workbook finito.

STATO: la parte di preflight/ingestione gira e ha test. La parte Excel
(`build`) non e' mai stata eseguita: qui non c'e' Excel. Vedi
docs/architettura.md §8.

Ordine, e perche':
  0. il file di output non deve essere aperto in Excel (lock `~$...`):
     sovrascriverlo produrrebbe un salvataggio a meta' o un errore COM
     poco chiaro.
  1. preflight su TUTTI i CSV, prima di aprire Excel. Aprire Excel per poi
     scoprire che manca una colonna costa 30 secondi e lascia processi appesi.
  2. copia template (o il workbook esistente, per un giro parziale) -> una
     copia TEMPORANEA, mai out_path direttamente (`_scrivi_con_copia_atomica`):
     se un run va male a meta', out_path resta quello di prima.
  3. scrittura dei dataset nella copia temporanea.
  3b. i due fogli DERIVATI ('AHT History', le righe nuove di 'Helper CaseType'):
     non vengono da un file di input, si calcolano dai dati di SF_DATABASE
     appena scritti. Vanno qui, dopo il punto 3 — perche' il resize di
     `AHT_Data` deve essere gia' avvenuto — e prima del punto 4, perche'
     'AHT Trend WoW' e 'CaseType Deepdive' sono tutte formule e devono
     ricalcolare su questi valori, non su quelli della settimana precedente.
  4. ricalcolo completo. Le formule usano XLOOKUP/FILTER/UNIQUE in array e
     INDIRECT: volatili, un calculate() semplice non propaga sempre.
  5. macro malpractice, in modalita' silenziosa.
  6. salva e chiudi, sempre, anche in caso di errore; le chiamate xlwings piu'
     fragili (apertura, salvataggio) passano da `_con_retry`, che assorbe un
     errore COM transitorio senza disturbare l'utente.
  7. solo ora, con Excel chiuso e il salvataggio riuscito, la copia
     temporanea prende il nome buono (`os.replace`, atomico).
"""

from __future__ import annotations

import os
import shutil
import time
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from ..core.contract import Contract
from ..core.csvsource import read_csv
from ..core.errors import PipelineError
from ..core.preflight import PreflightReport, check_dataset
from ..core.transform import Block, add_derived, build_block
from ..core.wfmsource import week_from_iso
from .settings import Settings
from .writer import WriteResult, write_block


@dataclass
class BuildResult:
    workbook: Path | None
    preflight: Path
    report: PreflightReport
    writes: list[WriteResult] = field(default_factory=list)
    macro_ran: bool = False
    # Celle che dopo il ricalcolo contengono un errore Excel. Il workbook esiste,
    # ma con #SPILL!/#REF!/#VALUE! dentro i suoi numeri non sono affidabili: va
    # detto, non lasciato scoprire a chi lo apre.
    error_cells: list = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return (
            self.workbook is not None and self.report.ok and not self.error_cells
        )

    @property
    def n_errors(self) -> int:
        return sum(e.total for e in self.error_cells)


def _open_source(contract: Contract, settings: Settings, dataset, ctx: dict):
    """Apre la sorgente giusta per il dataset, secondo `reader` del contratto.

    I 4 CSV passano da `csvsource`; `Turni` e `Slot Only Cases` dagli adattatori
    WFM, che riducono una matrice larga a righe tidy. A valle non cambia nulla:
    entrambi restituiscono `(headers, rows)`.
    """
    path = settings.input_path(dataset.name)

    # Il formato si riconosce dall'estensione, non si dichiara: gli export reali
    # sono misti (SF e PSAT come .csv, AT e ATwi come .xlsx) e dipendono da chi
    # li produce. `reader:` nel contratto dice cosa *e'* il file — una tabella o
    # una matrice larga — non come e' scritto.
    if dataset.reader in ("csv", "table"):
        from ..core.tablesource import is_excel, read_table

        if is_excel(path):
            return read_table(path)
        return read_csv(path)

    from ..core.wfmsource import read_backoffice, read_roster

    week = ctx.get("week")
    if week is None:
        raise PipelineError(
            f"{dataset.name}: non so a quale settimana ritagliare la sorgente.\n"
            f"  Le sorgenti WFM coprono mesi (il roster del W30 va dal 20/07 al "
            f"15/09), quindi la settimana è obbligatoria.\n"
            f"  Di norma si ricava da AT_DATASET!Start Time; se AT_DATASET non è\n"
            f"  leggibile, imposta sources.monday_serial in settings.yml."
        )

    if dataset.reader == "wfm_roster":
        src = read_roster(
            path,
            week=week,
            skills=settings.sources.skills,
            include_marked=settings.sources.include_marked_skills,
            contratti=settings.contratti,
            sheet_name=settings.sources.roster_sheet,
        )
        ctx["roster_notes"] = src.notes
        ctx["target_agents"] = set(src.notes.target_agents)
        return src

    if dataset.reader == "wfm_backoffice":
        src = read_backoffice(
            path,
            week=week,
            sections=settings.sources.backoffice_sections,
            aliases=ctx.get("aliases"),
            allowed_keys=ctx.get("target_agents"),
            sheet_name=settings.sources.backoffice_sheet,
        )
        ctx["backoffice_notes"] = src.notes
        return src

    raise PipelineError(f"{dataset.name}: reader {dataset.reader!r} non implementato.")


def _dataset_order(contract: Contract) -> list:
    """I dataset in ordine di dipendenza.

    `AT_DATASET` per primo: da lui si ricava la settimana. Poi il roster, che
    produce l'elenco degli agenti. Poi il back office, che ci si filtra sopra.
    """
    rank = {"csv": 0, "wfm_roster": 1, "wfm_backoffice": 2}
    return sorted(
        contract.datasets.values(),
        key=lambda d: (rank.get(d.reader, 9), 0 if d.name == "AT_DATASET" else 1, d.name),
    )


def _empty_block(dataset) -> Block:
    """Il blocco di un dataset opzionale la cui fonte manca.

    Zero righe, ma un Block vero: passa dallo stesso `write_block` di tutti gli
    altri, che quindi PULISCE il foglio (`_clear_data` gira comunque) e non ci
    scrive nulla sopra. E' la differenza fra "vuoto perche' cosi' deve essere"
    e "vuoto per omissione, con sotto i dati della settimana prima".
    """
    from ..core.contract import index_to_col

    return Block(
        dataset=dataset.name,
        start_col=index_to_col(dataset.start_index),
        end_col=index_to_col(dataset.last_input_index),
        rows=[],
        stats={},
    )


def _source_label(settings: Settings, dataset: str) -> str:
    """Da dove *doveva* arrivare il dataset, senza poter fallire.

    Serve nel ramo d'errore: se la sorgente manca, `input_path` solleva — ed e'
    giusto che lo faccia. Ma chiamarla proprio mentre si scrive il rapporto
    dell'errore faceva morire tutto il preflight sulla prima fonte assente,
    invece di elencare le sei righe e dire quali mancano. Il rapporto deve
    sopravvivere al guasto che sta descrivendo.
    """
    try:
        return str(settings.input_path(dataset))
    except PipelineError:
        pattern = settings.input_files.get(dataset)
        if pattern:
            return f"{settings.input_dir / pattern} (nessun file corrispondente)"
        return f"{settings.input_dir} (pattern non configurato)"


def run_preflight(
    contract: Contract,
    settings: Settings,
    *,
    week_number: int | None = None,
    build_blocks: bool = True,
    only: set[str] | None = None,
) -> tuple[PreflightReport, dict[str, Block]]:
    """Valida le sorgenti e, se richiesto, costruisce i blocchi canonici.

    Non tocca Excel: si puo' (e si deve) lanciare anche solo per controllare che
    gli export della settimana siano a posto.
    """
    report = PreflightReport(generated_at=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    blocks: dict[str, Block] = {}
    ctx: dict = {}

    if settings.sources.monday_serial:
        from ..core.wfmsource import week_bounds

        ctx["week"] = week_bounds(settings.sources.monday_serial)

    # La tabella alias vive nel template, dov'e' anche il VBA che la usa.
    if settings.template.is_file():
        try:
            from ..core.wfmsource import read_alias_map

            ctx["aliases"] = read_alias_map(settings.template)
        except PipelineError:
            ctx["aliases"] = {}

    for dataset in _dataset_order(contract):
        name = dataset.name
        if only and name not in only:
            continue
        try:
            source = _open_source(contract, settings, dataset, ctx)
        except PipelineError as exc:
            from ..core.preflight import DatasetReport

            if dataset.optional:
                # La fonte non c'e', ma il dataset e' dichiarato opzionale (nel
                # contratto, con la prova che nessuna formula del VBA lo legge):
                # si segnala forte e si procede. Il blocco vuoto e' cio' che fa
                # scrivere il foglio VUOTO invece di lasciarlo saltato — un
                # residuo della settimana scorsa nel template sarebbe un
                # report sbagliato, non uno che manca.
                report.datasets.append(DatasetReport(
                    name=name,
                    source=_source_label(settings, name),
                    skipped_reason=str(exc),
                ))
                if build_blocks:
                    blocks[name] = _empty_block(dataset)
                continue

            report.datasets.append(
                DatasetReport(name=name, source=_source_label(settings, name), error=str(exc))
            )
            continue

        ds_report = check_dataset(contract, dataset, source)
        ds_report.notes = getattr(source, "notes", None)
        report.datasets.append(ds_report)

        if ds_report.ok and build_blocks:
            assert ds_report.mapping is not None
            try:
                block = build_block(
                    dataset,
                    ds_report.mapping,
                    source.rows(),
                    max_uncoercible_ratio=settings.validation.max_uncoercible_ratio,
                    min_rows=settings.validation.min_rows,
                )
            except PipelineError as exc:
                ds_report.error = str(exc)
                continue
            ds_report.block = block
            blocks[name] = block

            # La settimana: dal numero passato a --week piu' l'anno ricavato dai
            # dati. NON dall'intervallo di AT_DATASET, che sborda (vedi
            # `_resolve_week`).
            if name == "AT_DATASET":
                ctx["at_start_times"] = _at_start_times(dataset, block)
                ctx["offset_hours"] = _timezone_offset(contract, settings)
                inferred = _resolve_week(
                    ctx["at_start_times"], ctx["offset_hours"]
                )
                ctx["week_inferred"] = inferred
                ctx["week_declared"] = week_number
                if "week" not in ctx and inferred:
                    ctx["week"] = week_from_iso(*inferred)

    # Nel rapporto, sempre: e' la settimana su cui verranno ritagliati turni e
    # slot, e va vista anche quando le fonti WFM mancano.
    report.week = ctx.get("week_inferred")
    report.week_bounds = ctx.get("week")

    report.coherence = _run_coherence(contract, settings, report, blocks, ctx)
    return report, blocks


def _colonna(dataset, block, canonical: str) -> list:
    """I valori di una colonna del blocco, per nome canonico.

    Serve ai passaggi che lavorano sui dati gia' letti (lo storico AHT, i case
    type) e che non devono sapere in che colonna del foglio quel campo finisce.
    """
    from ..core.contract import col_to_index

    fld = next((f for f in dataset.input_fields if f.canonical == canonical), None)
    if fld is None:
        return []
    off = fld.target_index - col_to_index(dataset.data_start_col)
    return [row[off] if off < len(row) else None for row in block.rows]


def _terne_sf(contract: Contract, blocks: dict) -> list[tuple]:
    """(canale, case type, AHT) riga per riga da SF_DATABASE.

    E' l'unico ingrediente dello storico settimanale: si prende dal blocco gia'
    validato dal preflight, non rileggendo il CSV. Cosi' lo storico e il foglio
    `SF_DATABASE` del report non possono raccontare due cose diverse.
    """
    block = blocks.get("SF_DATABASE")
    if not block or not block.rows:
        return []
    ds = contract.dataset("SF_DATABASE")
    canali = _colonna(ds, block, "Case Origin (group)")
    tipi = _colonna(ds, block, "Case Type")
    aht = _colonna(ds, block, "Case AHT (mins)")
    if not (canali and tipi):
        return []
    return list(zip(canali, tipi, aht or [None] * len(canali)))


def _at_start_times(dataset, block) -> list:
    """I valori di `AT_DATASET!Start Time`, per ricavare anno e copertura."""
    from ..core.contract import col_to_index

    fld = next((f for f in dataset.input_fields if f.canonical == "Start Time"), None)
    if fld is None:
        return []
    off = fld.target_index - col_to_index(dataset.data_start_col)
    return [row[off] for row in block.rows if off < len(row) and row[off] is not None]


def _resolve_week(start_times: list, offset_hours: float):
    """La settimana ISO su cui ritagliare `Turni` e `Slot Only Cases`.

    Si ricava **dai dati**: la settimana in cui `AT_DATASET` sta per la gran
    parte. Il numero passato a `--week` non la determina, la controlla — se i
    due non coincidono, il preflight blocca (vedi il controllo "settimana
    dichiarata").

    Le sorgenti WFM contengono molti piu' giorni e molte piu' persone del
    necessario (il roster del W30 copre due mesi e 130 agenti): il ritaglio alla
    settimana di AT_DATASET e' quello che le riduce a cio' che serve.

    NON si usa il min/max delle date, ed e' misurato: l'export di AT e' per data
    Seattle e `Data Milano` = `INT(Start Time + offset/24)` lo sposta avanti. Nel
    W30 le date Milano vanno dal 20 al **27** luglio — otto giorni, a cavallo di
    due settimane ISO — mentre il workbook copre i sette dal 20 al 26. Col
    min/max si caricherebbe un giorno in piu' di turni e slot.
    """
    from ..core.wfmsource import infer_iso_week

    return infer_iso_week(start_times, offset_hours)


def _run_coherence(contract, settings, report, blocks, ctx):
    from ..core.coherence import check_sources

    turni = blocks.get("Turni")
    slot = blocks.get("Slot Only Cases")
    # Gira SEMPRE, anche senza le fonti WFM. Ogni singolo controllo e' gia'
    # guardato dai propri dati, quindi quelli che non hanno di che lavorare si
    # saltano da se'. Uscire prima faceva perdere il confronto fra la settimana
    # dichiarata a --week e quella che dicono i dati: con `--week 30` su un
    # export della W31 il preflight rispondeva OK, e sarebbe uscito un
    # `Omni_Report_W30.xlsm` pieno di W31. Un report con l'etichetta sbagliata e'
    # peggio di un report che manca, perche' viene archiviato.
    email = None
    if settings.template.is_file():
        email = _read_email_agenti(settings.template)

    sf = blocks.get("SF_DATABASE")
    date_viewpoint = (
        _colonna(contract.dataset("SF_DATABASE"), sf, "Date Viewpoint") if sf else None
    )

    # Il perimetro dei duplicati si ricava dalle date dei casi chiusi, non dalle
    # righe di preambolo dell'export: quelle dicono l'intervallo RICHIESTO al
    # report, queste quello OTTENUTO.
    dup = blocks.get("DUP_DATASET")
    dup_closed = (
        _colonna(contract.dataset("DUP_DATASET"), dup, "Date/Time Closed")
        if dup and dup.rows else None
    )

    return check_sources(
        dup_closed=dup_closed,
        dup_capienze=_dup_capienze(settings),
        dup_conteggi=_dup_conteggi(contract, blocks),
        column_stats={n: b.stats for n, b in blocks.items() if b.stats},
        date_viewpoint=date_viewpoint,
        casetype_nuovi=_casetype_nuovi(contract, settings, blocks),
        casetype_esclusi=settings.casetype_esclusi,
        roster_notes=ctx.get("roster_notes"),
        backoffice_notes=ctx.get("backoffice_notes"),
        turni_rows=turni.rows if turni else None,
        slot_rows=slot.rows if slot else None,
        at_dates=ctx.get("week"),
        week=ctx.get("week"),
        at_start_times=ctx.get("at_start_times"),
        timezone_offset_hours=ctx.get("offset_hours", 0.0),
        week_declared=ctx.get("week_declared"),
        week_inferred=ctx.get("week_inferred"),
        case_owners=_case_owners(contract, blocks),
        email_agenti=email,
        contratti=settings.contratti,
        wanted_skills=settings.sources.skills,
        include_marked=settings.sources.include_marked_skills,
        aliases_available=bool(ctx.get("aliases")),
        aliases=ctx.get("aliases"),
        row_limits=_row_limits(contract, settings),
        last_rows=_last_rows(contract, blocks),
    )


def _casetype_nuovi(contract: Contract, settings: Settings, blocks: dict):
    """I case type nei dati che 'Helper CaseType' del template non elenca.

    Si legge il template OFFLINE, senza Excel: cosi' l'avviso arriva col
    preflight — prima del build — invece che dopo. Chi guarda il rapporto sa
    subito che nel deepdive mancheranno quei tipi, e puo' decidere se
    aggiungerli alla lista curata prima di generare.
    """
    if not settings.template.is_file():
        return None
    terne = _terne_sf(contract, blocks)
    if not terne:
        return None
    from ..core.casetype import coppie_dai_dati, da_appendere

    esistenti = _read_casetype_helper(settings.template)
    if esistenti is None:
        return None
    return da_appendere(esistenti, coppie_dai_dati((c, t) for c, t, _ in terne)) or None


def _read_casetype_helper(template) -> list[tuple[str, str]] | None:
    """Le coppie (canale, case type) elencate in 'Helper CaseType'!A:B.

    `None` (e non lista vuota) se il foglio non c'e': significa "non so", ed e'
    diverso da "la lista e' vuota". Con un template senza quel foglio il
    controllo si salta, non segnala 53 case type mancanti.
    """
    try:
        from ..core.xlsxsource import read_sheet

        sheet = read_sheet(template, "Helper CaseType")
    except PipelineError:
        return None
    out: list[tuple[str, str]] = []
    for rownum in sheet.rows:
        if rownum == 1:
            continue
        canale, ct = sheet.cell("A", rownum), sheet.cell("B", rownum)
        if canale and ct and str(canale).strip() and str(ct).strip():
            out.append((str(canale).strip(), str(ct).strip()))
    return out


def _assert_vba_compilabile(template) -> None:
    """Ferma il build prima di aprire Excel se il modulo non compila.

    Il modo peggiore in cui questo puo' andare male e' quello che e' andato male:
    Excel compila il VBA solo quando serve, quindi un modulo con le dichiarazioni
    nell'ordine sbagliato si salva senza un lamento, e l'errore esce quando la
    pipeline lancia la macro — come dialogo modale, che in automazione nessuno
    chiude. Il build resta appeso a tempo indeterminato con Excel aperto.

    Meglio due secondi di lettura del binario e un errore che dice cosa fare.
    Senza oletools non si puo' leggere il modulo: si passa, senza inventare un
    allarme che non si e' in grado di verificare.
    """
    from .vbapatch import DEFAULT_MODULE, check_declaration_order, read_module

    try:
        code = read_module(template, DEFAULT_MODULE)
    except PipelineError:
        return

    problemi = check_declaration_order(code)
    if problemi:
        raise PipelineError(
            f"Il modulo {DEFAULT_MODULE} del template non compila: "
            f"{len(problemi)} dichiarazioni di modulo stanno dopo una procedura.\n"
            f"  VBA lo rifiuta con 'dopo End Sub ... sono ammessi solo commenti',\n"
            f"  e lo fa con un dialogo che in automazione nessuno puo' chiudere.\n"
            + "".join(f"    {p}\n" for p in problemi[:5])
            + f"  Vanno spostate SOPRA la prima Sub del modulo. Il modulo corretto\n"
            f"  e' in template/CreaMalpractice_patched_da_incollare.vb: aprilo,\n"
            f"  copia tutto e incollalo sul modulo (Alt+F11, Ctrl+A, incolla).\n"
            f"  Poi:  omni-report check"
        )


def _row_limits(contract: Contract, settings: Settings) -> list | None:
    """I limiti di riga scritti nelle formule del template.

    Due scansioni, non una. Quella per intervalli
    (`SUMIFS(AT_DATASET!$P$2:$P$130000, ...)`) copre tutti i dataset; quella a
    cella singola solo quelli che dichiarano `read_by_row`, perche' li' il limite
    non e' un intervallo ma il riferimento puntuale piu' alto — `Duplicates Helper`
    legge `DUP_DATASET` cella per cella, e senza la seconda scansione quel dataset
    risulterebbe senza limite noto.

    Senza template non si puo' sapere, e non si inventa: si restituisce None e il
    controllo si salta.
    """
    if not settings.template.is_file():
        return None
    from ..core.templatescan import scan_cell_refs, scan_row_limits

    try:
        limiti = scan_row_limits(settings.template, tuple(contract.datasets))
        per_riga = tuple(
            n for n, d in contract.datasets.items() if d.read_by_row
        )
        if per_riga:
            limiti += scan_cell_refs(settings.template, per_riga)
        return limiti
    except PipelineError:
        return None


def _dup_capienze(settings: Settings) -> dict[str, int] | None:
    """La capienza degli elenchi dei fogli DC, letta dal template.

    Si misura sull'ultima riga con una formula, non su una costante nel codice: se
    le formule vengono tirate piu' in basso, il controllo lo segue da solo.
    """
    if not settings.template.is_file():
        return None
    from ..core.duplicati import SCAFFALI, SPILL_ORIGIN, SPILL_ORIGIN_CAPIENZA, punti_da_misurare
    from ..core.templatescan import scan_formula_extent

    try:
        estensioni = scan_formula_extent(settings.template, punti_da_misurare())
    except PipelineError:
        return None
    out = {
        s.etichetta: estensioni[(s.sheet, s.col)] - s.prima_riga + 1
        for s in SCAFFALI
        if (s.sheet, s.col) in estensioni
    }
    # Lo spill del menu Origin non e' una formula per riga: la sua capienza e' lo
    # spazio fra dove parte e cio' che lo blocca. Vive nel modulo, con la nota.
    out[SPILL_ORIGIN.etichetta] = SPILL_ORIGIN_CAPIENZA
    return out or None


def _dup_conteggi(contract: Contract, blocks: dict) -> dict[str, int] | None:
    """Quanti valori distinti la settimana chiede a ciascun elenco dei fogli DC."""
    block = blocks.get("DUP_DATASET")
    if not block or not block.rows:
        return None
    from ..core.contract import col_to_index
    from ..core.duplicati import conteggi

    ds = contract.dataset("DUP_DATASET")
    start = col_to_index(ds.data_start_col)
    offset = {f.canonical: f.target_index - start for f in ds.input_fields}
    return conteggi(block.rows, offset) or None


def _last_rows(contract: Contract, blocks: dict) -> dict[str, int]:
    """Ultima riga del FOGLIO che ciascun dataset occupera'.

    Non il numero di righe: la riga di Excel, intestazione compresa — e' quello
    che i limiti nelle formule esprimono.
    """
    out: dict[str, int] = {}
    for nome, block in blocks.items():
        if not block.rows:
            continue
        out[nome] = contract.dataset(nome).header_row + len(block.rows)
    return out


def _case_owners(contract: Contract, blocks: dict) -> dict[str, list]:
    """I nomi degli agenti che compaiono nelle fonti caso-per-caso.

    Sono le persone che nella settimana hanno lavorato casi. Confrontarle con
    'Email Agenti' e' un controllo che si puo' fare con i soli quattro dataset,
    senza roster ne' back office — cioe' subito, appena arrivano gli export.
    """
    from ..core.contract import col_to_index

    # DUP_DATASET e' la terza fonte caso-per-caso: chi ha lavorato duplicati e
    # non e' in 'Email Agenti' esce dallo stesso controllo, gratis.
    fonti = {
        "SF_DATABASE": "Employee Name",
        "PSAT_DATASET": "Agent Name",
        "DUP_DATASET": "Full Name",
    }
    out: dict[str, list] = {}
    for nome_ds, campo in fonti.items():
        block = blocks.get(nome_ds)
        if not block:
            continue
        ds = contract.dataset(nome_ds)
        fld = next((f for f in ds.input_fields if f.canonical == campo), None)
        if fld is None:
            continue
        off = fld.target_index - col_to_index(ds.data_start_col)
        out[nome_ds] = [
            r[off] for r in block.rows if off < len(r) and r[off]
        ]
    return out


def _timezone_offset(contract: Contract, settings: Settings) -> float:
    """L'offset Seattle->Milano, letto dal template dove vive.

    Sta in `Helper Malpractice`!B7 (= 9) e alimenta le formule `Data Milano` /
    `Ora Milano`. Duplicarlo in config vorrebbe dire due verita' che possono
    divergere.
    """
    ref = next(
        (ds.derive_offset_from for ds in contract.datasets.values() if ds.derive_offset_from),
        None,
    )
    if not ref or not settings.template.is_file():
        return 0.0
    sheet_name, _, cell = ref.rpartition("!")
    sheet_name = sheet_name.strip().strip("'")
    try:
        from ..core.xlsxsource import read_sheet

        sheet = read_sheet(settings.template, sheet_name)
        col = "".join(c for c in cell if c.isalpha())
        row = int("".join(c for c in cell if c.isdigit()))
        return float(sheet.cell(col, row))
    except Exception:
        return 0.0


def _read_email_agenti(template) -> set[str] | None:
    """Nomi presenti in `Email Agenti` del template, normalizzati."""
    try:
        from ..core.names import normalize_name
        from ..core.xlsxsource import read_sheet

        sheet = read_sheet(template, "Email Agenti")
    except PipelineError:
        return None
    out = set()
    for rownum in sheet.rows:
        if rownum == 1:
            continue
        nome = sheet.cell("A", rownum)
        if nome and str(nome).strip():
            out.add(normalize_name(nome))
    return out or None


def _bloccato_da_excel(path: Path) -> bool:
    """Il file `~$<nome>` che Excel crea quando un workbook e' aperto.

    Non e' infallibile — un Excel schiantato puo' lasciare il lock orfano — ma
    e' il segnale standard, e non richiede di aprire il file per scoprirlo.
    Senza questo controllo, scrivere sopra un file che un collega ha ancora
    aperto produce un salvataggio a meta' o un errore COM poco chiaro, invece
    di un messaggio che dice semplicemente "chiudilo".
    """
    return path.with_name(f"~${path.name}").exists()


def _tipi_com_transitori() -> tuple[type[BaseException], ...]:
    """I tipi di eccezione COM da ritentare, se esistono su questa macchina.

    Senza pywin32 (qui in sandbox, o su un sistema che non e' Windows) non
    esistono errori COM: la tupla resta vuota, e nessun retry scatta mai —
    semplicemente perche' non c'e' nulla di quel genere da ritentare.
    """
    try:
        import pywintypes

        return (pywintypes.com_error,)
    except ImportError:
        return ()


_MARCATORI_TRANSITORI = (
    "rpc_e_call_rejected",
    "call was rejected by callee",
    "server execution failed",
    "chiamata rifiutata dal chiamato",
)


def _e_transitorio(exc: BaseException) -> bool:
    if isinstance(exc, _tipi_com_transitori()):
        return True
    testo = str(exc).lower()
    return any(m in testo for m in _MARCATORI_TRANSITORI)


def _scrivi_con_copia_atomica(sorgente: Path, out_path: Path, scrivi) -> None:
    """Esegue `scrivi(percorso_temporaneo)` su una COPIA, e la promuove al nome
    buono solo se `scrivi` non solleva.

    `out_path` non viene mai toccato finche' il lavoro non e' riuscito per
    intero: se si interrompe a meta' — un errore COM, la macro che fallisce —
    resta quello di prima (intatto per un giro parziale, assente per un giro
    completo), non un file a meta' scrittura che un collega potrebbe prendere
    per completo perche' si chiama giusto.

    Separata da `build()` apposta: qui non c'e' Excel, quindi tutto cio' che
    serve xlwings non e' verificabile in questo ambiente — ma questa funzione
    non ne ha bisogno, e puo' essere testata per davvero con un `scrivi`
    finto.
    """
    temp_path = out_path.with_name(f"{out_path.stem}.building{out_path.suffix}")
    try:
        shutil.copy2(sorgente, temp_path)
        scrivi(temp_path)
        # Atomico sullo stesso filesystem, Windows compreso: non esiste un
        # istante in cui out_path e' a meta' scritto.
        os.replace(temp_path, out_path)
    except Exception:
        temp_path.unlink(missing_ok=True)
        raise


def _con_retry(fn, *, tentativi: int = 3, attesa_iniziale: float = 2.0):
    """Ritenta una chiamata xlwings che ha fallito per un errore COM transitorio.

    Excel a volte rifiuta una chiamata RPC per un istante — e' "occupato": un
    ridisegno, un altro comando ancora in corso — e la stessa chiamata,
    ripetuta poco dopo, funziona. Non e' un errore da mostrare all'utente: e'
    rumore di fondo di COM. Qualunque altro errore (un file mancante, un
    foglio che non c'e') non e' transitorio e deve fallire subito, non essere
    ritentato — altrimenti si perde solo tempo prima di fallire comunque.
    """
    attesa = attesa_iniziale
    for tentativo in range(1, tentativi + 1):
        try:
            return fn()
        except Exception as exc:
            if tentativo == tentativi or not _e_transitorio(exc):
                raise
            time.sleep(attesa)
            attesa *= 2


def _settimana_storico(contract: Contract, blocks: dict, report: PreflightReport):
    """(anno ISO, settimana) sotto cui archiviare gli aggregati di questa settimana.

    Si preferisce `Date Viewpoint` dell'export SF: gli aggregati sono di QUEL
    file, quindi la settimana giusta e' quella che il file dichiara. Il
    controllo di coerenza garantisce che coincida con quella del resto del
    report (e BLOCCA se non coincide), percio' preferire l'una o l'altra non
    cambia il risultato — cambia solo cosa succede quando una delle due manca.

    Se `Date Viewpoint` non c'e' (export vecchio, o colonna rinominata), si
    ripiega sulla settimana ricavata da `AT_DATASET`.
    """
    sf = blocks.get("SF_DATABASE")
    if sf and sf.rows:
        from datetime import date as _date
        from datetime import datetime as _dt

        giorni: dict[_date, int] = {}
        for v in _colonna(contract.dataset("SF_DATABASE"), sf, "Date Viewpoint"):
            if isinstance(v, _dt):
                v = v.date()
            if isinstance(v, _date):
                giorni[v] = giorni.get(v, 0) + 1
        if giorni:
            prevalente = max(giorni.items(), key=lambda kv: kv[1])[0]
            iso = prevalente.isocalendar()
            return (iso[0], iso[1])
    return report.week


def _scrivi_derivati(book, contract: Contract, settings: Settings, blocks: dict, settimana):
    """Storico AHT e case type nuovi: le due scritture che non vengono da un file.

    Se non si sa a che settimana appartengono i dati non si scrive niente e si
    dice perche': mettere gli aggregati sotto la settimana sbagliata
    corromperebbe lo storico in modo permanente, ed e' l'unica cosa del progetto
    che non si ricostruisce rilanciando il programma.

    Nota su cosa succede se il build fallisce DOPO questo punto (la macro va in
    errore, per esempio): il CSV e' gia' aggiornato mentre il workbook non viene
    promosso al nome buono. E' voluto e innocuo — gli aggregati sono di dati che
    il preflight ha validato, quindi restano corretti, e `unisci` e' idempotente:
    il giro successivo riscrive le stesse righe. Rimandare la scrittura a build
    riuscito vorrebbe dire tenere aperto il file dello storico attraverso tutta
    la sequenza Excel, per proteggersi da un caso che non produce danni.
    """
    from ..core.aht_history import aggrega, carica, righe_foglio, scrivi, unisci
    from ..core.casetype import coppie_dai_dati, da_appendere
    from .writer import (
        FOGLIO_STORICO,
        WriteResult,
        append_helper_casetype,
        leggi_coppie_helper,
        write_aht_history,
    )

    terne = _terne_sf(contract, blocks)
    if not terne:
        return []
    if not settimana:
        return [WriteResult(
            dataset=FOGLIO_STORICO,
            rows_written=0,
            rows_cleared=0,
            range_written="(saltato)",
            warnings=(
                "Non so a quale settimana attribuire gli aggregati AHT: manca sia "
                "'Date Viewpoint' in SF_DATABASE sia la settimana ricavata da "
                "AT_DATASET. Lo storico NON e' stato toccato — e' meglio di "
                "archiviarlo sotto la settimana sbagliata.",
            ),
        )]

    iso_year, week = settimana
    nuove = aggrega(
        terne, iso_year=iso_year, week=week, esclusi=settings.casetype_esclusi
    )
    storico = unisci(carica(settings.aht_history), nuove)
    scrivi(settings.aht_history, storico)

    out = [write_aht_history(book, righe_foglio(storico))]
    out.append(append_helper_casetype(
        book,
        da_appendere(
            leggi_coppie_helper(book),
            coppie_dai_dati((c, t) for c, t, _ in terne),
        ),
    ))
    return out


def build(
    contract: Contract,
    settings: Settings,
    week: str,
    *,
    week_number: int | None = None,
    only: set[str] | None = None,
) -> BuildResult:
    """Il "pulsante": da input/ a output/Omni_Report_W{week}.xlsm."""
    report, blocks = run_preflight(
        contract, settings, week_number=week_number, only=only
    )
    preflight_path = report.write(settings.preflight_path(week))

    if not report.ok:
        return BuildResult(workbook=None, preflight=preflight_path, report=report)

    out_path = settings.workbook_path(week)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    if _bloccato_da_excel(out_path):
        raise PipelineError(
            f"{out_path.name} e' aperto in Excel (trovato il file di lock "
            f"'~${out_path.name}').\n"
            f"  Chiudilo e rilancia: scrivere sopra un file aperto produce un "
            f"salvataggio a meta' o un errore poco chiaro, non un avviso."
        )

    # `--only` ricarica un SOTTOINSIEME: deve scrivere nel workbook che c'e'
    # gia', non ripartire dal template. Ripartendo dal template i fogli non
    # ricaricati resterebbero pieni dei dati della settimana del template — un
    # report mezzo W31 e mezzo W30, senza che nulla lo dica. Meglio pretendere
    # che il giro completo sia stato fatto almeno una volta.
    parziale = bool(only) and only != set(contract.datasets)
    if parziale:
        if not out_path.is_file():
            raise PipelineError(
                f"--only ricarica alcuni fogli di un workbook che esiste gia', ma "
                f"{out_path.name} non c'e'.\n"
                f"  Fai prima il giro completo:\n"
                f"    omni-report build --week {week}\n"
                f"  Ripartire dal template scrivendo solo "
                f"{', '.join(sorted(only))} lascerebbe gli altri fogli con i dati "
                f"della settimana del template."
            )
    else:
        if not settings.template.is_file():
            raise PipelineError(
                f"Template non trovato: {settings.template}\n"
                f"  Ricavalo dal workbook di una settimana chiusa svuotando i 4 fogli "
                f"DATASET e aggiungendo il flag SilentMode al VBA.\n"
                f"  Istruzioni: docs/architettura.md §6."
            )
        _assert_vba_compilabile(settings.template)

    try:
        import xlwings as xw
    except ImportError:
        raise PipelineError(
            "xlwings non e' installato: serve per scrivere e ricalcolare.\n"
            "  pip install -e '.[excel]'\n"
            "  Richiede Excel desktop (Windows o macOS): il ricalcolo delle formule\n"
            "  ad array 365 e del VBA non e' replicabile altrove (piano §10)."
        ) from None

    writes: list[WriteResult] = []
    macro_ran = False
    settimana_storico = _settimana_storico(contract, blocks, report)

    def _scrivi_workbook(temp_path: Path) -> None:
        nonlocal macro_ran
        app = None
        book = None
        try:
            app = _con_retry(lambda: xw.App(visible=settings.excel.visible, add_book=False))
            app.display_alerts = False
            app.screen_updating = False
            book = _con_retry(lambda: app.books.open(str(temp_path)))

            offset = _read_offset(book, contract) if settings.derived_mode == "python" else None

            # Solo i dataset per cui c'e' un blocco: con `--only` gli altri non
            # sono stati letti, e cercarli qui era un KeyError proprio nel caso
            # in cui `--only` serve (i turni che cambiano a giro iniziato).
            for name, dataset in contract.datasets.items():
                block = blocks.get(name)
                if block is None:
                    continue
                if settings.derived_mode == "python" and dataset.derived_fields:
                    block = add_derived(dataset, block, offset or 0.0)
                writes.append(write_block(book, contract, dataset, block))

            # I fogli derivati, DOPO che SF_DATABASE e' stato scritto (e la
            # tabella AHT_Data ridimensionata dal suo write_block) e PRIMA del
            # ricalcolo: 'AHT Trend WoW' e 'Helper CaseType' sono tutte formule,
            # e devono ricalcolare su questi dati, non su quelli di prima.
            writes.extend(
                _scrivi_derivati(book, contract, settings, blocks, settimana_storico)
            )

            if settings.excel.full_rebuild:
                app.api.CalculateFullRebuild()
            else:
                app.calculate()

            _run_macro(book, settings)
            macro_ran = True

            _con_retry(book.save)
        finally:
            if book is not None:
                book.close()
            if app is not None:
                app.quit()

    _scrivi_con_copia_atomica(
        out_path if parziale else settings.template, out_path, _scrivi_workbook
    )

    # Il workbook e' salvato: ora si guarda com'e' venuto. E' la controparte del
    # preflight — quello controlla cio' che entra, questo cio' che e' uscito, e
    # vede una famiglia di guasti che sulle sorgenti non e' visibile (un array
    # dinamico senza spazio per espandersi, una formula che legge uno spill
    # rimasto vuoto).
    from ..core.templatescan import scan_error_cells

    try:
        errori = scan_error_cells(out_path)
    except PipelineError:
        errori = []

    return BuildResult(
        workbook=out_path,
        preflight=preflight_path,
        report=report,
        writes=writes,
        macro_ran=macro_ran,
        error_cells=errori,
    )


def _read_offset(book, contract: Contract) -> float:
    """Legge l'offset fuso dal workbook, non da settings.yml.

    Sta in 'Helper Malpractice'!B7 (= 9, ore Seattle->Milano). Duplicarlo nella
    config vorrebbe dire due verita' che possono divergere.
    """
    ref = None
    for ds in contract.datasets.values():
        if ds.derive_offset_from:
            ref = ds.derive_offset_from
            break
    if not ref:
        return 0.0
    sheet, _, cell = ref.rpartition("!")
    sheet = sheet.strip().strip("'")
    try:
        value = book.sheets[sheet].range(cell).value
    except Exception:
        raise PipelineError(
            f"Impossibile leggere l'offset fuso da {ref} nel template."
        ) from None
    if value is None:
        raise PipelineError(f"{ref} e' vuota: l'offset fuso orario e' obbligatorio.")
    return float(value)


def _run_macro(book, settings: Settings) -> None:
    """Lancia la macro malpractice, silenziando i MsgBox.

    Il modulo CreaMalpractice ha due MsgBox (fine esecuzione ed errore): in
    automazione bloccano il processo a tempo indeterminato. Il template va
    patchato una volta con un flag pubblico — piano §12, dettagli in
    docs/architettura.md §6. Se il flag non c'e' si prosegue, ma con un errore
    parlante invece di un blocco muto.
    """
    flag = settings.excel.silent_mode_flag
    try:
        book.macro(f"CreaMalpractice.Set{flag}")(True)
    except Exception:
        try:
            book.app.api.Run(f"'{book.name}'!Set{flag}", True)
        except Exception:
            raise PipelineError(
                f"Il template non espone Set{flag}: la macro "
                f"{settings.excel.macro!r} mostrerebbe due MsgBox e bloccherebbe "
                f"il run a tempo indeterminato.\n"
                f"  Patcha il VBA come descritto in docs/architettura.md §6, "
                f"oppure lancia con excel.visible: true e chiudi i dialoghi a mano."
            ) from None

    try:
        book.macro(settings.excel.macro)()
    except Exception as exc:
        raise PipelineError(
            f"La macro {settings.excel.macro!r} e' andata in errore: {exc}\n"
            f"  Il workbook non e' stato salvato. Le macro sono abilitate?"
        ) from None
