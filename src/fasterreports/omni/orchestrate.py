"""Flusso end-to-end: CSV -> workbook finito.

STATO: la parte di preflight/ingestione gira e ha test. La parte Excel
(`build`) non e' mai stata eseguita: qui non c'e' Excel. Vedi
docs/architettura.md §8.

Ordine, e perche':
  1. preflight su TUTTI i CSV, prima di aprire Excel. Aprire Excel per poi
     scoprire che manca una colonna costa 30 secondi e lascia processi appesi.
  2. copia template -> output. Il template non viene mai aperto in scrittura:
     se un run va male, resta intatto.
  3. scrittura dei 4 dataset.
  4. ricalcolo completo. Le formule usano XLOOKUP/FILTER/UNIQUE in array e
     INDIRECT: volatili, un calculate() semplice non propaga sempre.
  5. macro malpractice, in modalita' silenziosa.
  6. salva e chiudi, sempre, anche in caso di errore.
"""

from __future__ import annotations

import shutil
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

    return check_sources(
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

    Senza template non si puo' sapere, e non si inventa: si restituisce None e il
    controllo si salta.
    """
    if not settings.template.is_file():
        return None
    from ..core.templatescan import scan_row_limits

    try:
        return scan_row_limits(settings.template, tuple(contract.datasets))
    except PipelineError:
        return None


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

    fonti = {"SF_DATABASE": "Employee Name", "PSAT_DATASET": "Agent Name"}
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
        shutil.copy2(settings.template, out_path)

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
    app = None
    book = None
    try:
        app = xw.App(visible=settings.excel.visible, add_book=False)
        app.display_alerts = False
        app.screen_updating = False
        book = app.books.open(str(out_path))

        offset = _read_offset(book, contract) if settings.derived_mode == "python" else None

        # Solo i dataset per cui c'e' un blocco: con `--only` gli altri non sono
        # stati letti, e cercarli qui era un KeyError proprio nel caso in cui
        # `--only` serve (i turni che cambiano a giro iniziato).
        for name, dataset in contract.datasets.items():
            block = blocks.get(name)
            if block is None:
                continue
            if settings.derived_mode == "python" and dataset.derived_fields:
                block = add_derived(dataset, block, offset or 0.0)
            writes.append(write_block(book, contract, dataset, block))

        if settings.excel.full_rebuild:
            app.api.CalculateFullRebuild()
        else:
            app.calculate()

        _run_macro(book, settings)
        macro_ran = True

        book.save()
    finally:
        if book is not None:
            book.close()
        if app is not None:
            app.quit()

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
