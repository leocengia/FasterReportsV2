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
from .settings import Settings
from .writer import WriteResult, write_block


@dataclass
class BuildResult:
    workbook: Path | None
    preflight: Path
    report: PreflightReport
    writes: list[WriteResult] = field(default_factory=list)
    macro_ran: bool = False

    @property
    def ok(self) -> bool:
        return self.workbook is not None and self.report.ok


def _open_source(contract: Contract, settings: Settings, dataset, ctx: dict):
    """Apre la sorgente giusta per il dataset, secondo `reader` del contratto.

    I 4 CSV passano da `csvsource`; `Turni` e `Slot Only Cases` dagli adattatori
    WFM, che riducono una matrice larga a righe tidy. A valle non cambia nulla:
    entrambi restituiscono `(headers, rows)`.
    """
    path = settings.input_path(dataset.name)
    if dataset.reader == "csv":
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


def run_preflight(
    contract: Contract,
    settings: Settings,
    *,
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

            report.datasets.append(
                DatasetReport(name=name, source=str(settings.input_path(name)), error=str(exc))
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

            # La settimana si ricava da AT_DATASET, non si scrive due volte.
            if name == "AT_DATASET" and "week" not in ctx:
                ctx["week"] = _week_from_at(dataset, ds_report.mapping, block)

    report.coherence = _run_coherence(contract, settings, report, blocks, ctx)
    return report, blocks


def _week_from_at(dataset, mapping, block):
    """Intervallo di date coperto da `AT_DATASET!Start Time`."""
    from ..core.contract import col_to_index
    from ..core.wfmsource import week_from_dates

    fld = next((f for f in dataset.input_fields if f.canonical == "Start Time"), None)
    if fld is None:
        return None
    off = fld.target_index - col_to_index(dataset.data_start_col)
    return week_from_dates(row[off] for row in block.rows if off < len(row))


def _run_coherence(contract, settings, report, blocks, ctx):
    from ..core.coherence import check_sources

    turni = blocks.get("Turni")
    slot = blocks.get("Slot Only Cases")
    if not (ctx.get("roster_notes") or turni or slot):
        return None

    email = None
    if settings.template.is_file():
        email = _read_email_agenti(settings.template)

    return check_sources(
        roster_notes=ctx.get("roster_notes"),
        backoffice_notes=ctx.get("backoffice_notes"),
        turni_rows=turni.rows if turni else None,
        slot_rows=slot.rows if slot else None,
        at_dates=ctx.get("week"),
        email_agenti=email,
        contratti=settings.contratti,
        wanted_skills=settings.sources.skills,
        include_marked=settings.sources.include_marked_skills,
        aliases_available=bool(ctx.get("aliases")),
    )


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


def build(contract: Contract, settings: Settings, week: str) -> BuildResult:
    """Il "pulsante": da input/ a output/Omni_Report_W{week}.xlsm."""
    report, blocks = run_preflight(contract, settings)
    preflight_path = report.write(settings.preflight_path(week))

    if not report.ok:
        return BuildResult(workbook=None, preflight=preflight_path, report=report)

    if not settings.template.is_file():
        raise PipelineError(
            f"Template non trovato: {settings.template}\n"
            f"  Ricavalo dal workbook di una settimana chiusa svuotando i 4 fogli "
            f"DATASET e aggiungendo il flag SilentMode al VBA.\n"
            f"  Istruzioni: docs/architettura.md §6."
        )

    out_path = settings.workbook_path(week)
    out_path.parent.mkdir(parents=True, exist_ok=True)
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

        for name, dataset in contract.datasets.items():
            block = blocks[name]
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

    return BuildResult(
        workbook=out_path,
        preflight=preflight_path,
        report=report,
        writes=writes,
        macro_ran=macro_ran,
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
