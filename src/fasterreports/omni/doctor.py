"""`omni-report check`: l'ambiente e il template sono pronti?

Esiste per una ragione precisa: senza questi controlli, un template preparato
male fallisce durante il `build` — a Excel già aperto, magari appeso su un
`MsgBox` invisibile che nessuno può chiudere. Meglio saperlo prima, in due
secondi, con l'indicazione di cosa sistemare.

Quasi tutti i controlli girano **senza Excel**: leggono il template come archivio
zip. Solo l'ultimo lo apre davvero, e solo se xlwings è installato.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from ..core.names import normalize_name
from ..core.xlsxsource import read_sheet, read_sheet_names

OK = "OK"
MANCA = "MANCA"
ATTENZIONE = "ATTENZIONE"
SALTATO = "SALTATO"


@dataclass
class Check:
    name: str
    status: str
    detail: str = ""
    hint: str = ""

    @property
    def blocking(self) -> bool:
        return self.status == MANCA


@dataclass
class CheckReport:
    checks: list[Check] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not any(c.blocking for c in self.checks)

    def add(self, name, status, detail="", hint="") -> None:
        self.checks.append(Check(name, status, detail, hint))

    def render(self) -> str:
        width = max((len(c.name) for c in self.checks), default=10)
        lines = ["", "Controllo di ambiente e template", "=" * 60]
        for c in self.checks:
            lines.append(f"  [{c.status:^10}] {c.name.ljust(width)}  {c.detail}")
            if c.hint and c.status in (MANCA, ATTENZIONE):
                for line in c.hint.splitlines():
                    lines.append(f"  {' ' * 12} -> {line}")
        lines.append("=" * 60)
        if self.ok:
            lines.append("Tutto pronto per il build.")
        else:
            n = sum(1 for c in self.checks if c.blocking)
            lines.append(
                f"{n} controlli da sistemare prima del build.\n"
                f"Istruzioni passo per passo: template/COME_PREPARARE_IL_TEMPLATE.md"
            )
        return "\n".join(lines) + "\n"


# ---------------------------------------------------------------------------

def run_checks(contract, settings, *, try_excel: bool = True) -> CheckReport:
    rep = CheckReport()
    _check_python_deps(rep)
    tpl_ok = _check_template_file(rep, settings)
    if tpl_ok:
        _check_template_sheets(rep, settings, contract)
        _check_vba(rep, settings)
        _check_alias_table(rep, settings)
        _check_email_agenti(rep, settings)
    _check_input_dir(rep, settings, contract)
    if try_excel:
        _check_excel(rep)
    else:
        rep.add("Excel raggiungibile", SALTATO, "richiesto --no-excel")
    return rep


def _check_python_deps(rep: CheckReport) -> None:
    try:
        import xlwings  # noqa: F401
    except ImportError:
        rep.add(
            "xlwings installato", MANCA, "",
            "pip install -e \".[excel]\"\n"
            "Serve per scrivere e ricalcolare: il ricalcolo delle formule ad\n"
            "array 365 e del VBA non e' replicabile senza Excel vero.",
        )
        return
    import xlwings

    rep.add("xlwings installato", OK, f"versione {xlwings.__version__}")


def _check_template_file(rep: CheckReport, settings) -> bool:
    path = Path(settings.template)
    if not path.is_file():
        rep.add(
            "template presente", MANCA, str(path),
            "Copia un workbook di una settimana chiusa e rinominalo cosi'.\n"
            "Non serve svuotare i DATASET: il build li pulisce da se'.\n"
            "Istruzioni: template/COME_PREPARARE_IL_TEMPLATE.md",
        )
        return False

    if path.suffix.lower() != ".xlsm":
        rep.add(
            "template con macro", MANCA, f"estensione {path.suffix}",
            "Il motore include il VBA: il file deve essere .xlsm.\n"
            "Se l'hai salvato come .xlsx hai perso tutte le macro.",
        )
        return False

    mb = path.stat().st_size / 1_048_576
    rep.add("template presente", OK, f"{path.name} ({mb:.1f} MB)")

    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
    except zipfile.BadZipFile:
        rep.add(
            "template leggibile", MANCA, "l'archivio non si apre",
            "Il file e' corrotto o non e' un vero .xlsm.",
        )
        return False

    if "xl/vbaProject.bin" not in names:
        rep.add(
            "macro nel template", MANCA, "nessun vbaProject.bin",
            "Il template non contiene VBA: senza la macro\n"
            "Refresh_Dettaglio_Malpractice il foglio 'Dettaglio Malpractice'\n"
            "resta vuoto. Probabilmente e' stato salvato come .xlsx e\n"
            "rinominato a mano.",
        )
        return False
    rep.add("macro nel template", OK, "vbaProject.bin presente")
    return True


def _check_template_sheets(rep: CheckReport, settings, contract) -> None:
    try:
        present = set(read_sheet_names(settings.template))
    except Exception as exc:
        rep.add("fogli del template", MANCA, str(exc))
        return

    wanted = {ds.sheet for ds in contract.datasets.values()}
    extra_needed = {"Anagrafica", "Email Agenti", "Helper Turni", "Helper Malpractice",
                    "Report Agenti", "Dettaglio Malpractice", "Malpractice Recap"}
    missing = sorted((wanted | extra_needed) - present)
    if missing:
        rep.add(
            "fogli attesi", MANCA, f"mancano: {', '.join(missing)}",
            "Il template non e' un Omni Report completo: la pipeline scriverebbe\n"
            "in fogli che non esistono. Riparti da un workbook di una settimana\n"
            "chiusa invece di costruirlo da zero.",
        )
    else:
        rep.add("fogli attesi", OK, f"{len(present)} fogli, tutti quelli richiesti")


def _check_vba(rep: CheckReport, settings) -> None:
    """`SetSilentMode` c'e'?

    Con oletools si legge il codice e si verifica anche che i MsgBox siano dietro
    il flag. Senza, si cerca il nome nel binario: i nomi delle Sub pubbliche
    stanno nello stream `dir` non compresso, quindi la ricerca e' affidabile per
    dire "c'e'" — meno per dire com'e' fatto il resto.
    """
    flag_sub = f"Set{settings.excel.silent_mode_flag}"

    code = None
    try:
        from oletools.olevba import VBA_Parser

        parser = VBA_Parser(str(settings.template))
        if parser.detect_vba_macros():
            code = "\n".join(c for _f, _s, _n, c in parser.extract_macros())
    except ImportError:
        pass
    except Exception:
        code = None

    if code is not None:
        if flag_sub not in code:
            rep.add(
                f"{flag_sub} nel VBA", MANCA, "non trovata nel codice",
                _vba_hint(settings),
            )
            return
        rep.add(f"{flag_sub} nel VBA", OK, "trovata")

        nudi = [
            line.strip()
            for line in code.splitlines()
            if "MsgBox" in line and "SilentMode" not in line and not line.strip().startswith("'")
        ]
        if nudi:
            rep.add(
                "MsgBox silenziati", ATTENZIONE,
                f"{len(nudi)} MsgBox non protetti dal flag",
                "In automazione bloccano il run in attesa di un clic che nessuno\n"
                "dara', e con Excel invisibile il dialogo non si vede nemmeno.\n"
                "Righe da sistemare:\n  " + "\n  ".join(nudi[:4]),
            )
        else:
            rep.add("MsgBox silenziati", OK, "tutti dietro il flag")

        if "Err.Raise" not in code:
            rep.add(
                "errore propagato", ATTENZIONE, "nessun Err.Raise nel modulo",
                "Silenziare il MsgBox d'errore senza propagarlo fa fallire la\n"
                "macro senza che la pipeline lo sappia: il workbook verrebbe\n"
                "salvato con un 'Dettaglio Malpractice' incompleto.\n"
                "Aggiungi in CleanFail:  If SilentMode Then Err.Raise Err.Number, , Err.Description",
            )
        else:
            rep.add("errore propagato", OK, "Err.Raise presente")
        return

    # Ripiego: ricerca nel binario.
    with zipfile.ZipFile(settings.template) as z:
        blob = z.read("xl/vbaProject.bin")
    if flag_sub.encode("latin-1") in blob or flag_sub.encode("utf-16-le") in blob:
        rep.add(
            f"{flag_sub} nel VBA", OK,
            "nome trovato nel binario (installa oletools per il controllo completo)",
        )
    else:
        rep.add(
            f"{flag_sub} nel VBA", MANCA,
            "nome non trovato nel binario",
            _vba_hint(settings) + "\n"
            "Nota: questo controllo cerca nel binario compresso. Per una verifica\n"
            "completa del codice:  pip install -e \".[audit]\"",
        )


def _vba_hint(settings) -> str:
    flag = settings.excel.silent_mode_flag
    return (
        f"Apri l'editor VBA (Alt+F11), modulo CreaMalpractice, e aggiungi:\n"
        f"    Public {flag} As Boolean\n"
        f"    Public Sub Set{flag}(ByVal value As Boolean)\n"
        f"        {flag} = value\n"
        f"    End Sub\n"
        f"poi metti i due MsgBox dietro `If Not {flag} Then ...`.\n"
        f"Passo per passo: template/COME_PREPARARE_IL_TEMPLATE.md"
    )


def _check_alias_table(rep: CheckReport, settings) -> None:
    try:
        from ..core.wfmsource import read_alias_map

        aliases = read_alias_map(settings.template)
    except Exception as exc:
        rep.add(
            "tabella alias nomi", ATTENZIONE, f"non leggibile ({exc})",
            "Vive in 'Helper Malpractice'!D:E, dov'e' anche il VBA che la usa.",
        )
        return
    if not aliases:
        rep.add(
            "tabella alias nomi", ATTENZIONE, "vuota",
            "Il file back-office scrive alcuni nomi diversamente dal roster\n"
            "('alessandro passierello' invece di 'passariello'). Senza la tabella\n"
            "quegli agenti risultano orfani e il preflight segnala un problema che\n"
            "non esiste.",
        )
    else:
        rep.add("tabella alias nomi", OK, f"{len(aliases)} alias in Helper Malpractice!D:E")


def _check_email_agenti(rep: CheckReport, settings) -> None:
    try:
        sheet = read_sheet(settings.template, "Email Agenti")
    except Exception as exc:
        rep.add("Email Agenti", MANCA, str(exc))
        return
    nomi = {
        normalize_name(sheet.cell("A", r))
        for r in sheet.rows
        if r > 1 and sheet.cell("A", r)
    }
    nomi.discard("")
    con_email = sum(
        1 for r in sheet.rows if r > 1 and sheet.cell("A", r) and sheet.cell("B", r)
    )
    if not nomi:
        rep.add(
            "Email Agenti", MANCA, "nessun nome",
            "Il VBA lavora per email: senza questa mappa NESSUNA regola di\n"
            "malpractice si applica a nessuno.",
        )
    elif con_email < len(nomi):
        rep.add(
            "Email Agenti", ATTENZIONE,
            f"{len(nomi)} nomi, {con_email} con email",
            "Gli agenti senza email vengono scartati da AddATRules in silenzio.",
        )
    else:
        rep.add("Email Agenti", OK, f"{len(nomi)} agenti con email")


def _check_input_dir(rep: CheckReport, settings, contract) -> None:
    missing = []
    found = []
    for name in contract.datasets:
        try:
            path = settings.input_path(name)
        except Exception:
            missing.append(f"{name} (nome file non configurato)")
            continue
        (found if path.is_file() else missing).append(path.name)
    if missing:
        rep.add(
            "fonti in input/", ATTENZIONE,
            f"{len(found)}/{len(contract.datasets)} presenti; mancano: {', '.join(missing)}",
            "Non blocca il check: le fonti si mettono al momento del run.\n"
            "I nomi attesi sono in config/settings.yml -> input_files.",
        )
    else:
        rep.add("fonti in input/", OK, f"tutte e {len(found)}")


def _check_excel(rep: CheckReport) -> None:
    try:
        import xlwings as xw
    except ImportError:
        rep.add("Excel raggiungibile", SALTATO, "xlwings non installato")
        return

    app = None
    try:
        app = xw.App(visible=False, add_book=False)
        version = app.version
        rep.add("Excel raggiungibile", OK, f"versione {version}")
    except Exception as exc:
        rep.add(
            "Excel raggiungibile", MANCA, f"{type(exc).__name__}: {exc}",
            "Serve Excel desktop (Windows o macOS) sulla macchina dove gira il\n"
            "pulsante: gli array dinamici e il VBA vanno valutati da Excel vero.\n"
            "Su macOS la prima volta va concesso il permesso di automazione.",
        )
    finally:
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass
