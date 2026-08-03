"""La patch della modalita' silenziosa al VBA: applicazione e verifica.

La macro `Refresh_Dettaglio_Malpractice` mostra due `MsgBox` — uno a fine
esecuzione, uno nel gestore d'errore. In automazione bloccano il processo a
tempo indeterminato, in attesa di un clic che nessuno dara', e con Excel
invisibile il dialogo non si vede nemmeno.

Tre modi di applicare la patch, in ordine di comodita':

1. **`omni-report patch-template`** — la pipeline apre il template con xlwings e
   riscrive il modulo passando dall'**API di VBA**. Nessun editor da aprire, e
   nessun problema di p-code: e' VBA stesso a ricompilare. Richiede la spunta
   *Considera attendibile l'accesso al modello a oggetti dei progetti VBA*.
2. **`tools/make_vba_patch.py`** — genera il modulo patchato come file, da
   incollare nell'editor. Non richiede nessuna impostazione.
3. A mano, seguendo `template/COME_PREPARARE_IL_TEMPLATE.md`.

Cosa NON si fa: riscrivere `xl/vbaProject.bin` a file chiuso. I moduli sono
compressi in un contenitore OLE e accanto al sorgente c'e' il **p-code
compilato**; quando le versioni combaciano Excel esegue quello, non il sorgente.
Si otterrebbe un file che *sembra* patchato e continua a eseguire il codice
vecchio: un guasto silenzioso, cioe' il tipo di cosa che questo progetto esiste
per eliminare.
"""

from __future__ import annotations

import re
from pathlib import Path

from ..core.errors import PipelineError

DEFAULT_MODULE = "CreaMalpractice"
DEFAULT_FLAG = "SilentMode"

# La dichiarazione va dopo Option Explicit, prima di qualunque altra cosa.
# Solo ASCII, come il resto del modulo (l'autore scrive "verita'" invece di
# "verita" accentata): il testo passa dalla clipboard e da editor di testo, e un
# carattere non-ASCII e' un problema di codifica in attesa di accadere.
_DECL = """
' ==== Modalita' silenziosa, per l'esecuzione automatica ====
' In automazione un MsgBox blocca il processo a tempo indeterminato, in attesa
' di un clic che nessuno dara', e con Excel invisibile il dialogo non si vede
' nemmeno. La pipeline chiama Set{flag}(True) prima di lanciare la macro.
Public {flag} As Boolean

Public Sub Set{flag}(ByVal value As Boolean)
    {flag} = value
End Sub

"""


class PatchError(PipelineError):
    """La patch non si puo' applicare, e si dice perche'."""


# ---------------------------------------------------------------------------
# La trasformazione del testo
# ---------------------------------------------------------------------------

def patch(code: str, flag: str = DEFAULT_FLAG) -> tuple[str, list[str]]:
    """Applica le tre modifiche. Ritorna (codice, elenco di cosa e' cambiato).

    Idempotente: su un modulo gia' patchato non fa nulla.
    """
    lines = code.replace("\r\n", "\n").split("\n")
    done: list[str] = []

    # --- 1. dichiarazione del flag ------------------------------------------
    if f"Public Sub Set{flag}" in code:
        done.append(f"Set{flag} gia' presente, non ridichiarata")
    else:
        idx = next(
            (i for i, l in enumerate(lines) if l.strip().startswith("Option Explicit")),
            None,
        )
        if idx is None:
            # Nessun Option Explicit: si mette dopo l'ultimo Attribute, non in
            # cima al file (le Attribute devono restare prime).
            idx = max(
                (i for i, l in enumerate(lines) if l.strip().startswith("Attribute ")),
                default=-1,
            )
        block = _DECL.format(flag=flag).rstrip("\n").split("\n")
        lines[idx + 1 : idx + 1] = block
        done.append(f"aggiunta la dichiarazione di {flag} e Set{flag}")

    # --- 2. MsgBox dietro il flag ------------------------------------------
    wrapped = 0
    for i, line in enumerate(lines):
        stripped = line.strip()
        if "MsgBox" not in stripped or stripped.startswith("'"):
            continue
        if flag in stripped:
            continue  # gia' protetto
        if not stripped.startswith("MsgBox"):
            # Es. `x = MsgBox(...)`: non e' un avviso, e' una domanda.
            # Riscriverla cambierebbe la semantica, quindi si segnala e si lascia.
            done.append(
                f"ATTENZIONE riga {i+1}: MsgBox non in prima posizione, "
                f"non modificato -> {stripped[:60]}"
            )
            continue
        indent = line[: len(line) - len(line.lstrip())]
        lines[i] = f"{indent}If Not {flag} Then {stripped}"
        wrapped += 1
    if wrapped:
        done.append(f"{wrapped} MsgBox messi dietro `If Not {flag} Then`")

    # --- 3. propagazione dell'errore in CleanFail --------------------------
    raise_line = f"If {flag} Then Err.Raise Err.Number, , Err.Description"
    if "Err.Raise" in code:
        done.append("Err.Raise gia' presente, non aggiunto")
    else:
        label = next(
            (i for i, l in enumerate(lines) if re.match(r"^\s*CleanFail\s*:", l)), None
        )
        if label is None:
            done.append(
                "ATTENZIONE nessuna etichetta CleanFail: l'errore NON viene "
                "propagato, va aggiunto a mano"
            )
        else:
            end = next(
                (i for i in range(label, len(lines)) if lines[i].strip() == "End Sub"),
                None,
            )
            if end is None:
                done.append("ATTENZIONE End Sub di CleanFail non trovato")
            else:
                indent = "    "
                for i in range(label + 1, end):
                    if lines[i].strip():
                        indent = lines[i][: len(lines[i]) - len(lines[i].lstrip())]
                lines[end:end] = [
                    f"{indent}' Senza propagare l'errore la macro fallirebbe in silenzio:",
                    f"{indent}' il workbook verrebbe salvato con un Dettaglio Malpractice",
                    f"{indent}' incompleto e nessuno se ne accorgerebbe.",
                    f"{indent}{raise_line}",
                ]
                done.append("aggiunta la propagazione dell'errore in CleanFail")

    return "\n".join(lines), done


def verify(code: str, flag: str = DEFAULT_FLAG) -> list[str]:
    """Gli stessi criteri di `omni/doctor.py::_check_vba`. Vuoto = a posto."""
    problems = []
    if f"Public Sub Set{flag}" not in code:
        problems.append(f"Set{flag} assente")
    nudi = [
        l.strip()
        for l in code.splitlines()
        if "MsgBox" in l and flag not in l and not l.strip().startswith("'")
    ]
    if nudi:
        problems.append(f"{len(nudi)} MsgBox non protetti: {nudi[0][:60]}")
    if "Err.Raise" not in code:
        problems.append("Err.Raise assente: l'errore non verrebbe propagato")
    return problems


# ---------------------------------------------------------------------------
# Lettura del modulo a file chiuso (oletools)
# ---------------------------------------------------------------------------

def read_module(path: str | Path, module: str = DEFAULT_MODULE) -> str:
    path = Path(path)
    try:
        from oletools.olevba import VBA_Parser
    except ImportError:
        raise PatchError(
            "oletools non installato: serve per leggere il VBA a file chiuso.\n"
            '  pip install -e ".[audit]"'
        ) from None

    parser = VBA_Parser(str(path))
    if not parser.detect_vba_macros():
        raise PatchError(
            f"{path.name}: nessuna macro nel file.\n"
            f"  Probabilmente e' stato salvato come .xlsx e rinominato a mano:\n"
            f"  in quel passaggio tutto il VBA e' andato perso."
        )
    found = {}
    for _fname, _stream, name, code in parser.extract_macros():
        found[name] = code
    for key in (module, f"{module}.bas"):
        if key in found:
            return found[key]
    raise PatchError(
        f"{path.name}: modulo {module!r} non trovato.\n"
        f"  Moduli presenti: {', '.join(sorted(found))}"
    )


# ---------------------------------------------------------------------------
# Scrittura dei file da incollare / importare
# ---------------------------------------------------------------------------

def write_module_files(base: Path, patched: str) -> dict[str, tuple[Path, str]]:
    """Due file, perche' i due modi di applicarlo vogliono contenuti diversi.

    `Attribute VB_Name = "..."` e' valido solo in un file esportato: incollarlo
    nel corpo di un modulo dall'editor da' un errore di compilazione. Chi fa
    Ctrl+A e incolla — cioe' quasi tutti — ci sbatterebbe contro.
    """
    base = Path(base).with_suffix("")
    out = {}

    for_import = base.with_suffix(".bas")
    out["import"] = (for_import, write_encoded(for_import, patched, for_import_route=True))

    body = "\n".join(
        l for l in patched.splitlines() if not l.strip().startswith("Attribute VB_")
    ).lstrip("\n")
    for_paste = base.with_name(base.name + "_da_incollare").with_suffix(".vb")
    out["paste"] = (for_paste, write_encoded(for_paste, body, for_import_route=False))
    return out


def write_encoded(path: Path, text: str, *, for_import_route: bool) -> str:
    """Scrive il modulo nella codifica che il suo destinatario si aspetta.

    Non e' pignoleria. Il modulo contiene caratteri non-ASCII **originali**: il
    grado in `"AHT alto (>" & ... & "° pct)"` e, cosa piu' delicata, gli accenti
    dentro `NormKey` (`Replace(s, "a-grave", "a")` e compagnia).

    Se quel file viene letto con la codifica sbagliata quei caratteri si
    corrompono e **NormKey smette di normalizzare gli accenti senza dire
    niente**: le chiavi non combaciano piu' e le ore previste vanno a zero.

    Quindi: `.bas` da importare in cp1252 (ANSI), che e' cio' che l'editor VBA
    produce esportando e si aspetta importando; `.vb` da incollare in UTF-8 con
    BOM, cosi' gli editor di Windows la riconoscono invece di indovinare ANSI.
    """
    data = text.replace("\r\n", "\n").replace("\n", "\r\n")
    if for_import_route:
        try:
            path.write_bytes(data.encode("cp1252"))
        except UnicodeEncodeError as exc:
            raise PatchError(
                f"{path.name}: il modulo contiene un carattere non rappresentabile "
                f"in ANSI/cp1252 ({exc.reason} a offset {exc.start}).\n"
                f"  L'import dell'editor VBA lo corromperebbe. Usa la via "
                f"'incolla' con il file .vb."
            ) from None
        return "cp1252 (ANSI, per l'import di Excel)"
    path.write_bytes(b"\xef\xbb\xbf" + data.encode("utf-8"))
    return "UTF-8 con BOM (per aprirlo e copiarlo senza sorprese)"


# ---------------------------------------------------------------------------
# Applicazione diretta, via xlwings
# ---------------------------------------------------------------------------

TRUST_HINT = (
    "Excel non consente l'accesso programmatico al progetto VBA.\n"
    "  Abilitalo una volta:\n"
    "    File > Opzioni > Centro protezione > Impostazioni Centro protezione\n"
    "      > Impostazioni macro > spunta\n"
    "        'Considera attendibile l'accesso al modello a oggetti dei progetti VBA'\n"
    "  Poi chiudi tutte le finestre di Excel e rilancia.\n"
    "\n"
    "  Se preferisci non abilitarlo, usa la via manuale:\n"
    "    python tools/make_vba_patch.py template/Omni_Report_TEMPLATE.xlsm\n"
    "  che genera il modulo da incollare nell'editor."
)


def apply_via_xlwings(
    template: Path,
    *,
    module: str = DEFAULT_MODULE,
    flag: str = DEFAULT_FLAG,
    backup: bool = True,
    visible: bool = False,
) -> tuple[list[str], Path | None]:
    """Patcha il modulo nel template, passando dall'API di VBA.

    Va per l'API e non per il binario: e' VBA stesso a ricompilare, quindi il
    p-code resta coerente col sorgente. Il rischio del "sembra patchato ma
    esegue il vecchio" non esiste.

    Ritorna (cosa e' cambiato, percorso del backup).
    """
    template = Path(template)
    if not template.is_file():
        raise PatchError(f"Template non trovato: {template}")

    try:
        import xlwings as xw
    except ImportError:
        raise PatchError(
            "xlwings non installato: serve per patchare il template.\n"
            '  pip install -e ".[excel]"'
        ) from None

    backup_path = None
    if backup:
        import shutil

        backup_path = template.with_name(template.stem + "_prima_della_patch.xlsm")
        shutil.copy2(template, backup_path)

    app = None
    book = None
    try:
        app = xw.App(visible=visible, add_book=False)
        app.display_alerts = False
        book = app.books.open(str(template))

        try:
            project = book.api.VBProject
            components = project.VBComponents
        except Exception as exc:
            raise PatchError(f"{TRUST_HINT}\n\n  Dettaglio: {exc}") from None

        try:
            component = components(module)
        except Exception:
            nomi = []
            try:
                nomi = [c.Name for c in components]
            except Exception:
                pass
            raise PatchError(
                f"Modulo {module!r} non trovato nel template.\n"
                f"  Moduli presenti: {', '.join(nomi) if nomi else '(non elencabili)'}"
            ) from None

        code_module = component.CodeModule
        n = code_module.CountOfLines
        original = code_module.Lines(1, n) if n else ""

        patched, done = patch(original, flag)
        problems = verify(patched, flag)
        if problems:
            raise PatchError(
                "Il risultato non passa i controlli, non lo scrivo:\n  - "
                + "\n  - ".join(problems)
            )

        if patched.replace("\r\n", "\n") == original.replace("\r\n", "\n"):
            return done, backup_path

        # Sostituzione completa: cancella e riscrive. Passando da AddFromString
        # e' VBA a fare il parsing, quindi un errore di sintassi si manifesta
        # qui e non a runtime.
        if n:
            code_module.DeleteLines(1, n)
        code_module.AddFromString(patched.replace("\r\n", "\n").replace("\n", "\r\n"))

        book.save()
        return done, backup_path
    finally:
        if book is not None:
            try:
                book.close()
            except Exception:
                pass
        if app is not None:
            try:
                app.quit()
            except Exception:
                pass
