#!/usr/bin/env python3
"""Genera il modulo VBA gia' patchato, da incollare nell'editor.

Perche' non si patcha direttamente il file: il codice sta in
`xl/vbaProject.bin`, un contenitore OLE dove i moduli sono compressi, e accanto
al sorgente c'e' il **p-code compilato**. Excel, quando le versioni combaciano,
esegue il p-code e non il sorgente: riscrivendo solo il testo si otterrebbe un
file che sembra patchato e continua a eseguire il codice vecchio. Modo di
fallire silenzioso, cioe' il tipo di guasto che questo progetto elimina.

Quindi: si legge il VBA, si applicano le tre modifiche al testo, e si scrive un
`.bas` che tu incolli nell'editor. Due minuti, zero rischio, e il risultato e'
verificabile con `omni-report check`.

    pip install -e ".[audit]"
    python tools/make_vba_patch.py template/Omni_Report_TEMPLATE.xlsm

Poi, nell'editor VBA (Alt+F11):
  - apri il modulo CreaMalpractice
  - seleziona tutto (Ctrl+A) e incolla il contenuto del .bas generato
  - salva mantenendo il formato .xlsm

L'operazione e' idempotente: su un modulo gia' patchato non fa nulla.
"""

from __future__ import annotations

import argparse
import difflib
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

DEFAULT_MODULE = "CreaMalpractice"
DEFAULT_FLAG = "SilentMode"

# La dichiarazione va dopo Option Explicit, prima di qualunque altra cosa.
_DECL = """
' ==== Modalita' silenziosa, per l'esecuzione automatica ====
' In automazione un MsgBox blocca il processo a tempo indeterminato, in attesa
' di un clic che nessuno dara' — e con Excel invisibile il dialogo non si vede
' nemmeno. La pipeline chiama Set{flag}(True) prima della macro.
Public {flag} As Boolean

Public Sub Set{flag}(ByVal value As Boolean)
    {flag} = value
End Sub
"""


class PatchError(Exception):
    pass


def read_module(path: Path, module: str) -> str:
    try:
        from oletools.olevba import VBA_Parser
    except ImportError:
        raise PatchError(
            "oletools non installato: serve per leggere il VBA.\n"
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
    # olevba usa sia 'CreaMalpractice' sia 'CreaMalpractice.bas'
    for key in (module, f"{module}.bas"):
        if key in found:
            return found[key]
    raise PatchError(
        f"{path.name}: modulo {module!r} non trovato.\n"
        f"  Moduli presenti: {', '.join(sorted(found))}\n"
        f"  Se il modulo si chiama diversamente, passa --module."
    )


def patch(code: str, flag: str = DEFAULT_FLAG) -> tuple[str, list[str]]:
    """Applica le tre modifiche. Ritorna (codice, elenco di cosa e' cambiato)."""
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
            # Nessun Option Explicit: si mette dopo l'ultimo Attribute.
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
            # Es. `x = MsgBox(...)`: non e' un avviso, e' una domanda. Non si
            # tocca, ma va segnalato: in automazione bloccherebbe comunque.
            done.append(f"ATTENZIONE riga {i+1}: MsgBox non in prima posizione, "
                        f"non modificato -> {stripped[:60]}")
            continue
        indent = line[: len(line) - len(line.lstrip())]
        lines[i] = f"{indent}If Not {flag} Then {stripped}"
        wrapped += 1
    if wrapped:
        done.append(f"{wrapped} MsgBox messi dietro `If Not {flag} Then`")

    # --- 3. propagazione dell'errore in CleanFail --------------------------
    raise_line = f"If {flag} Then Err.Raise Err.Number, , Err.Description"
    if raise_line.split(" Then ")[1] in code:
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
            # Dopo l'ultima riga del blocco prima di End Sub.
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
    """Gli stessi controlli che fa `omni-report check`."""
    problems = []
    if f"Public Sub Set{flag}" not in code:
        problems.append(f"Set{flag} assente")
    nudi = [
        l.strip() for l in code.splitlines()
        if "MsgBox" in l and flag not in l and not l.strip().startswith("'")
    ]
    if nudi:
        problems.append(f"{len(nudi)} MsgBox non protetti: {nudi[0][:60]}")
    if "Err.Raise" not in code:
        problems.append("Err.Raise assente: l'errore non verrebbe propagato")
    return problems


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("workbook", type=Path)
    ap.add_argument("--module", default=DEFAULT_MODULE)
    ap.add_argument("--flag", default=DEFAULT_FLAG)
    ap.add_argument("-o", "--output", type=Path, help="default: accanto al workbook")
    ap.add_argument("--diff", action="store_true", help="mostra il diff completo")
    args = ap.parse_args(argv)

    if not args.workbook.is_file():
        print(f"File non trovato: {args.workbook}", file=sys.stderr)
        return 2

    try:
        original = read_module(args.workbook, args.module)
    except PatchError as exc:
        print(f"\nERRORE: {exc}\n", file=sys.stderr)
        return 1

    patched, done = patch(original, args.flag)

    print(f"Modulo {args.module!r} letto da {args.workbook.name} "
          f"({len(original.splitlines())} righe)")
    print("\nModifiche:")
    for d in done:
        mark = "  !" if d.startswith("ATTENZIONE") else "  ·"
        print(f"{mark} {d}")

    diff = list(difflib.unified_diff(
        original.replace("\r\n", "\n").splitlines(),
        patched.splitlines(),
        fromfile=f"{args.module} (attuale)",
        tofile=f"{args.module} (patchato)",
        lineterm="",
        n=0 if not args.diff else 3,
    ))
    if diff:
        print("\nDiff:")
        for line in diff:
            print(f"  {line}")
    else:
        print("\nNessuna modifica: il modulo e' gia' a posto.")

    problems = verify(patched, args.flag)
    if problems:
        print("\nIl risultato NON passa i controlli:")
        for p in problems:
            print(f"  - {p}")
        print("\nNon scrivo il file: andrebbe sistemato a mano.")
        return 1

    out = args.output or args.workbook.with_name(f"{args.module}_patched.bas")
    # CRLF: e' quello che l'editor VBA si aspetta.
    out.write_text(patched.replace("\n", "\r\n"), encoding="utf-8")
    print(f"\nScritto: {out}")
    print(
        "\nCome applicarlo (Alt+F11 nell'editor VBA):\n"
        f"  1. apri il modulo {args.module}\n"
        "  2. Ctrl+A, poi incolla il contenuto del file generato\n"
        "  3. salva MANTENENDO il formato .xlsm\n"
        "  4. verifica con:  omni-report check\n"
        "\nIn alternativa: click destro sul modulo -> Remove, poi File -> Import\n"
        "File. Incollare sopra e' piu' sicuro: se qualcosa va storto il modulo\n"
        "originale e' ancora li'."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
