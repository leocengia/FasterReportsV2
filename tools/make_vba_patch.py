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
# Solo ASCII, come il resto del modulo (l'autore scrive "verita'" invece di
# "verita" accentata): il file passa dalla clipboard e da editor di testo, e un
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

    base = args.output or args.workbook.with_name(f"{args.module}_patched")
    base = base.with_suffix("")

    # Due file, perche' i due modi di applicarlo vogliono contenuti diversi.
    #
    # `Attribute VB_Name = "..."` e' valido solo in un file esportato: incollarlo
    # nel corpo di un modulo dall'editor da' un errore di compilazione. Chi fa
    # Ctrl+A e incolla — cioe' quasi tutti — ci sbatte contro.
    for_import = base.with_suffix(".bas")
    enc_import = _write(for_import, patched, for_import_route=True)

    for_paste = base.with_name(base.name + "_da_incollare").with_suffix(".vb")
    body = "\n".join(
        l for l in patched.splitlines() if not l.strip().startswith("Attribute VB_")
    ).lstrip("\n")
    enc_paste = _write(for_paste, body, for_import_route=False)

    n_special = sum(1 for ch in patched if ord(ch) > 127)
    print(f"\nScritti due file, uno per ciascun modo di applicarlo:")
    print(f"  {for_paste.name}")
    print(f"      da INCOLLARE (senza la riga Attribute) · {enc_paste}")
    print(f"  {for_import.name}")
    print(f"      da IMPORTARE (modulo completo) · {enc_import}")
    if n_special:
        print(
            f"\n  Nota: il modulo contiene {n_special} caratteri non-ASCII "
            f"ORIGINALI (il grado\n"
            f"  nell'etichetta AHT e gli accenti dentro NormKey). Per questo i due\n"
            f"  file hanno codifiche diverse: letti male, NormKey smetterebbe di\n"
            f"  normalizzare gli accenti in silenzio. Dopo aver incollato,\n"
            f"  controlla che in NormKey si leggano ancora le vocali accentate."
        )
    print(
        f"\nVia consigliata — incollare (Alt+F11 nell'editor VBA):\n"
        f"  1. apri il modulo {args.module}\n"
        f"  2. Ctrl+A per selezionare tutto\n"
        f"  3. incolla il contenuto di {for_paste.name}\n"
        f"  4. salva MANTENENDO il formato .xlsm\n"
        f"  5. verifica con:  omni-report check\n"
        f"\nVia alternativa — importare: click destro sul modulo -> Remove (alla\n"
        f"domanda 'esportare prima?' rispondi No), poi File -> Import File e\n"
        f"scegli {for_import.name}.\n"
        f"\nIncollare e' piu' sicuro: non si rimuove niente, e se qualcosa va\n"
        f"storto basta annullare."
    )
    return 0


def _write(path: Path, text: str, *, for_import_route: bool) -> str:
    """Scrive il modulo nella codifica che il suo destinatario si aspetta.

    Non e' pignoleria. Il modulo contiene caratteri non-ASCII **originali**:
    il grado in `"AHT alto (>" & ... & "° pct)"` e, cosa piu' delicata, gli
    accenti dentro `NormKey`:

        s = Replace(s, "a-grave", "a"): s = Replace(s, "e-grave", "e")

    Se quel file viene letto con la codifica sbagliata, quei caratteri si
    corrompono e **NormKey smette di normalizzare gli accenti senza dire
    niente**: le chiavi non combaciano piu' e le ore previste vanno a zero.
    E' il tipo di guasto silenzioso che questo progetto esiste per togliere.

    Quindi:
      - `.bas` da importare -> cp1252 (ANSI), che e' quello che l'editor VBA
        produce esportando e si aspetta importando;
      - `.vb` da incollare  -> UTF-8 **con BOM**, cosi' gli editor di Windows
        la riconoscono invece di indovinare ANSI.

    Se un carattere non e' rappresentabile in cp1252 si solleva: meglio fermarsi
    che scrivere un file corrotto.
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


if __name__ == "__main__":
    raise SystemExit(main())
