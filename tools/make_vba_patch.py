#!/usr/bin/env python3
"""Genera il modulo VBA gia' patchato, da incollare o importare nell'editor.

Non modifica il workbook: produce due file di testo. Per applicare la patch
**senza aprire l'editor** c'e' invece:

    omni-report patch-template

che passa dall'API di VBA via xlwings (serve la spunta "Considera attendibile
l'accesso al modello a oggetti dei progetti VBA"). Questo strumento e' la via
che non richiede nessuna impostazione.

    pip install -e ".[audit]"
    python tools/make_vba_patch.py template/Omni_Report_TEMPLATE.xlsm

Poi, nell'editor VBA (Alt+F11): apri il modulo CreaMalpractice, Ctrl+A, e
incolla il contenuto del file `_da_incollare.vb`. Salva mantenendo il `.xlsm`.

Idempotente: su un modulo gia' patchato non fa nulla.

La logica sta in `fasterreports.omni.vbapatch`, condivisa con `patch-template`
e con i controlli di `omni-report check`: una regola sola, in un posto solo.
"""

from __future__ import annotations

import argparse
import difflib
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from fasterreports.omni.vbapatch import (  # noqa: E402
    DEFAULT_FLAG,
    DEFAULT_MODULE,
    PatchError,
    patch,
    read_module,
    verify,
    write_encoded,
    write_module_files,
)

# Ri-esportati per i test e per chi importa questo modulo direttamente.
_write = write_encoded
__all__ = ["patch", "verify", "read_module", "write_encoded", "_write", "PatchError"]


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("workbook", type=Path)
    ap.add_argument("--module", default=DEFAULT_MODULE)
    ap.add_argument("--flag", default=DEFAULT_FLAG)
    ap.add_argument("-o", "--output", type=Path, help="default: accanto al workbook")
    ap.add_argument("--diff", action="store_true", help="mostra il diff con contesto")
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
        print(f"  {'!' if d.startswith('ATTENZIONE') else '·'} {d}")

    diff = list(difflib.unified_diff(
        original.replace("\r\n", "\n").splitlines(),
        patched.splitlines(),
        fromfile=f"{args.module} (attuale)",
        tofile=f"{args.module} (patchato)",
        lineterm="",
        n=3 if args.diff else 0,
    ))
    if diff:
        print("\nDiff:")
        for line in diff:
            print(f"  {line}")
    else:
        print("\nNessuna modifica: il modulo e' gia' a posto.")

    problems = verify(patched, args.flag)
    if problems:
        print("\nIl risultato NON passa i controlli:", file=sys.stderr)
        for p in problems:
            print(f"  - {p}", file=sys.stderr)
        print("\nNon scrivo i file: andrebbe sistemato a mano.", file=sys.stderr)
        return 1

    base = args.output or args.workbook.with_name(f"{args.module}_patched")
    try:
        written = write_module_files(base, patched)
    except PatchError as exc:
        print(f"\nERRORE: {exc}\n", file=sys.stderr)
        return 1

    paste_path, paste_enc = written["paste"]
    import_path, import_enc = written["import"]
    n_special = sum(1 for ch in patched if ord(ch) > 127)

    print("\nScritti due file, uno per ciascun modo di applicarlo:")
    print(f"  {paste_path.name}")
    print(f"      da INCOLLARE (senza la riga Attribute) · {paste_enc}")
    print(f"  {import_path.name}")
    print(f"      da IMPORTARE (modulo completo) · {import_enc}")
    if n_special:
        print(
            f"\n  Nota: il modulo contiene {n_special} caratteri non-ASCII ORIGINALI\n"
            f"  (il grado nell'etichetta AHT e gli accenti dentro NormKey). Per questo\n"
            f"  i due file hanno codifiche diverse: letti male, NormKey smetterebbe di\n"
            f"  normalizzare gli accenti in silenzio. Dopo aver incollato, controlla\n"
            f"  che in NormKey si leggano ancora le vocali accentate."
        )
    print(
        f"\nVia consigliata — incollare (Alt+F11 nell'editor VBA):\n"
        f"  1. apri il modulo {args.module}\n"
        f"  2. Ctrl+A per selezionare tutto\n"
        f"  3. incolla il contenuto di {paste_path.name}\n"
        f"  4. salva MANTENENDO il formato .xlsm\n"
        f"  5. verifica con:  omni-report check\n"
        f"\nVia alternativa — importare: click destro sul modulo -> Remove (alla\n"
        f"domanda 'esportare prima?' rispondi No), poi File -> Import File e\n"
        f"scegli {import_path.name}.\n"
        f"\nSenza toccare l'editor:  omni-report patch-template"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
