#!/usr/bin/env python3
"""Toglie il limite di riga dalle formule di 'Profilo Colonne SF'.

Il foglio e' un diagnostico: una riga per ciascuna delle 143 colonne di
SF_DATABASE, con quante celle sono popolate, quanti valori distinti e un giudizio
("VUOTA nell'export"). Serve a capire dove possano stare duplicati ed esclusioni.

Le sue formule leggevano un intervallo con la riga finale **scritta dentro**:

    D6 = COUNTIF(INDEX(SF_DATABASE!$A$2:$EM$3389, 0, $A6), "<>")

3389 era l'estensione dei dati del W30. Con 2944 righe funziona — le righe in piu'
sono vuote e non contano — ma il giorno in cui l'export supera le 3389 righe il
foglio conta **di meno**, e non lo dice: "% riempimento" e "valori distinti"
diventano numeri plausibili e sbagliati.

La correzione non alza il limite: lo **toglie**. `AHT_Data` e' il ListObject che
copre i dati di SF_DATABASE e viene ridimensionato a ogni build, quindi
`INDEX(AHT_Data, 0, n)` segue i dati da se', per sempre.

    D6 = COUNTIF(INDEX(AHT_Data, 0, $A6), "<>")

Perche' uno strumento e non l'XML a mano: le formule di colonna F sono array
dinamici con metadati (`cm="1"`), e riscriverle nel file a mano vuol dire
rischiare di rompere quei metadati. Passando da Excel le formule vengono
reinserite e ricalcolate come si deve — stesso ragionamento del VBA e del p-code.

Uso, sulla macchina con Excel:

    python tools/fix_profilo_colonne_sf.py template/Omni_Report_TEMPLATE.xlsm

Poi:  omni-report check   e   omni-report preflight --week NN
La segnalazione "formule vicine al limite per SF_DATABASE" deve sparire.
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

FOGLIO = "Profilo Colonne SF"
PRIMA_RIGA = 6
ULTIMA_RIGA = 148  # 143 colonne + 5 righe di intestazione

# `AHT_Data` senza specificatore = il corpo dati della tabella, intestazione
# esclusa: esattamente cio' che serve, e si ridimensiona da se'.
F_POPOLATE = '=COUNTIF(INDEX(AHT_Data,0,$A{r}),"<>")'
F_DISTINTI = (
    "=IFERROR(COUNTA(UNIQUE(FILTER(INDEX(AHT_Data,0,$A{r}),"
    'INDEX(AHT_Data,0,$A{r})<>""))),0)'
)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("workbook", type=Path, help="il template (o un workbook prodotto)")
    ap.add_argument("--visible", action="store_true")
    ap.add_argument("--no-backup", action="store_true")
    ap.add_argument(
        "--dry-run", action="store_true",
        help="mostra le formule che scriverebbe, senza aprire Excel",
    )
    args = ap.parse_args(argv)

    if args.dry_run:
        print(f"Foglio: {FOGLIO}, righe {PRIMA_RIGA}..{ULTIMA_RIGA}")
        for r in (PRIMA_RIGA, PRIMA_RIGA + 1, ULTIMA_RIGA):
            print(f"  D{r} = {F_POPOLATE.format(r=r)}")
            print(f"  F{r} = {F_DISTINTI.format(r=r)}")
        return 0

    if not args.workbook.is_file():
        print(f"File non trovato: {args.workbook}", file=sys.stderr)
        return 2

    try:
        import xlwings as xw
    except ImportError:
        print(
            "xlwings non installato: serve Excel per reinserire le formule.\n"
            '  pip install -e ".[excel]"\n'
            "\n"
            "Alternativa a mano, senza strumenti: apri il foglio "
            f"{FOGLIO!r}, scrivi in D{PRIMA_RIGA} e F{PRIMA_RIGA}\n"
            f"  {F_POPOLATE.format(r=PRIMA_RIGA)}\n"
            f"  {F_DISTINTI.format(r=PRIMA_RIGA)}\n"
            f"e trascinale in basso fino a riga {ULTIMA_RIGA}.",
            file=sys.stderr,
        )
        return 3

    if not args.no_backup:
        backup = args.workbook.with_name(
            args.workbook.stem + "_prima_dei_limiti" + args.workbook.suffix
        )
        shutil.copy2(args.workbook, backup)
        print(f"Copia di sicurezza: {backup.name}")

    app = None
    book = None
    try:
        app = xw.App(visible=args.visible, add_book=False)
        app.display_alerts = False
        book = app.books.open(str(args.workbook.resolve()))
        try:
            sht = book.sheets[FOGLIO]
        except Exception:
            print(
                f"Il foglio {FOGLIO!r} non c'e'. Fogli presenti: "
                f"{', '.join(s.name for s in book.sheets)}",
                file=sys.stderr,
            )
            return 4

        scritte = 0
        for r in range(PRIMA_RIGA, ULTIMA_RIGA + 1):
            # Formula2 e' la via degli array dinamici: senza, Excel entra in
            # modalita' CSE legacy e UNIQUE/FILTER non si espandono.
            sht.range(f"D{r}").api.Formula2 = F_POPOLATE.format(r=r)
            sht.range(f"F{r}").api.Formula2 = F_DISTINTI.format(r=r)
            scritte += 2
        print(f"{scritte} formule riscritte su {FOGLIO!r}.")

        app.api.CalculateFullRebuild()
        book.save()
        print("Salvato.")
    finally:
        if book is not None:
            book.close()
        if app is not None:
            app.quit()

    print(
        "\nVerifica:\n"
        "  omni-report preflight --week NN\n"
        "La riga 'formule vicine al limite per SF_DATABASE' deve essere sparita."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
