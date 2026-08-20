#!/usr/bin/env python3
"""Le due patch al template per la sezione Duplicate Cases, senza aprire Excel.

    python tools/patch_template_duplicates.py template/Omni_Report_TEMPLATE.xlsm
    python tools/patch_template_duplicates.py FILE --dry-run

Perche' offline e non a mano in Excel. Sono due interventi che si fanno **una
volta**, e a mano il primo vuol dire riscrivere due celle e trascinarle fino a
riga 1001 su un foglio da 13 120 formule: fattibile, ma sono anche due minuti in
cui si puo' sbagliare una riga e non accorgersene. Qui invece la modifica e'
esattamente descrivibile, ripetibile, e verificabile senza Excel.

Perche' NON openpyxl: riscriverebbe tutto il workbook, e su questo file
significa perdere le formule ad array dinamico e i valori in cache — cioe'
corrompere il motore (vedi la testa di omni/writer.py). Qui si tocca il testo di
alcune formule dentro l'XML e si lascia identico tutto il resto, byte per byte.

--- PATCH 1: 'Duplicates Helper'!M/N smettono di parsare il testo ------------

Erano:  =IF($K2="","",DATE(VALUE(MID(K2,FIND("/",K2,FIND("/",K2)+1)+1,4)), ... ))
Ora:    =IF($K2="","",$K2)

Il contratto legge `DUP_DATASET!H` e `I` come `datetime` con `date_format`
dichiarato, quindi la pipeline scrive DATE VERE. Con la formula vecchia,
`FIND("/")` su una data da' `#VALUE!` — e con lei cadono `Opened`, `Closed`,
`Day`, `Weekday`, `Open hour`, `TTC (hours)` e `TTC bucket`, cioe' meta' dei
fogli DC.

Scriverle come testo NON era un'alternativa: Excel in locale italiano converte
`"8/4/2026 3:59 PM"` in data al momento della scrittura, e la formula vecchia si
romperebbe comunque — solo in modo imprevedibile invece che sempre.

Le colonne M e N sono memorizzate come UN master esplicito (`M2`, `N2`) piu' 16
gruppi di formule CONDIVISE (`<f t="shared" ref="M3:M66" si="0">…`). Le celle
eredi non hanno testo: cambiare il master cambia tutto il gruppo. Quindi le
sostituzioni sono 34, non 2000.

I valori in cache NON si toccano, ed e' voluto: il tipo del risultato non cambia
(un seriale di data resta un seriale, `""` resta `""`), e la pipeline fa
`CalculateFullRebuild` prima di salvare. Cancellarli sarebbe piu' lavoro e piu'
rischio per zero guadagno.

--- PATCH 2: via il collegamento esterno ------------------------------------

Il pacchetto contiene un `externalLink` verso

    L:\\Admin\\Individuali\\Leonardo\\Higiene Reports\\Duplicates\\Duplicates Report.xlsx

cioe' il workbook da cui i fogli nuovi sono stati copiati. **Nessuna formula lo
usa** (zero riferimenti `[1]` in tutto il file): e' un residuo. Ma sopravvive nel
pacchetto, e chi apre il report da un PC che non vede l'unita' `L:` si prende il
dialogo «Questa cartella di lavoro contiene collegamenti a origini dati
esterne». In automazione non blocca (`display_alerts = False`), a mano si'.

Togliere un external link vuol dire togliere quattro cose insieme, e se ne
manca una il file diventa illeggibile:

  1. le due parti  xl/externalLinks/externalLink1.xml  e il suo .rels
  2. la <Relationship> in xl/_rels/workbook.xml.rels
  3. l'<Override> in [Content_Types].xml
  4. il blocco <externalReferences> in xl/workbook.xml

Lo strumento e' **idempotente**: rilanciarlo su un template gia' patchato non
cambia niente e lo dice.
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
import zipfile
from pathlib import Path

FOGLIO = "Duplicates Helper"

# Colonna da riscrivere -> colonna da cui prende il valore.
# M ('Opened') viene da K ('Opened (text)'), N ('Closed') da L.
SORGENTE = {"M": "K", "N": "L"}

# Come si riconosce la formula VECCHIA: il parsing a mano del testo della data.
# Si pretende `DATE(VALUE(MID(` e `FIND("/"` per non toccare per sbaglio una
# formula gia' patchata, o una diversa da quella attesa.
#
# Le virgolette NON sono escapate: nel testo di una formula xlsx si escapano solo
# `<`, `>` e `&`, quindi qui `"` e' un `"` letterale. (Cercare `&quot;` fa
# rifiutare tutte e 34 le formule, che e' come e' andata la prima volta — e per
# fortuna, perche' significa che la guardia funziona: quando non riconosce, non
# tocca.)
_VECCHIA = re.compile(
    r'^IF\(\$(?P<src>[KL])(?P<row>\d+)="","",DATE\(VALUE\(MID\(.*FIND\("/".*\)$'
)


def nuova_formula(col: str, row: int) -> str:
    """La formula che sostituisce il parsing: prende il valore cosi' com'e'.

    `$K2` e non `K2`: e' la forma che hanno le altre colonne del foglio (`O`,
    `P`, `Q` usano `$M3`), e la colonna assoluta rende innocuo un trascinamento
    laterale.
    """
    src = SORGENTE[col]
    return f'IF(${src}{row}="","",${src}{row})'


def patch_duplicates_helper(raw: str) -> tuple[str, list[str], list[str]]:
    """Riscrive M/N nel foglio. Restituisce (xml, fatte, non riconosciute)."""
    fatte: list[str] = []
    saltate: list[str] = []

    # Solo le celle delle colonne M e N che portano il TESTO di una formula:
    # `<c r="M2" ...><f ...>TESTO</f>`. Le eredi (`<f t="shared" si="0"/>`) non
    # hanno testo e non matchano — ed e' esattamente il punto: ereditano.
    pat = re.compile(
        r'(?P<pre><c r="(?P<cella>[MN]\d+)"(?:[^>"]|"[^"]*")*?>)'
        r"(?P<apertura><f(?:[^>\"]|\"[^\"]*\")*?>)(?P<f>[^<]*)</f>",
        re.S,
    )

    out = []
    pos = 0
    for m in pat.finditer(raw):
        out.append(raw[pos : m.start()])
        testo = m.group("f")
        cella = m.group("cella")
        apertura = m.group("apertura")
        col = re.match(r"([A-Z]+)", cella).group(1)
        rif = re.search(r'ref="([A-Z]+)(\d+):', apertura)
        row = int(rif.group(2)) if rif else int(re.search(r"(\d+)", cella).group(1))

        vecchia = _VECCHIA.match(testo)
        if not vecchia:
            saltate.append(f"{cella}: formula inattesa, non toccata ({testo[:40]}…)")
            out.append(m.group(0))
        elif vecchia.group("src") != SORGENTE[col] or int(vecchia.group("row")) != row:
            saltate.append(
                f"{cella}: legge {vecchia.group('src')}{vecchia.group('row')}, "
                f"atteso {SORGENTE[col]}{row}"
            )
            out.append(m.group(0))
        else:
            fatte.append(f"{col}{row}")
            out.append(
                m.group("pre") + apertura + nuova_formula(col, row) + "</f>"
            )
        pos = m.end()
    out.append(raw[pos:])
    return "".join(out), fatte, saltate


def togli_external_link(parti: dict[str, bytes]) -> list[str]:
    """Rimuove il collegamento esterno, tutte e quattro le sue tracce."""
    tolte: list[str] = []

    presenti = [n for n in parti if n.startswith("xl/externalLinks/")]
    if not presenti:
        return tolte
    for n in presenti:
        del parti[n]
        tolte.append(n)

    rels = parti["xl/_rels/workbook.xml.rels"].decode("utf8")
    ids = re.findall(
        r'<Relationship Id="(rId\d+)"[^>]*/relationships/externalLink"[^>]*/>', rels
    )
    rels_nuovo = re.sub(
        r'<Relationship Id="rId\d+"[^>]*/relationships/externalLink"[^>]*/>', "", rels
    )
    parti["xl/_rels/workbook.xml.rels"] = rels_nuovo.encode("utf8")
    tolte.append(f"workbook.xml.rels: {', '.join(ids) or 'nessuna relazione'}")

    ct = parti["[Content_Types].xml"].decode("utf8")
    parti["[Content_Types].xml"] = re.sub(
        r"<Override PartName=\"/xl/externalLinks/[^\"]+\"[^>]*/>", "", ct
    ).encode("utf8")
    tolte.append("[Content_Types].xml: Override della parte")

    wb = parti["xl/workbook.xml"].decode("utf8")
    wb_nuovo = re.sub(r"<externalReferences>.*?</externalReferences>", "", wb, flags=re.S)
    if wb_nuovo != wb:
        tolte.append("workbook.xml: blocco <externalReferences>")
    parti["xl/workbook.xml"] = wb_nuovo.encode("utf8")
    return tolte


def _foglio_target(parti: dict[str, bytes], nome: str) -> str | None:
    wb = parti["xl/workbook.xml"].decode("utf8", "replace")
    rels = parti["xl/_rels/workbook.xml.rels"].decode("utf8", "replace")
    relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    for m in re.finditer(r"<sheet ([^>]*)/>", wb):
        n = re.search(r'name="([^"]*)"', m.group(1))
        rid = re.search(r'r:id="(rId\d+)"', m.group(1))
        if n and rid and n.group(1) == nome:
            t = relmap.get(rid.group(1), "").lstrip("/")
            return t if t.startswith("xl/") else "xl/" + t
    return None


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("workbook", type=Path)
    ap.add_argument("--dry-run", action="store_true", help="dice cosa farebbe, non scrive")
    ap.add_argument(
        "--no-backup", action="store_true",
        help="non salva il .prima-della-patch (di norma lo si vuole)",
    )
    args = ap.parse_args(argv)

    if not args.workbook.is_file():
        print(f"File non trovato: {args.workbook}", file=sys.stderr)
        return 2

    with zipfile.ZipFile(args.workbook) as z:
        ordine = z.namelist()
        parti = {n: z.read(n) for n in ordine}

    target = _foglio_target(parti, FOGLIO)
    if target is None:
        print(f"Il workbook non contiene il foglio {FOGLIO!r}.", file=sys.stderr)
        return 1

    raw = parti[target].decode("utf8")
    nuovo, fatte, saltate = patch_duplicates_helper(raw)
    parti[target] = nuovo.encode("utf8")

    print(f"PATCH 1 — {FOGLIO}!M/N (parsing del testo -> valore diretto)")
    if fatte:
        print(f"  {len(fatte)} formule riscritte (master espliciti + gruppi condivisi):")
        print("    " + ", ".join(fatte))
    else:
        print("  niente da fare: nessuna formula col parsing del testo.")
    for s in saltate:
        print(f"  ! {s}")

    tolte = togli_external_link(parti)
    print("\nPATCH 2 — collegamento esterno")
    if tolte:
        for t in tolte:
            print(f"  tolto: {t}")
    else:
        print("  niente da fare: nessun collegamento esterno.")

    if not fatte and not tolte:
        print("\nIl template e' gia' patchato. Nessuna scrittura.")
        return 0

    if args.dry_run:
        print("\n--dry-run: non ho scritto niente.")
        return 0

    if not args.no_backup:
        backup = args.workbook.with_name(
            f"{args.workbook.stem}_prima_delle_patch_duplicates{args.workbook.suffix}"
        )
        shutil.copy2(args.workbook, backup)
        print(f"\nCopia di sicurezza: {backup.name}")

    # Si riscrive il pacchetto nell'ORDINE originale delle parti. Excel non lo
    # pretende, ma un diff fra i due zip resta leggibile — e su un file da 8 MB
    # poter dire "e' cambiato solo questo" vale il piccolo sforzo.
    tmp = args.workbook.with_suffix(args.workbook.suffix + ".patching")
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for n in ordine:
                if n in parti:
                    z.writestr(n, parti[n])
        tmp.replace(args.workbook)
    except Exception:
        tmp.unlink(missing_ok=True)
        raise
    print(f"Scritto: {args.workbook}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
