#!/usr/bin/env python3
"""Le due patch al template per la sezione Only Cases, senza aprire Excel.

    python tools/patch_template_only_cases.py template/Omni_Report_TEMPLATE.xlsm
    python tools/patch_template_only_cases.py FILE --dry-run

I due fogli `OC Eventi` e `Only Cases W33` sono arrivati nel template il
2026-08-21, disegnati e formattati a mano. Servono due correzioni che NON sono
questioni di aspetto, e per questo si fanno qui invece che in Excel: la
formattazione non viene toccata da nessuna delle due.

Perche' NON openpyxl: riscriverebbe tutto il workbook, e su questo file
significa perdere le formule ad array dinamico e i valori in cache — cioe'
corrompere il motore (vedi la testa di omni/writer.py). Qui si tocca il testo di
sei formule e tre stringhe di nome, e si lascia identico tutto il resto.

--- PATCH 1: il foglio non si chiama piu' come la settimana ------------------

Era:  `Only Cases W33`
Ora:  `Only Cases Dashboard`

Il nome portava dentro il numero di settimana. Excel se la cava — rinominando un
foglio aggiorna da se' le formule che lo citano — ma il **codice** no: il
controllo delle capienze (`core/onlycases.py`) cerca il foglio per nome, e il
lunedi' dopo un rename non lo troverebbe piu'. Un controllo che sparisce in
silenzio e' peggio di un controllo che non c'e', perche' il preflight continua a
dire OK.

La settimana resta scritta dove deve stare: nelle celle del foglio, che si
ricalcolano dai dati (`B6`/`B7` sono il primo e l'ultimo giorno dell'export).

Il nome compare in TRE punti, e ci sono tutti e tre:

  1. `xl/workbook.xml`      -> l'elemento `<sheet name="...">`
  2. `xl/worksheets/*.xml`  -> `'Only Cases W33'!$B$5`, citato da 'OC Eventi'
  3. `docProps/app.xml`     -> l'elenco `TitlesOfParts`

Nessun grafico lo cita (verificato: la stringa non compare in `xl/charts/`), e
`calcChain.xml` indirizza i fogli per indice, non per nome.

--- PATCH 2: il tetto di AT_DATASET torna a 130.000 -------------------------

Le formule dei due fogli leggono `AT_DATASET` fino a riga **60000**:

    LET(r, SEQUENCE(59999,1,2), s, AT_DATASET!$F$2:$F$60000, FILTER(r, s=stato))
    MINIFS(AT_DATASET!$P$2:$P$60000, AT_DATASET!$F$2:$F$60000, $B$5)

Il resto del template arriva a 130000 (`Report Agenti`, e le formule `P`/`Q`
sono pre-riempite fin li'). Un limite piu' basso in un foglio solo diventa il
limite di TUTTI: `coherence._check_row_limits` prende il piu' stretto, e da 60000
in su BLOCCA il build. Misurato sul W30, l'export ha 26 541 righe — il 44% —
quindi non morde oggi; ma il margine era 4,9x e diventerebbe 2,3x, e la riga in
piu' arriva da sola con il volume.

`SEQUENCE(59999,1,2)` e l'intervallo `$2:$60000` sono la stessa cosa detta due
volte (le righe da 2 a 60000 sono 59999), quindi si spostano INSIEME: 129999 e
130000. Se si muovesse solo uno dei due, `FILTER` confronterebbe due array di
lunghezza diversa e darebbe `#VALUE!` — motivo per cui questo strumento non
riscrive niente se non trova entrambi nella forma attesa.

Sono 6 sostituzioni in tutto: 5 intervalli e 1 `SEQUENCE`.

Cosa NON tocca, e perche'. `OC Eventi` legge anche `'Slot Only Cases'!$A$2:$E$2000`
e `Turni!$E$2:$I$5000`. Sono limiti veri e nuovi, ma stanno dentro 33 000 e
39 000 formule per riga scritte per esteso: allargarli qui vorrebbe dire
riscrivere 72 000 formule e gonfiare il file, per un margine che oggi e' 1291 su
2000 (65%) e 252 su 5000 (5%). Li tiene d'occhio il preflight, che segnala
all'80% — cioe' a 1600 righe di back office. Quando succedera' sara' una
modifica al template, decisa con il numero in mano.

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

NOME_VECCHIO = "Only Cases W33"
NOME_NUOVO = "Only Cases Dashboard"

# Il tetto di AT_DATASET nei due fogli nuovi, e quello del resto del template.
RIGA_VECCHIA = 60000
RIGA_NUOVA = 130000

# `AT_DATASET!$F$2:$F$60000` — un intervallo con la riga finale scritta dentro
# la formula. Si pretende che parta da riga 2: un intervallo che parte altrove
# non e' "tutto il dataset" ma una selezione, e allargarlo cambierebbe il senso.
_INTERVALLO = re.compile(
    rf"AT_DATASET!\$([A-Z]{{1,3}})\$2:\$([A-Z]{{1,3}})\${RIGA_VECCHIA}\b"
)
# `SEQUENCE(59999,1,2)`, con o senza il prefisso `_xlfn.` che Excel scrive per le
# funzioni recenti.
_SEQUENZA = re.compile(
    rf"(SEQUENCE\()({RIGA_VECCHIA - 1})(,1,2\))"
)


def rinomina_foglio(parti: dict[str, bytes]) -> list[str]:
    """Sostituisce il nome del foglio nei tre punti in cui e' scritto."""
    fatte: list[str] = []

    wb = parti["xl/workbook.xml"].decode("utf8")
    if f'name="{NOME_NUOVO}"' in wb:
        return fatte  # gia' rinominato
    if f'name="{NOME_VECCHIO}"' not in wb:
        return fatte  # non c'e' niente da rinominare

    parti["xl/workbook.xml"] = wb.replace(
        f'name="{NOME_VECCHIO}"', f'name="{NOME_NUOVO}"'
    ).encode("utf8")
    fatte.append("xl/workbook.xml: <sheet name=...>")

    # Nelle formule il nome sta fra apici, perche' contiene spazi:
    # `'Only Cases W33'!$B$5`. Si cerca la forma con gli apici e basta: senza,
    # `Only Cases W33` non sarebbe un riferimento valido.
    for nome in [n for n in parti if n.startswith("xl/worksheets/")]:
        raw = parti[nome].decode("utf8")
        n = raw.count(f"'{NOME_VECCHIO}'!")
        if not n:
            continue
        parti[nome] = raw.replace(
            f"'{NOME_VECCHIO}'!", f"'{NOME_NUOVO}'!"
        ).encode("utf8")
        fatte.append(f"{nome}: {n} riferimento/i nelle formule")

    app = parti.get("docProps/app.xml")
    if app and NOME_VECCHIO.encode("utf8") in app:
        parti["docProps/app.xml"] = app.decode("utf8").replace(
            f"<vt:lpstr>{NOME_VECCHIO}</vt:lpstr>",
            f"<vt:lpstr>{NOME_NUOVO}</vt:lpstr>",
        ).encode("utf8")
        fatte.append("docProps/app.xml: TitlesOfParts")

    return fatte


def allarga_tetto(raw: str) -> tuple[str, int, int]:
    """Porta a 130000 gli intervalli su AT_DATASET, e con loro la SEQUENCE.

    Restituisce `(xml, n_intervalli, n_sequenze)`. Non tocca niente se non
    riconosce la forma attesa: e' la stessa guardia dello strumento dei
    duplicati, e serve allo stesso scopo — quando non si e' sicuri di cosa si sta
    guardando, l'unica mossa senza rischi e' lasciare stare.
    """
    nuovo, n_int = _INTERVALLO.subn(
        rf"AT_DATASET!$\1$2:$\2${RIGA_NUOVA}", raw
    )
    nuovo, n_seq = _SEQUENZA.subn(rf"\g<1>{RIGA_NUOVA - 1}\g<3>", nuovo)
    return nuovo, n_int, n_seq


def _fogli(parti: dict[str, bytes]) -> dict[str, str]:
    """`{nome foglio: parte dello zip}`."""
    wb = parti["xl/workbook.xml"].decode("utf8", "replace")
    rels = parti["xl/_rels/workbook.xml.rels"].decode("utf8", "replace")
    relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    out: dict[str, str] = {}
    for m in re.finditer(r"<sheet ([^>]*)/>", wb):
        n = re.search(r'name="([^"]*)"', m.group(1))
        rid = re.search(r'r:id="(rId\d+)"', m.group(1))
        if not (n and rid and rid.group(1) in relmap):
            continue
        t = relmap[rid.group(1)].lstrip("/")
        out[n.group(1)] = t if t.startswith("xl/") else "xl/" + t
    return out


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("workbook", type=Path)
    ap.add_argument("--dry-run", action="store_true", help="dice cosa farebbe, non scrive")
    ap.add_argument(
        "--no-backup", action="store_true",
        help="non salva la copia di sicurezza (di norma la si vuole)",
    )
    args = ap.parse_args(argv)

    if not args.workbook.is_file():
        print(f"File non trovato: {args.workbook}", file=sys.stderr)
        return 2

    with zipfile.ZipFile(args.workbook) as z:
        ordine = z.namelist()
        parti = {n: z.read(n) for n in ordine}

    print(f"PATCH 1 — {NOME_VECCHIO!r} -> {NOME_NUOVO!r}")
    rinominate = rinomina_foglio(parti)
    if rinominate:
        for r in rinominate:
            print(f"  {r}")
    else:
        print("  niente da fare: il foglio ha gia' il nome stabile (o non c'e').")

    # Dopo il rename, i fogli si ritrovano con i nomi aggiornati.
    fogli = _fogli(parti)
    mancanti = [n for n in (NOME_NUOVO, "OC Eventi") if n not in fogli]
    if mancanti:
        print(
            f"Il workbook non contiene {', '.join(repr(m) for m in mancanti)}.",
            file=sys.stderr,
        )
        return 1

    print(f"\nPATCH 2 — tetto AT_DATASET {RIGA_VECCHIA} -> {RIGA_NUOVA}")
    tot_int = tot_seq = 0
    for nome in (NOME_NUOVO, "OC Eventi"):
        target = fogli[nome]
        raw = parti[target].decode("utf8")
        nuovo, n_int, n_seq = allarga_tetto(raw)
        parti[target] = nuovo.encode("utf8")
        tot_int += n_int
        tot_seq += n_seq
        if n_int or n_seq:
            print(f"  {nome}: {n_int} intervalli, {n_seq} SEQUENCE")
    if not (tot_int or tot_seq):
        print(f"  niente da fare: nessun riferimento a riga {RIGA_VECCHIA}.")

    # I due numeri si muovono insieme o non si muove niente: un intervallo
    # allargato senza la sua SEQUENCE farebbe confrontare a FILTER due array di
    # lunghezza diversa, cioe' `#VALUE!` al primo ricalcolo.
    if tot_int and not tot_seq:
        print(
            f"\nRIFIUTO: {tot_int} intervalli a {RIGA_VECCHIA} ma nessuna "
            f"SEQUENCE({RIGA_VECCHIA - 1}).\n"
            f"  Allargare solo gli intervalli romperebbe FILTER. Non ho scritto niente.",
            file=sys.stderr,
        )
        return 1

    if not rinominate and not tot_int and not tot_seq:
        print("\nIl template e' gia' patchato. Nessuna scrittura.")
        return 0

    if args.dry_run:
        print("\n--dry-run: non ho scritto niente.")
        return 0

    if not args.no_backup:
        backup = args.workbook.with_name(
            f"{args.workbook.stem}_prima_delle_patch_only_cases{args.workbook.suffix}"
        )
        shutil.copy2(args.workbook, backup)
        print(f"\nCopia di sicurezza: {backup.name}")

    # Si riscrive il pacchetto nell'ORDINE originale delle parti, come fa lo
    # strumento dei duplicati: un diff fra i due zip resta leggibile.
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
