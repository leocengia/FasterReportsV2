#!/usr/bin/env python3
"""Compila una copia di `August_Transfer_Data_Template_v2.xlsx` con i numeri misurati.

Non usa openpyxl: riscrive `xl/worksheets/sheet1.xml` dentro una COPIA dello zip,
lasciando intatti stili, formule e tutto il resto del pacchetto. openpyxl
riscriverebbe il file da capo e con lui se ne andrebbero formattazione e
proprieta' — e il progetto lo vieta comunque in scrittura.

    python analisi/transfer-mbr-agosto/compila_template.py --breakdown analisi/transfer-mbr-agosto/breakdown_W34.md

Le correzioni ai difetti del template (formula D20, anno 2025->2026) sono
esplicite qui sotto e documentate in `rilievi_template.md`: non sono ritocchi
silenziosi.
"""

from __future__ import annotations

import argparse
import re
import shutil
import zipfile
from pathlib import Path

QUI = Path(__file__).parent
SORGENTE = QUI / "August_Transfer_Data_Template_v2.xlsx"
FOGLIO = "xl/worksheets/sheet1.xml"


def _esc(testo: str) -> str:
    return (testo.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def scrivi_cella(xml: str, rif: str, valore, formula: str | None = None) -> str:
    """Sostituisce il contenuto della cella `rif`, conservandone lo stile `s`.

    Se `valore` e' un numero lo scrive come numero; se e' testo usa una stringa
    inline, cosi' non si tocca `sharedStrings.xml` e non si spostano gli indici
    delle stringhe gia' presenti.
    """
    m = re.search(rf'<c r="{rif}"((?: [a-zA-Z:]+="[^"]*")*)\s*(/>|>.*?</c>)', xml, re.S)
    if not m:
        raise SystemExit(f"cella {rif} non trovata nel foglio")
    attributi = m.group(1)
    attributi = re.sub(r'\s+t="[^"]*"', "", attributi)   # il tipo lo decidiamo qui

    if formula is not None:
        corpo = f"<f>{_esc(formula)}</f>" + (f"<v>{valore}</v>" if valore is not None else "")
        tipo = ' t="str"' if isinstance(valore, str) else ""
    elif valore is None or valore == "":
        corpo, tipo = "", ""
    elif isinstance(valore, (int, float)):
        corpo, tipo = f"<v>{valore}</v>", ""
    else:
        corpo, tipo = f"<is><t>{_esc(str(valore))}</t></is>", ' t="inlineStr"'

    nuova = f'<c r="{rif}"{attributi}{tipo}>{corpo}</c>' if corpo else f'<c r="{rif}"{attributi}/>'
    return xml[:m.start()] + nuova + xml[m.end():]


def compila(dati: dict, destinazione: Path) -> None:
    if not SORGENTE.exists():
        raise SystemExit(f"template di partenza non trovato: {SORGENTE}")
    shutil.copy(SORGENTE, destinazione)

    origine = zipfile.ZipFile(SORGENTE)
    xml = origine.read(FOGLIO).decode("utf-8")
    stringhe = origine.read("xl/sharedStrings.xml").decode("utf-8")

    # --- Rilievo 1: le intestazioni dicono 2025, ma la richiesta e' per il 2026.
    stringhe = stringhe.replace("August 2025", "August 2026")

    adv_v, adv_b = dati["advanced_voice"], dati["advanced_bof"]
    bas_v, bas_b = dati["basic_voice"], dati["basic_bof"]
    tot_v, tot_b = adv_v + bas_v, adv_b + bas_b

    # Sezione 1 — destinazione x canale di provenienza.
    xml = scrivi_cella(xml, "B9", adv_v)
    xml = scrivi_cella(xml, "C9", adv_b)
    xml = scrivi_cella(xml, "B10", bas_v)
    xml = scrivi_cella(xml, "C10", bas_b)
    # Le due somme sono gia' formule del template: si aggiorna solo il valore in
    # cache, cosi' il file mostra il numero giusto anche prima di un ricalcolo.
    xml = scrivi_cella(xml, "B11", tot_v, formula="IFERROR(SUM(B9:B10),0)")
    xml = scrivi_cella(xml, "C11", tot_b, formula="IFERROR(SUM(C9:C10),0)")

    # Rilievo 12: la sezione 1 non ha una colonna Combined, quindi il totale di
    # riga promesso da A7 non esiste. Va nelle note.
    xml = scrivi_cella(xml, "D9", f"Row total: {adv_v + adv_b}. Destination inferred from "
                                  f"work_function of the receiving case (no 'transferred to "
                                  f"queue' field exists in the export).")
    xml = scrivi_cella(xml, "D10", f"Row total: {bas_v + bas_b}. Voice = 'Transferred from "
                                   f"Phone'; BOF = 'Transferred from Live Agent'.")

    # Sezione 2 — stessi totali della sezione 1 (rilievo 14: niente li collega).
    xml = scrivi_cella(xml, "B16", tot_v)
    xml = scrivi_cella(xml, "C16", tot_b)
    xml = scrivi_cella(xml, "D16", tot_v + tot_b, formula="IFERROR(B16+C16,0)")

    # --- Rilievo 10: D20 sta nella colonna Combined ma calcolava B19/B16,
    # cioe' il solo Voice. Corretta in D19/D16.
    xml = scrivi_cella(xml, "D20", "—", formula='IFERROR(D19/D16,"—")')
    # Rilievo 11: B20/C20 erano input a mano di valori derivabili. Ora sono formule.
    xml = scrivi_cella(xml, "B20", "—", formula='IFERROR(B19/B16,"—")')
    xml = scrivi_cella(xml, "C20", "—", formula='IFERROR(C19/C16,"—")')

    xml = scrivi_cella(xml, "B4", dati["sito"])
    xml = scrivi_cella(xml, "F16", dati["nota_periodo"])
    xml = scrivi_cella(xml, "F17", dati["nota_scrub"])

    # La calcChain elenca le formule nell'ordine in cui Excel le ha calcolate.
    # Avendone aggiunte due (B20, C20), la via sicura e' toglierla: Excel la
    # ricostruisce da solo alla prima apertura. Lasciarla disallineata e' l'unico
    # modo per far comparire un avviso di file danneggiato.
    #
    # Togliere la parte non basta: [Content_Types].xml la dichiara, e un Override
    # che punta a una parte assente e' anch'esso un file danneggiato. Le due cose
    # vanno fatte insieme, nella stessa scrittura.
    tipi = origine.read("[Content_Types].xml").decode("utf-8")
    tipi = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", tipi)

    sostituzioni = {
        FOGLIO: xml,
        "xl/sharedStrings.xml": stringhe,
        "[Content_Types].xml": tipi,
    }
    with zipfile.ZipFile(destinazione, "w", zipfile.ZIP_DEFLATED) as fuori:
        for voce in origine.infolist():
            if voce.filename == "xl/calcChain.xml":
                continue
            if voce.filename in sostituzioni:
                fuori.writestr(voce, sostituzioni[voce.filename])
            else:
                fuori.writestr(voce, origine.read(voce.filename))


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--advanced-voice", type=int, required=True)
    ap.add_argument("--advanced-bof", type=int, required=True)
    ap.add_argument("--basic-voice", type=int, required=True)
    ap.add_argument("--basic-bof", type=int, required=True)
    ap.add_argument("--sito", default="Milan (Orchidea) — EMEA")
    ap.add_argument("--periodo", required=True, help="testo del periodo coperto")
    ap.add_argument("--out", type=Path, required=True)
    args = ap.parse_args(argv)

    compila(
        {
            "advanced_voice": args.advanced_voice,
            "advanced_bof": args.advanced_bof,
            "basic_voice": args.basic_voice,
            "basic_bof": args.basic_bof,
            "sito": args.sito,
            "nota_periodo": args.periodo,
            "nota_scrub": "Scrub in progress — see methodology annex. Rows 17-19 to be "
                          "completed once case-level marking is done.",
        },
        args.out,
    )
    print(f"scritto {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
