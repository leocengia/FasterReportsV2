#!/usr/bin/env python3
"""Ricostruisce lo storico AHT da export SF o da Omni Report gia' prodotti.

A cosa serve. Lo storico seminato dal foglio 'AHT History' del template era
stato costruito a mano dal file WOW, e ha ereditato la sua lista curata di case
type: confrontandolo con il workbook consegnato del W30, volumi e AHT
combaciavano al centesimo su tutte le 45 combinazioni presenti, ma ne mancavano
9 — fra cui 'Call Assignment' e 'Specialty Functions', le stesse assenti da
'Helper CaseType'. Le heat map di quelle settimane erano quindi incomplete, non
sbagliate.

La cosa importante: **non serve rigenerare gli Omni Report.** L'export SF di
ogni settimana passata e' ancora dentro il workbook che quella settimana ha
prodotto, nel foglio SF_DATABASE. Questo strumento lo legge da li'.

  python tools/ricostruisci_storico.py output/Omni_Report_W3*.xlsm
  python tools/ricostruisci_storico.py "input/SF DATABASE W33.csv" --dry-run

Accetta indifferentemente:
  · un .xlsm/.xlsx  -> legge il foglio SF_DATABASE
  · un .csv         -> lo tratta come export SF

Le colonne si cercano per NOME nell'intestazione, come fa tutta la pipeline:
funziona sia sul workbook (dove stanno nelle posizioni del contratto) sia su un
export che ha cambiato ordine. La settimana si ricava da 'Date Viewpoint', e
deve cadere di lunedi': se non lo fa, il file viene saltato con un messaggio,
non archiviato sotto una settimana inventata.

Rilanciarlo sugli stessi file e' innocuo: l'aggiornamento dello storico
sostituisce le righe di quella settimana, non le somma.
"""

from __future__ import annotations

import argparse
import csv
import sys
from collections import Counter
from datetime import date, datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from fasterreports.core import aht_history as H  # noqa: E402
from fasterreports.core.coerce import Uncoercible, coerce  # noqa: E402
from fasterreports.core.errors import PipelineError  # noqa: E402
from fasterreports.core.normalize import normalize_raw  # noqa: E402

CANALE = "Case Origin (group)"
TIPO = "Case Type"
AHT = "Case AHT (mins)"
DATA = "Date Viewpoint"
FOGLIO = "SF_DATABASE"


def _mappa(intestazioni: dict[str, str]) -> dict[str, str]:
    """{nome canonico: lettera} risolvendo per nome normalizzato.

    Per NOME, non per posizione: e' cio' che fa funzionare lo stesso strumento
    sul foglio del workbook (colonne nelle posizioni del contratto) e su un
    export che ha cambiato ordine — come e' successo fra il W30 e il W33.
    """
    per_nome = {normalize_raw(v): k for k, v in intestazioni.items() if v}
    out = {}
    for nome in (CANALE, TIPO, AHT, DATA):
        lettera = per_nome.get(normalize_raw(nome))
        if lettera:
            out[nome] = lettera
    return out


def leggi_xlsx(path: Path):
    from fasterreports.core.xlsxsource import read_sheet

    sh = read_sheet(path, FOGLIO)
    righe = sorted(sh.rows)
    if not righe:
        raise PipelineError(f"{path}: il foglio {FOGLIO} e' vuoto.")
    mappa = _mappa(sh.row(1))
    mancanti = [n for n in (CANALE, TIPO, AHT) if n not in mappa]
    if mancanti:
        raise PipelineError(
            f"{path}: nel foglio {FOGLIO} non trovo {', '.join(mancanti)}."
        )
    terne = [
        (sh.cell(mappa[CANALE], r), sh.cell(mappa[TIPO], r), sh.cell(mappa[AHT], r))
        for r in righe if r > 1
    ]
    date_viewpoint = (
        [sh.cell(mappa[DATA], r) for r in righe if r > 1] if DATA in mappa else []
    )
    return terne, date_viewpoint


def leggi_csv(path: Path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        header = next(reader)
        idx = {}
        per_nome = {normalize_raw(h): i for i, h in enumerate(header)}
        for nome in (CANALE, TIPO, AHT, DATA):
            i = per_nome.get(normalize_raw(nome))
            if i is not None:
                idx[nome] = i
        mancanti = [n for n in (CANALE, TIPO, AHT) if n not in idx]
        if mancanti:
            raise PipelineError(f"{path}: non trovo {', '.join(mancanti)}.")
        terne, date_viewpoint = [], []
        for row in reader:
            def g(nome):
                i = idx.get(nome)
                return row[i] if i is not None and i < len(row) else None
            terne.append((g(CANALE), g(TIPO), g(AHT)))
            if DATA in idx:
                date_viewpoint.append(g(DATA))
    return terne, date_viewpoint


def settimana(valori, formato: str) -> tuple[int, int]:
    """(anno ISO, settimana) dal lunedi' prevalente in 'Date Viewpoint'."""
    giorni: Counter[date] = Counter()
    for v in valori:
        if v is None or (isinstance(v, str) and not v.strip()):
            continue
        try:
            d = coerce(v, "datetime", formato if isinstance(v, str) and "/" in v else "")
        except Uncoercible:
            continue
        if isinstance(d, datetime):
            giorni[d.date()] += 1
    if not giorni:
        raise PipelineError(
            "nessun valore leggibile in 'Date Viewpoint': non so a che settimana "
            "appartiene questo file. Passa --week per dirlo a mano."
        )
    prevalente = giorni.most_common(1)[0][0]
    if prevalente.weekday() != 0:
        raise PipelineError(
            f"'Date Viewpoint' vale {prevalente.isoformat()}, che non e' un lunedi'.\n"
            f"  Quasi sempre significa che la data e' stata letta al contrario.\n"
            f"  Prova un --date-format diverso, o passa --week a mano."
        )
    iso = prevalente.isocalendar()
    return (iso[0], iso[1])


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("files", type=Path, nargs="+")
    ap.add_argument("--out", type=Path, default=Path("data/aht_history.csv"))
    ap.add_argument(
        "--date-format", default="%m/%d/%Y %I:%M:%S %p",
        help="formato di 'Date Viewpoint' quando e' testo (default: americano, "
             "come l'export Tableau)",
    )
    ap.add_argument(
        "--week", type=int, default=None,
        help="forza anno e settimana per TUTTI i file, nella forma AAAASS "
             "(es. 202630). Usalo solo se 'Date Viewpoint' manca.",
    )
    ap.add_argument(
        "--settings", type=Path, default=Path("config/settings.yml"),
        help="da dove leggere aht_history.casetype_esclusi (default: config/settings.yml)",
    )
    ap.add_argument(
        "--tutti-i-casetype", action="store_true",
        help="ignora le esclusioni e includi ogni case type. Serve per capire cosa "
             "verrebbe escluso, non per produrre lo storico buono.",
    )
    ap.add_argument("--dry-run", action="store_true", help="mostra cosa cambierebbe, non scrive")
    args = ap.parse_args(argv)

    esclusi: tuple[str, ...] = ()
    if not args.tutti_i_casetype:
        try:
            from fasterreports.omni.settings import load_settings

            esclusi = load_settings(
                args.settings, root=args.settings.resolve().parent.parent
            ).casetype_esclusi
        except Exception as exc:  # config assente o illeggibile: si dice, non si tace
            print(f"ATTENZIONE: non ho letto le esclusioni da {args.settings} ({exc}).",
                  file=sys.stderr)
            print("  Lo storico includerebbe OGNI case type, anche quelli che nelle "
                  "heat map non devono starci.", file=sys.stderr)
            return 2
    print(f"case type esclusi dal trend: {len(esclusi)}"
          + (f" ({', '.join(esclusi)})" if esclusi else " — nessuno"))

    storico = H.carica(args.out)
    prima = {r.chiave: r for r in storico}
    print(f"storico di partenza: {len(storico)} righe, "
          f"{len(H.settimane(storico))} settimane\n")

    problemi = 0
    for path in args.files:
        if not path.is_file():
            print(f"  SALTATO {path}: non esiste", file=sys.stderr)
            problemi += 1
            continue
        try:
            if path.suffix.lower() in (".xlsm", ".xlsx"):
                terne, dv = leggi_xlsx(path)
            else:
                terne, dv = leggi_csv(path)
            if args.week:
                iso_year, week = divmod(args.week, 100)
            else:
                iso_year, week = settimana(dv, args.date_format)
        except PipelineError as exc:
            print(f"  SALTATO {path.name}: {exc}", file=sys.stderr)
            problemi += 1
            continue

        nuove = H.aggrega(terne, iso_year=iso_year, week=week, esclusi=esclusi)
        vecchie = [r for r in storico if r.week_key == iso_year * 100 + week]
        storico = H.unisci(storico, nuove)
        agg = len(nuove) - len(vecchie)
        print(f"  {path.name}: {iso_year}-W{week:02d} · {len(terne)} casi · "
              f"{len(nuove)} combinazioni"
              + (f" ({agg:+d} rispetto alle {len(vecchie)} che c'erano)" if vecchie
                 else " (settimana nuova)"))

    print()
    print(f"storico finale: {len(storico)} righe, {len(H.settimane(storico))} settimane")

    # Le differenze si contano con una tolleranza, e separando i due casi. Senza
    # la tolleranza ogni riga risulta "cambiata": il CSV di partenza porta l'AHT
    # con 15 cifre e il ricalcolo ne produce altrettante, ma non identiche
    # all'ultimo bit. Un contatore che dice "29 righe cambiate" quando non e'
    # cambiato niente insegna a ignorarlo, che e' peggio di non averlo.
    nuove_righe = [r for r in storico if r.chiave not in prima]
    diverse = [
        (prima[r.chiave], r) for r in storico
        if r.chiave in prima and (
            prima[r.chiave].volume != r.volume
            or abs(prima[r.chiave].aht - r.aht) > 0.005
        )
    ]
    print(f"  combinazioni nuove: {len(nuove_righe)}")
    print(f"  valori cambiati (oltre 0.005 sull'AHT o volume diverso): {len(diverse)}")
    for vecchia, nuova in diverse[:10]:
        print(f"    {nuova.week}  {nuova.channel:9} {nuova.case_type[:34]:34} "
              f"vol {vecchia.volume}->{nuova.volume}  "
              f"aht {vecchia.aht:.2f}->{nuova.aht:.2f}")
    if len(diverse) > 10:
        print(f"    ... e altre {len(diverse) - 10}")
    if not nuove_righe and not diverse:
        print("  => il ricalcolo conferma lo storico che c'era, riga per riga.")

    if args.dry_run:
        print("\n--dry-run: niente scritto.")
    else:
        H.scrivi(args.out, storico)
        print(f"\nScritto in {args.out}. Ricordati di committarlo.")
    return 1 if problemi else 0


if __name__ == "__main__":
    raise SystemExit(main())
