#!/usr/bin/env python3
"""Confronta un workbook generato dalla pipeline con la stessa settimana fatta a mano.

E' il collaudo che conta: il golden test delle sorgenti WFM verifica
l'ingestione, questo verifica i **numeri finali** — `Report Agenti`,
`Malpractice Recap`, `Outbound Exploitation`, `Recap PSAT Positive`. Se tornano,
la pipeline sostituisce il processo manuale; se non tornano, la differenza dice
dove guardare.

    python tools/golden_report.py \\
        --generato   output/Omni_Report_W30.xlsm \\
        --riferimento "samples/omni-report/Omni Report W30.xlsm"

Legge i **valori in cache** di entrambi i file, senza aprire Excel. Il workbook
generato li ha perche' `build` salva dopo il ricalcolo completo.

Cosa aspettarsi la prima volta: qualche differenza c'e' quasi sempre, e non
tutte sono bug. Le tre note sono dichiarate qui sotto (`ATTESE`). Tutto il resto
va guardato: l'output raggruppa per foglio e mostra le prime celle divergenti,
cosi' si capisce se e' un numero o un'intera colonna.
"""

from __future__ import annotations

import argparse
import sys
from dataclasses import dataclass, field
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from fasterreports.core.xlsxsource import col_to_index, read_sheet, read_sheet_names  # noqa: E402

# I fogli che contengono i numeri del report, in ordine di importanza.
FOGLI_OUTPUT = [
    "Report Agenti",
    "Malpractice Recap",
    "Dettaglio Malpractice",
    "Outbound Exploitation",
    "Recap PSAT Positive",
    "Analisi Status",
    "Verifica AHT",
    "AHT Outliers",
    "Helper Turni",
]

# Fogli che NON si confrontano: sono gli input, e sappiamo gia' che differiscono
# per le tre ragioni dichiarate nel golden test delle sorgenti.
FOGLI_INPUT = {
    "AT_DATASET", "ATwi_DATASET", "SF_DATABASE", "PSAT_DATASET",
    "Turni", "Slot Only Cases",
}

ATTESE = """Differenze attese, dichiarate in anticipo:
  · 'Turni' e 'Slot Only Cases': celle vuote invece della spazzatura del
    processo manuale sulle righe non lavorate, e 'NO BOT' scritto per esteso.
    Nessuna formula le legge (filtrano su Stato="LAVORA"), quindi non
    propagano — ma per questo i fogli di input non si confrontano affatto.
  · 'Recap PSAT Positive': l'"Elogio della settimana" punta alla riga FISSA
    130 di PSAT_DATASET, cioe' e' una scelta manuale. Se il numero di righe
    PSAT cambia, quella cella pesca un commento diverso."""


@dataclass
class DiffFoglio:
    nome: str
    confrontate: int = 0
    diverse: list[tuple[str, object, object]] = field(default_factory=list)
    solo_generato: int = 0
    solo_riferimento: int = 0
    assente_in: str = ""

    @property
    def quota(self) -> float:
        return len(self.diverse) / self.confrontate if self.confrontate else 0.0


def _num(v):
    if v is None:
        return None
    s = str(v).strip()
    if not s:
        return None
    try:
        return float(s.replace(",", "."))
    except ValueError:
        return s


def _uguali(a, b, tol: float) -> bool:
    a, b = _num(a), _num(b)
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    if isinstance(a, float) and isinstance(b, float):
        if a == b:
            return True
        scala = max(abs(a), abs(b), 1.0)
        return abs(a - b) <= tol * scala
    return str(a) == str(b)


def _leggi(path: Path, nome: str, etichetta: str) -> tuple[object | None, str]:
    """Legge un foglio distinguendo 'non c'e'' da 'non si legge'.

    Confonderli manda fuori strada: un file corrotto verrebbe archiviato come
    "foglio mancante" e si cercherebbe il problema nel posto sbagliato.
    """
    try:
        if nome not in read_sheet_names(path):
            return None, f"assente nel {etichetta}"
    except Exception as exc:
        return None, f"{etichetta} non leggibile: {exc}"
    try:
        return read_sheet(path, nome), ""
    except Exception as exc:
        return None, f"{etichetta} illeggibile ({type(exc).__name__}): {exc}"


def confronta_foglio(gen_path: Path, rif_path: Path, nome: str, tol: float,
                     max_col: int) -> DiffFoglio:
    out = DiffFoglio(nome=nome)
    gen, problema = _leggi(gen_path, nome, "generato")
    if gen is None:
        out.assente_in = problema
        return out
    rif, problema = _leggi(rif_path, nome, "riferimento")
    if rif is None:
        out.assente_in = problema
        return out

    righe = set(gen.rows) | set(rif.rows)
    for r in sorted(righe):
        g, f = gen.row(r), rif.row(r)
        for col in sorted(set(g) | set(f), key=col_to_index):
            if col_to_index(col) > max_col:
                continue
            gv, fv = g.get(col), f.get(col)
            if gv is None and fv is None:
                continue
            if gv is None:
                out.solo_riferimento += 1
                continue
            if fv is None:
                out.solo_generato += 1
                continue
            out.confrontate += 1
            if not _uguali(gv, fv, tol):
                out.diverse.append((f"{col}{r}", gv, fv))
    return out


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("--generato", type=Path, required=True)
    ap.add_argument("--riferimento", type=Path, required=True)
    ap.add_argument("--fogli", nargs="*", help="default: i fogli di output")
    ap.add_argument("--tol", type=float, default=1e-6,
                    help="tolleranza relativa sui numeri (default 1e-6)")
    ap.add_argument("--max-col", type=int, default=60,
                    help="ultima colonna da confrontare (default 60 = BH)")
    ap.add_argument("--esempi", type=int, default=6)
    args = ap.parse_args(argv)

    for p in (args.generato, args.riferimento):
        if not p.is_file():
            print(f"File non trovato: {p}", file=sys.stderr)
            return 2

    if args.fogli:
        fogli = args.fogli
    else:
        presenti = set(read_sheet_names(args.generato)) & set(read_sheet_names(args.riferimento))
        fogli = [f for f in FOGLI_OUTPUT if f in presenti]
        # Fogli di output che non avevo previsto: meglio confrontarli che ignorarli.
        extra = sorted(presenti - set(FOGLI_OUTPUT) - FOGLI_INPUT)
        fogli += extra

    print(f"generato:     {args.generato}")
    print(f"riferimento:  {args.riferimento}")
    print(f"tolleranza relativa: {args.tol:g} · colonne fino a "
          f"{args.max_col} · fogli: {len(fogli)}\n")

    risultati = [
        confronta_foglio(args.generato, args.riferimento, f, args.tol, args.max_col)
        for f in fogli
    ]

    identici, con_diff, assenti = [], [], []
    for d in risultati:
        if d.assente_in:
            assenti.append(d)
        elif d.diverse or d.solo_generato or d.solo_riferimento:
            con_diff.append(d)
        else:
            identici.append(d)

    print("=" * 74)
    if identici:
        print(f"IDENTICI ({len(identici)}):")
        for d in identici:
            print(f"  · {d.nome}  ({d.confrontate} celle)")
    if assenti:
        print("\nNON CONFRONTATI:")
        for d in assenti:
            print(f"  ! {d.nome}: {d.assente_in}")

    for d in con_diff:
        print(f"\n{'-' * 74}\n{d.nome}")
        print(f"  celle confrontate: {d.confrontate}   diverse: {len(d.diverse)}"
              f" ({d.quota:.1%})")
        if d.solo_generato:
            print(f"  celle solo nel generato:    {d.solo_generato}")
        if d.solo_riferimento:
            print(f"  celle solo nel riferimento: {d.solo_riferimento}")
        if d.diverse:
            # Le colonne piu' colpite dicono se e' un numero o una formula intera.
            colonne: dict[str, int] = {}
            for ref, _g, _f in d.diverse:
                c = "".join(ch for ch in ref if ch.isalpha())
                colonne[c] = colonne.get(c, 0) + 1
            top = sorted(colonne.items(), key=lambda kv: -kv[1])[:6]
            print(f"  colonne piu' colpite: "
                  + ", ".join(f"{c} ({n})" for c, n in top))
            for ref, g, f in d.diverse[:args.esempi]:
                print(f"    {ref:>7}  generato={g!r:<22} riferimento={f!r}")
            if len(d.diverse) > args.esempi:
                print(f"    ... e altre {len(d.diverse) - args.esempi}")

    print("\n" + "=" * 74)
    tot_diff = sum(len(d.diverse) for d in con_diff)
    tot_celle = sum(d.confrontate for d in risultati)
    if not con_diff and not assenti:
        print(f"ESITO: tutti i {len(identici)} fogli identici ({tot_celle} celle).")
        print("La pipeline riproduce il processo manuale.")
        return 0
    print(f"ESITO: {len(con_diff)} fogli con differenze, {tot_diff} celle su {tot_celle}.")
    print()
    print(ATTESE)
    print(
        "\nCome leggerlo: una colonna intera diversa e' una formula che legge la\n"
        "colonna sbagliata o un input mancante; poche celle sparse sono piu'\n"
        "probabilmente arrotondamenti o scelte manuali. Parti dal foglio con la\n"
        "quota piu' alta."
    )
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
