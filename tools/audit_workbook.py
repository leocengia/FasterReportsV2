#!/usr/bin/env python3
"""Ispezione di un workbook Excel senza aprire Excel — e senza openpyxl.

Legge direttamente l'XML dentro il file .xlsm/.xlsx. Non carica il foglio in
memoria come oggetto, quindi funziona anche sul WOW da 60 MB dove openpyxl in
scrittura va in OOM (contesto WOW §5).

Tre usi:

  --headers        intestazioni di riga 1 dei fogli indicati (o di tutti)
  --usage          quali colonne dei dataset sono consumate da quali formule
                   -> e' il test di auto-verifica del contratto (piano §11)
  --tables         i ListObject e se il loro intervallo copre i dati

  python tools/audit_workbook.py "samples/Omni Report W30.xlsm" --all
  python tools/audit_workbook.py FILE --usage --datasets AT_DATASET SF_DATABASE
  python tools/audit_workbook.py FILE --headers --json > tests/fixtures/headers.json

Nota sul matching dei nomi foglio nelle formule: `AT_DATASET` e' una
sottostringa di `PSAT_DATASET`. Senza un confine a sinistra si attribuiscono ad
AT tutte le colonne di PSAT — errore facile e silenzioso.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from collections import defaultdict
from pathlib import Path

# Tutti e SEI i fogli dati, non solo i 4 CSV. `Turni` e `Slot Only Cases` sono
# entrati nel contratto dopo, e restando fuori da questa lista la fixture dei test
# nasceva senza di loro: un intero pezzo di contratto non verificato.
DEFAULT_DATASETS = [
    "AT_DATASET",
    "ATwi_DATASET",
    "SF_DATABASE",
    "PSAT_DATASET",
    "Turni",
    "Slot Only Cases",
    # DUP_DATASET entra qui insieme al suo dataset nel contratto, non prima:
    # `test_tutti_i_dataset_del_contratto_sono_nei_default` confronta le due
    # liste in ENTRAMBE le direzioni, e un nome qui senza contratto e' un
    # foglio che si prova a leggere senza sapere cosa contiene.
]

# Percorso del contratto rispetto alla radice del repo. Da qui si legge in che
# riga stanno le intestazioni di ciascun foglio dati, invece di riscriverlo:
# `header_row` e' gia' dichiarato nel contratto, ed e' quello che la pipeline
# usa per scrivere. Due numeri che possono divergere sarebbero il difetto che
# questo file serve a scoprire.
CONTRATTO = Path(__file__).resolve().parents[1] / "config" / "columns.yml"


def header_rows(contratto: Path = CONTRATTO) -> dict[str, int]:
    """`{nome foglio: riga delle intestazioni}` dal contratto.

    Best-effort: questo file e' uno strumento diagnostico e deve funzionare anche
    su un workbook di cui non si ha il contratto (o con un contratto rotto —
    magari e' proprio quello che si sta cercando). Se non si riesce a leggerlo, si
    assume riga 1 per tutti, che e' il caso di cinque dataset su sette.
    """
    try:
        import yaml

        body = yaml.safe_load(contratto.read_text(encoding="utf-8")) or {}
        return {
            str(d.get("sheet", nome)): int(d.get("header_row", 1))
            for nome, d in (body.get("datasets") or {}).items()
            if isinstance(d, dict)
        }
    except Exception:
        return {}


# Una cella del foglio: `<c r="B14" s="410" t="s"><v>9</v></c>` — ma anche
# `<c r="A14" s="396"/>`, cioe' formattata e VUOTA.
#
# Il ramo `/>` non e' una rifinitura. Con un pattern che pretende
# `>...</c>`, una cella auto-chiusa non chiude il match: questo prosegue fino al
# primo `</c>` che trova, cioe' **si mangia il valore della cella successiva** e
# lo attribuisce a quella vuota. Finora non si vedeva perche' tutti i fogli dati
# avevano le intestazioni in riga 1 a partire da A senza buchi. La riga 14 di
# `DUP_DATASET` ha A e C vuote-ma-formattate, e il risultato misurato era
# `A: '9136'`, `C: '1364'` (indici grezzi di sharedStrings) con `Full Name` e
# `Case Number` **spariti** dalla fixture.
#
# `(?:[^>"]|"[^"]*")*?` invece di `[^>]*`: gli attributi possono contenere `>`
# dentro le virgolette.
_CELL = re.compile(
    r'<c r="([A-Z]+)(\d+)"((?:[^>"]|"[^"]*")*?)(?:/>|>(.*?)</c>)', re.S
)

# Una formula: `<f>SUM(A1:A9)</f>`, ma anche `<f t="shared" si="6"/>` — le
# formule condivise, che nel corpo non hanno testo. Stesso difetto del pattern
# delle celle: `<f[^>]*>` considera `<f t="shared" si="6"/>` un tag di
# APERTURA (il `[^>]*` si mangia anche lo slash), e cattura tutto fino al
# `</f>` successivo — incollando insieme il testo di celle diverse. In questo
# workbook le formule condivise sono 6993.
_FORMULA = re.compile(r"<f(?:[^>\"]|\"[^\"]*\")*?(?:/>|>(.*?)</f>)", re.S)


def formule(raw: str) -> list[str]:
    """Il testo di tutte le formule di un foglio, senza le condivise vuote."""
    return [_unescape(f) for f in _FORMULA.findall(raw) if f]


def col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def col_letters(index: int) -> str:
    out = ""
    while index:
        index, rem = divmod(index - 1, 26)
        out = chr(65 + rem) + out
    return out


class Workbook:
    def __init__(self, path: Path) -> None:
        self.path = path
        self.z = zipfile.ZipFile(path)
        wb = self.z.read("xl/workbook.xml").decode("utf8", "replace")
        rels = self.z.read("xl/_rels/workbook.xml.rels").decode("utf8", "replace")
        relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
        self.sheets: dict[str, str] = {}
        self.hidden: set[str] = set()
        for m in re.finditer(r"<sheet ([^>]*)/>", wb):
            attrs = m.group(1)
            name = re.search(r'name="([^"]*)"', attrs)
            rid = re.search(r'r:id="(rId\d+)"', attrs)
            if not (name and rid):
                continue
            target = relmap.get(rid.group(1), "").lstrip("/")
            if target and not target.startswith("xl/"):
                target = "xl/" + target
            # I nomi vanno de-escapati: nel workbook.xml il foglio si chiama
            # `DC Agents &amp; Categories`, non `DC Agents & Categories`.
            # `templatescan._sheet_targets` lo faceva gia'; qui no, e lo stesso
            # foglio risultava con due nomi diversi a seconda di chi lo leggeva.
            nome = _unescape(name.group(1))
            self.sheets[nome] = target
            if 'state="hidden"' in attrs or 'state="veryHidden"' in attrs:
                self.hidden.add(nome)
        self._shared: list[str] | None = None

    @property
    def shared(self) -> list[str]:
        if self._shared is None:
            self._shared = []
            if "xl/sharedStrings.xml" in self.z.namelist():
                data = self.z.read("xl/sharedStrings.xml").decode("utf8", "replace")
                for si in re.finditer(r"<si>(.*?)</si>", data, re.S):
                    self._shared.append(
                        "".join(re.findall(r"<t[^>]*>(.*?)</t>", si.group(1), re.S))
                    )
        return self._shared

    def xml(self, sheet: str) -> str:
        return self.z.read(self.sheets[sheet]).decode("utf8", "replace")

    def dimension(self, sheet: str) -> str:
        m = re.search(r'<dimension ref="([^"]+)"', self.xml(sheet))
        return m.group(1) if m else "?"

    def last_data_row(self, sheet: str) -> int:
        """L'ultima riga che contiene davvero un valore o una formula.

        Non `<dimension>`: quell'attributo e' un promemoria che Excel scrive e
        non sempre restringe. In `SF_DATABASE` dice `A1:EM3568` mentre l'ultima
        riga con un valore e' la 3389 — e il confronto con l'intervallo della
        tabella `AHT_Data` (`A1:EM3389`) stampava un ATTENZIONE per 179 righe di
        dati che non esistono. Un falso allarme in uno strumento che serve a
        trovare quelli veri e' peggio di nessun allarme.
        """
        raw = self.xml(sheet)
        ultima = 0
        for m in re.finditer(r'<row r="(\d+)"', raw):
            riga = int(m.group(1))
            if riga <= ultima:
                continue
            fine = raw.find("</row>", m.end())
            corpo = raw[m.end() : fine if fine != -1 else None]
            if "<v>" in corpo or "<is>" in corpo or "<f" in corpo:
                ultima = riga
        return ultima

    def header_row(self, sheet: str, row: int = 1) -> dict[str, str]:
        """Le intestazioni di una riga: `{lettera colonna: testo}`.

        `row` non e' sempre 1. `DUP_DATASET` e' il report Salesforce formattato
        incollato tale e quale, e le sue intestazioni stanno in riga 14, sotto
        titolo, `As of ...` e il blocco `Filtered By`.
        """
        raw = self.xml(sheet)
        m = re.search(rf'<row r="{row}"[^>]*>(.*?)</row>', raw, re.S)
        if not m:
            return {}
        out: dict[str, str] = {}
        for col, _rownum, attrs, body in _CELL.findall(m.group(1)):
            t = re.search(r't="([^"]+)"', attrs)
            v = re.search(r"<v>(.*?)</v>", body, re.S)
            inline = re.search(r"<is>.*?<t[^>]*>(.*?)</t>", body, re.S)
            if t and t.group(1) == "s" and v:
                val = self.shared[int(v.group(1))]
            elif inline:
                val = inline.group(1)
            elif v:
                val = v.group(1)
            else:
                continue
            out[col] = _unescape(val)
        return out

    def table_columns(self) -> dict[str, tuple[str, dict[str, str]]]:
        """`AHT_Data` -> (foglio, {nome colonna: lettera}).

        Serve a risolvere i riferimenti strutturati: nelle formule una colonna di
        tabella si scrive `AHT_Data[Case Type]`, non `SF_DATABASE!$Z`.
        """
        out: dict[str, tuple[str, dict[str, str]]] = {}
        owner: dict[str, str] = {}
        for name, target in self.sheets.items():
            rel = target.replace("worksheets/", "worksheets/_rels/") + ".rels"
            if rel not in self.z.namelist():
                continue
            x = self.z.read(rel).decode("utf8", "replace")
            for t in re.findall(r'Target="([^"]*tables/[^"]+)"', x):
                owner[Path(t).name] = name
        for n in sorted(p for p in self.z.namelist() if "/tables/" in p):
            x = self.z.read(n).decode("utf8", "replace")
            m = re.search(r'<table[^>]*name="([^"]+)"[^>]*ref="([^"]+)"', x)
            if not m:
                continue
            prima = re.match(r"([A-Z]+)", m.group(2))
            base = col_index(prima.group(1)) if prima else 1
            nomi = re.findall(r"<tableColumn[^>]*name=\"([^\"]+)\"", x)
            lettere = {
                _unescape(nome): col_letters(base + i) for i, nome in enumerate(nomi)
            }
            out[m.group(1)] = (owner.get(Path(n).name, "?"), lettere)
        return out

    def formula_usage(self, datasets: list[str]) -> dict[str, dict[str, set[str]]]:
        usage: dict[str, dict[str, set[str]]] = {d: defaultdict(set) for d in datasets}
        # Tre insidie, tutte scoperte col sangue su questo workbook:
        #
        # 1. I nomi di foglio con spazi, nelle formule, sono fra apici:
        #    `'Slot Only Cases'!$A$2`. Un pattern che pretende `Cases!` non li
        #    vede, e il foglio risulta "letto da nessuna formula" — falso.
        # 2. Un nome di foglio puo' essere sottostringa di un altro:
        #    `AT_DATASET` sta dentro `PSAT_DATASET`, `Turni` dentro
        #    `Helper Turni`. Serve un confine a sinistra che escluda anche uno
        #    spazio, altrimenti si attribuiscono a uno le colonne dell'altro.
        # 3. Una colonna dentro un ListObject si cita per NOME, non per lettera:
        #    `AHT_Data[Case Type]`. Cercando solo `SF_DATABASE!$Z` quelle
        #    colonne risultano lette da nessuno — e infatti due colonne vere
        #    (`Case Type`, `Primary Category`) sono rimaste fuori dal contratto,
        #    e il foglio 'AHT Outliers' e' uscito vuoto: le sue formule
        #    filtravano una colonna che la pipeline non riempiva.
        patterns = {
            d: re.compile(
                r"(?<![A-Za-z0-9_ ])" + re.escape(d) + r"'?!\$?([A-Z]{1,3})\$?\d*"
            )
            for d in datasets
        }
        tabelle = self.table_columns()
        strutturati = {
            nome: (foglio, lettere)
            for nome, (foglio, lettere) in tabelle.items()
            if foglio in datasets
        }
        for name in self.sheets:
            for f in formule(self.xml(name)):
                for d, pat in patterns.items():
                    for m in pat.finditer(f):
                        usage[d][m.group(1)].add(name)
                for tab, (foglio, lettere) in strutturati.items():
                    for m in re.finditer(re.escape(tab) + r"\[([^\]\[]+)\]", f):
                        col = m.group(1)
                        if col in lettere:
                            usage[foglio][lettere[col]].add(name)
        return usage

    def tables(self) -> list[dict]:
        out = []
        owner: dict[str, str] = {}
        for name, target in self.sheets.items():
            rel = target.replace("worksheets/", "worksheets/_rels/") + ".rels"
            if rel not in self.z.namelist():
                continue
            x = self.z.read(rel).decode("utf8", "replace")
            for t in re.findall(r'Target="([^"]*tables/[^"]+)"', x):
                owner[Path(t).name] = name
        for n in sorted(p for p in self.z.namelist() if "/tables/" in p):
            x = self.z.read(n).decode("utf8", "replace")
            m = re.search(r'<table[^>]*name="([^"]+)"[^>]*ref="([^"]+)"', x)
            if not m:
                continue
            sheet = owner.get(Path(n).name, "?")
            out.append(
                {
                    "name": m.group(1),
                    "ref": m.group(2),
                    "sheet": sheet,
                    "columns": len(re.findall(r"<tableColumn[^>]*name=", x)),
                    "sheet_dimension": self.dimension(sheet) if sheet in self.sheets else "?",
                    "sheet_last_data_row": (
                        self.last_data_row(sheet) if sheet in self.sheets else 0
                    ),
                }
            )
        return out


def _unescape(s: str) -> str:
    return (
        s.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&apos;", "'")
        .replace("&amp;", "&")
    )


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("workbook", type=Path)
    ap.add_argument("--headers", action="store_true")
    ap.add_argument("--usage", action="store_true")
    ap.add_argument("--tables", action="store_true")
    ap.add_argument("--all", action="store_true", help="tutti i controlli")
    ap.add_argument("--sheets", nargs="*", help="limita ai fogli indicati")
    ap.add_argument("--datasets", nargs="*", default=DEFAULT_DATASETS)
    ap.add_argument(
        "--header-row", nargs="*", metavar="FOGLIO=N", default=[],
        help="riga delle intestazioni per un foglio (es. DUP_DATASET=14). "
             "Di norma non serve: si legge da config/columns.yml.",
    )
    ap.add_argument("--json", action="store_true", help="output JSON (per le fixture)")
    args = ap.parse_args(argv)

    righe = header_rows()
    for voce in args.header_row:
        foglio, _, n = voce.rpartition("=")
        if not foglio or not n.isdigit():
            print(f"--header-row: attesa la forma FOGLIO=N, non {voce!r}", file=sys.stderr)
            return 2
        righe[foglio] = int(n)

    if not args.workbook.is_file():
        print(f"File non trovato: {args.workbook}", file=sys.stderr)
        return 2
    if not (args.headers or args.usage or args.tables or args.all):
        args.all = True

    wb = Workbook(args.workbook)
    payload: dict = {"workbook": args.workbook.name}

    if args.headers or args.all:
        targets = args.sheets or args.datasets
        payload["headers"] = {}
        for s in targets:
            if s not in wb.sheets:
                print(f"[skip] foglio assente: {s}", file=sys.stderr)
                continue
            riga = righe.get(s, 1)
            payload["headers"][s] = {
                "dimension": wb.dimension(s),
                "hidden": s in wb.hidden,
                "header_row": riga,
                "columns": wb.header_row(s, riga),
            }

    if args.usage or args.all:
        usage = wb.formula_usage(args.datasets)
        payload["usage"] = {
            d: {c: sorted(v) for c, v in sorted(cols.items(), key=lambda kv: col_index(kv[0]))}
            for d, cols in usage.items()
        }

    if args.tables or args.all:
        payload["tables"] = wb.tables()

    if args.json:
        print(json.dumps(payload, indent=2, ensure_ascii=False))
        return 0

    print(f"Workbook: {args.workbook}  ({len(wb.sheets)} fogli, {len(wb.hidden)} nascosti)")

    for sheet, info in payload.get("headers", {}).items():
        riga = f", riga {info['header_row']}" if info.get("header_row", 1) != 1 else ""
        print(f"\n=== INTESTAZIONI {sheet}  (dim {info['dimension']}{riga}) ===")
        for col, val in sorted(info["columns"].items(), key=lambda kv: col_index(kv[0])):
            print(f"  {col:>3} | {val}")

    for ds, cols in payload.get("usage", {}).items():
        print(f"\n=== COLONNE DI {ds} CONSUMATE DA FORMULE ===")
        if not cols:
            print("  (nessuna)")
        for col, sheets in cols.items():
            print(f"  {col:>3} <- {', '.join(sheets)}")

    if "tables" in payload:
        print("\n=== TABELLE (ListObject) ===")
        for t in payload["tables"]:
            note = ""
            m = re.search(r"[A-Z]+(\d+)$", t["ref"])
            # Sulle righe vere, non su `<dimension>`: vedi `last_data_row`.
            ultima = t.get("sheet_last_data_row") or 0
            if m and ultima > int(m.group(1)):
                note = (
                    f"  <-- ATTENZIONE: i dati del foglio arrivano a riga {ultima}, "
                    f"la tabella si ferma a {m.group(1)}"
                )
            print(f"  {t['name']} su '{t['sheet']}' ref={t['ref']} ({t['columns']} col.){note}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
