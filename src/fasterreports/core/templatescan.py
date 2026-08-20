"""I limiti di riga scritti a mano nelle formule del template.

Molte formule del workbook non leggono una colonna intera ma un intervallo con la
riga finale **scritta nella formula**:

    Report Agenti          -> SUMIFS(AT_DATASET!$P$2:$P$130000, ...)
    Anagrafica             -> FILTER(SF_DATABASE!$BB$2:$BB$50000, ...)
    Profilo Colonne SF     -> INDEX(SF_DATABASE!$A$2:$EM$3389, 0, n)

Finche' i dati stanno sotto quel numero non cambia niente. Il giorno in cui lo
superano, la formula smette di vedere le righe in eccesso — e **non c'e' nessun
errore**: escono medie su un sottoinsieme, conteggi piu' bassi, percentuali
plausibili. E' il guasto silenzioso di sempre, con la differenza che qui il
detonatore e' la crescita del volume: arriva da solo, senza che nessuno tocchi
niente.

Questo modulo li trova leggendo le formule dal file. Il confronto con le righe
davvero scritte lo fa `coherence._check_row_limits`, che ha i due numeri.

Non serve Excel: le formule stanno nell'XML dentro lo zip.
"""

from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass
from pathlib import Path

from .errors import SourceError

# `'Slot Only Cases'!$A$2:$E$1291` e `AT_DATASET!$P$2:$P$130000`.
# Il confine a sinistra serve perche' `AT_DATASET` e' sottostringa di
# `PSAT_DATASET` e `Turni` di `Helper Turni`.
_RANGE = r"(?<![A-Za-z0-9_ ])'?{}'?!\$?[A-Z]{{1,3}}\$?(\d+):\$?[A-Z]{{1,3}}\$?(\d+)"
_REF = r"(?<![A-Za-z0-9_ ])'?{}'?!(\$?[A-Z]{{1,3}}\$?\d+:\$?[A-Z]{{1,3}}\$?\d+)"
# Un riferimento a UNA cella: `DUP_DATASET!B15`. Serve solo ai dataset letti
# cella per cella — vedi `scan_cell_refs`.
_CELL_REF = r"(?<![A-Za-z0-9_ ])'?{}'?!\$?([A-Z]{{1,3}})\$?(\d+)(?![\d:])"

# Una formula nell'XML: `<f>SUM(A1:A9)</f>`, ma anche `<f t="shared" si="6"/>`,
# che nel corpo non ha testo. Un pattern `<f[^>]*>(.*?)</f>` considera la forma
# auto-chiusa un tag di APERTURA — il `[^>]*` si mangia anche lo slash — e
# cattura tutto fino al `</f>` successivo, incollando insieme il testo di celle
# diverse. Nel template ci sono 6993 formule condivise.
_FORMULA = re.compile(r"<f(?:[^>\"]|\"[^\"]*\")*?(?:/>|>(.*?)</f>)", re.S)


def _formule(raw: str) -> list[str]:
    """Il testo delle formule di un foglio, senza le condivise vuote."""
    return [_unescape(f) for f in _FORMULA.findall(raw) if f]


@dataclass(frozen=True)
class RowLimit:
    """Un intervallo che si fermerebbe se i dati crescessero."""

    dataset: str
    max_row: int
    sheet: str
    ref: str

    def __str__(self) -> str:
        return f"{self.sheet}: {self.dataset}!{self.ref} (fino a riga {self.max_row})"


def scan_row_limits(path: str | Path, datasets: tuple[str, ...]) -> list[RowLimit]:
    """Tutti gli intervalli con riga finale esplicita verso i fogli dati.

    Un riferimento a colonna intera (`AT_DATASET!$K:$K`) non ha limite e non
    viene riportato: e' la forma robusta, ed e' quella da preferire quando si
    scrive una formula nuova. Anche i riferimenti a tabella (`AHT_Data[...]`) non
    hanno limite, perche' la tabella viene ridimensionata ai dati a ogni build.
    """
    path = Path(path)
    if not path.is_file():
        raise SourceError(f"Template non trovato: {path}")
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        raise SourceError(f"{path.name}: non e' un file .xlsm/.xlsx valido.") from None

    out: list[RowLimit] = []
    visti: set[tuple[str, str, str]] = set()
    with z:
        for nome, target in _sheet_targets(z, path).items():
            try:
                raw = z.read(target).decode("utf8", "replace")
            except KeyError:
                continue
            for f in _formule(raw):
                for ds in datasets:
                    for m in re.finditer(_RANGE.format(re.escape(ds)), f):
                        fine = max(int(m.group(1)), int(m.group(2)))
                        ref_m = re.search(_REF.format(re.escape(ds)), f[m.start() :])
                        ref = ref_m.group(1).replace("$", "") if ref_m else f"?{fine}"
                        chiave = (nome, ds, ref)
                        if chiave in visti:
                            continue
                        visti.add(chiave)
                        out.append(
                            RowLimit(dataset=ds, max_row=fine, sheet=nome, ref=ref)
                        )
    return sorted(out, key=lambda l: (l.dataset, l.max_row, l.sheet, l.ref))


def scan_cell_refs(path: str | Path, datasets: tuple[str, ...]) -> list[RowLimit]:
    """Il limite dei dataset letti **cella per cella**, non per intervalli.

    `Duplicates Helper!A2` e' `IF(OR(DUP_DATASET!B15="",...),"",DUP_DATASET!B15)`,
    `A3` punta a `B16`, e cosi' via fino a `A1001` -> `B1014`. Non c'e' nessun
    intervallo da trovare: c'e' una griglia di 13 120 riferimenti puntuali, e la
    riga piu' alta fra quelli **e' il limite** — la riga 1015 di `DUP_DATASET`
    non la leggerebbe nessuno.

    PERCHE' SOLO PER I DATASET CHE LO DICHIARANO (`read_by_row` nel contratto), e
    non per tutti: `Recap PSAT Positive` punta alla riga **fissa** `PSAT_DATASET!DQ130`
    — l'"elogio della settimana", scelto a mano. Una scansione indiscriminata
    leggerebbe quel 130 come il limite di `PSAT_DATASET` e **bloccherebbe** ogni
    settimana con piu' di 130 risposte al sondaggio: un guasto inventato dal
    controllo, che e' il peggior tipo.

    Restituisce un `RowLimit` per (foglio, dataset, colonna) — la stessa forma di
    `scan_row_limits`, cosi' `_check_row_limits` non ha bisogno di sapere da quale
    delle due scansioni arriva il numero.
    """
    path = Path(path)
    if not path.is_file():
        raise SourceError(f"Template non trovato: {path}")
    if not datasets:
        return []
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        raise SourceError(f"{path.name}: non e' un file .xlsm/.xlsx valido.") from None

    # (foglio, dataset, colonna) -> riga massima
    massimi: dict[tuple[str, str, str], int] = {}
    with z:
        for nome, target in _sheet_targets(z, path).items():
            try:
                raw = z.read(target).decode("utf8", "replace")
            except KeyError:
                continue
            for f in _formule(raw):
                for ds in datasets:
                    if ds not in f:  # taglia corto: il 99% delle formule non lo cita
                        continue
                    for m in re.finditer(_CELL_REF.format(re.escape(ds)), f):
                        col, riga = m.group(1), int(m.group(2))
                        chiave = (nome, ds, col)
                        if riga > massimi.get(chiave, 0):
                            massimi[chiave] = riga

    out = [
        RowLimit(dataset=ds, max_row=riga, sheet=nome, ref=f"{col}{riga} (cella per cella)")
        for (nome, ds, col), riga in massimi.items()
    ]
    return sorted(out, key=lambda l: (l.dataset, l.max_row, l.sheet, l.ref))


def scan_formula_extent(
    path: str | Path, punti: tuple[tuple[str, str, int], ...]
) -> dict[tuple[str, str], int]:
    """Fin dove arrivano le formule di una colonna: `(foglio, colonna) -> riga`.

    Serve ai fogli che presentano un elenco con **una formula per riga**: l'elenco
    dei nomi e' un array dinamico e cresce da se', ma le colonne accanto (il
    conteggio, la media, la percentuale) hanno una formula scritta riga per riga e
    si fermano dove le ha tirate chi ha fatto il foglio. La voce in eccesso
    compare nell'elenco **senza nessun numero accanto**: presente e invisibile
    insieme, senza un solo errore. E' lo stesso difetto tolto a 'Helper CaseType'.

    `punti` sono `(foglio, colonna, prima riga)`; si restituisce l'ultima riga
    >= `prima riga` che in quella colonna ha una formula.

    La misura si legge dal TEMPLATE e non si scrive nel codice: se domani le
    formule vengono tirate piu' in basso, il controllo lo segue da solo — lo
    stesso principio dei formati numerici di 'AHT History'.
    """
    path = Path(path)
    if not path.is_file():
        raise SourceError(f"Template non trovato: {path}")
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        raise SourceError(f"{path.name}: non e' un file .xlsm/.xlsx valido.") from None

    volute: dict[str, list[tuple[str, int]]] = {}
    for foglio, col, prima in punti:
        volute.setdefault(foglio, []).append((col, prima))

    out: dict[tuple[str, str], int] = {}
    with z:
        fogli = _sheet_targets(z, path)
        for foglio, colonne in volute.items():
            target = fogli.get(foglio)
            if target is None:
                continue
            try:
                raw = z.read(target).decode("utf8", "replace")
            except KeyError:
                continue
            for col, prima in colonne:
                # Una cella con formula: `<c r="B40" ...><f ...` — anche quando la
                # formula e' condivisa e il tag e' auto-chiuso, il `<f` c'e'.
                pat = re.compile(rf'<c r="{col}(\d+)"[^>]*>\s*<f[^>]*[>/]')
                righe = [
                    int(m.group(1)) for m in pat.finditer(raw) if int(m.group(1)) >= prima
                ]
                if righe:
                    out[(foglio, col)] = max(righe)
    return out


@dataclass(frozen=True)
class ErrorCells:
    """Celle di errore trovate in un foglio, per tipo."""

    sheet: str
    kinds: dict[str, int]
    examples: tuple[str, ...]

    @property
    def total(self) -> int:
        return sum(self.kinds.values())

    def __str__(self) -> str:
        tipi = ", ".join(f"{n} {k}" for k, n in sorted(self.kinds.items()))
        return f"{self.sheet}: {tipi} (es. {', '.join(self.examples)})"


def scan_error_cells(path: str | Path, limit: int = 5) -> list[ErrorCells]:
    """Le celle che dopo il ricalcolo contengono un errore Excel.

    E' il controllo piu' economico che esista sul workbook finito, e intercetta
    una famiglia di guasti che nessun controllo sulle sorgenti puo' vedere:

    - `#SPILL!` — un array dinamico non ha spazio per espandersi. Arriva **da
      solo** quando gli agenti o i casi crescono: la formula di ieri stava, quella
      di oggi no.
    - `#REF!` — una formula legge lo spill di un'altra che e' rimasta vuota.
    - `#VALUE!` — tipicamente una MEDIAN/QUARTILE su un insieme vuoto.
    - `#N/D` — un XLOOKUP che non trova, spesso un nome che non fa match.

    Misurato sul primo W31: 116 celle di errore, tutte discendenti dallo stesso
    guasto (due colonne di SF_DATABASE non riempite). Nessuna era visibile
    guardando i fogli principali.
    """
    path = Path(path)
    if not path.is_file():
        raise SourceError(f"Workbook non trovato: {path}")
    try:
        z = zipfile.ZipFile(path)
    except zipfile.BadZipFile:
        raise SourceError(f"{path.name}: non e' un file .xlsm/.xlsx valido.") from None

    out: list[ErrorCells] = []
    with z:
        for nome, target in _sheet_targets(z, path).items():
            try:
                raw = z.read(target).decode("utf8", "replace")
            except KeyError:
                continue
            trovate = re.findall(
                r'<c r="([A-Z]+\d+)"[^>]*t="e"[^>]*>.*?<v>(.*?)</v>', raw, re.S
            )
            if not trovate:
                continue
            kinds: dict[str, int] = {}
            for _ref, val in trovate:
                v = _unescape(val)
                kinds[v] = kinds.get(v, 0) + 1
            out.append(
                ErrorCells(
                    sheet=nome,
                    kinds=kinds,
                    examples=tuple(ref for ref, _v in trovate[:limit]),
                )
            )
    return sorted(out, key=lambda e: -e.total)


def _sheet_targets(z: zipfile.ZipFile, path: Path) -> dict[str, str]:
    try:
        wb = z.read("xl/workbook.xml").decode("utf8", "replace")
        rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8", "replace")
    except KeyError:
        raise SourceError(f"{path.name}: struttura .xlsx inattesa.") from None
    relmap = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels))
    fogli: dict[str, str] = {}
    for m in re.finditer(r"<sheet ([^>]*?)/>", wb):
        nome = re.search(r'name="([^"]*)"', m.group(1))
        rid = re.search(r'r:id="(rId\d+)"', m.group(1))
        if not (nome and rid and rid.group(1) in relmap):
            continue
        target = relmap[rid.group(1)].lstrip("/")
        fogli[_unescape(nome.group(1))] = (
            target if target.startswith("xl/") else "xl/" + target
        )
    return fogli


def binding_limits(limits: list[RowLimit]) -> dict[str, RowLimit]:
    """Per ogni dataset, il limite piu' BASSO: e' quello che morde per primo."""
    out: dict[str, RowLimit] = {}
    for lim in limits:
        attuale = out.get(lim.dataset)
        if attuale is None or lim.max_row < attuale.max_row:
            out[lim.dataset] = lim
    return out


def _unescape(s: str) -> str:
    return (
        s.replace("&lt;", "<")
        .replace("&gt;", ">")
        .replace("&quot;", '"')
        .replace("&apos;", "'")
        .replace("&amp;", "&")
    )
