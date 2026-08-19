"""Lettura di un dataset tabellare da un foglio Excel.

I 4 dataset dell'Omni Report non arrivano tutti come CSV: gli export reali sono
misti — `SF DATABASE W31.csv` e `PSAT DATASET W31.csv` da un lato,
`AT DATASET W31.xlsx` e `ATwi DATASET W31.xlsx` dall'altro. Il formato dipende da
chi produce l'export, non da noi.

Questo lettore fa per un foglio quello che `csvsource` fa per un CSV: trova la
riga di intestazione, restituisce gli header e le righe. Da lì in poi la pipeline
e' identica — matcher, coercizioni, preflight, writer.

Differenza da `wfmsource`: quello riduce una **matrice larga** (un agente per
riga, una colonna per giorno) a forma lunga. Qui il foglio e' gia' una tabella:
una riga per record, gli header in cima.

Le date arrivano come **seriali Excel** invece che come testo, e i decimali come
numeri veri invece che con virgola o punto: `coerce.to_datetime` e `to_float`
gestiscono entrambe le forme, quindi non serve convertire nulla qui. E' anzi il
caso piu' facile — un CSV con date in formato ambiguo e' piu' rischioso di un
foglio dove la data e' un numero.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterator

from .errors import SourceError
from .xlsxsource import col_to_index, read_sheet, read_sheet_names

# Estensioni riconosciute come "foglio Excel" invece che come testo delimitato.
EXCEL_SUFFIXES = (".xlsx", ".xlsm", ".xltx", ".xltm")


@dataclass
class TableSource:
    """Sostituto di `CsvSource` per un dataset che arriva come foglio Excel."""

    path: Path
    headers: list[str]
    data: list[list]
    encoding: str = "xlsx"
    delimiter: str = "(foglio)"
    sheet: str = ""
    # Le righe SOPRA l'intestazione, nell'ordine, ognuna larga quanto le
    # intestazioni. Un export normale non ne ha; il report Salesforce formattato
    # ne ha 13 (titolo, `As of <quando>`, il blocco `Filtered By`), e sono la
    # sola cosa che dice quale intervallo e' stato chiesto al report. Il writer le
    # ricopia nel foglio, cosi' `DUP_DATASET` non dichiara per sempre la
    # settimana in cui e' stato costruito il template.
    preamble: list[list] = field(default_factory=list)
    # In che riga del FILE stavano le intestazioni. Non e' `header_row` del
    # contratto — quella e' la riga del foglio di destinazione. Servono
    # entrambe, e confonderle vorrebbe dire spostare i dati quando il preambolo
    # dell'export cambia lunghezza.
    header_row: int = 1

    def rows(self) -> Iterator[list]:
        return iter(self.data)

    @property
    def n_cols(self) -> int:
        return len(self.headers)


def is_excel(path: str | Path) -> bool:
    return Path(path).suffix.lower() in EXCEL_SUFFIXES


# Fin dove cercare la riga di intestazione.
#
# Era 10, e non bastava: il report Salesforce formattato (`DUP_DATASET`) mette le
# intestazioni in **riga 14**, sotto il titolo, la riga `As of <quando>` e il
# blocco `Filtered By` — una riga per filtro. Con 10 il lettore rispondeva
# «nessuna riga di intestazione riconoscibile» su un file perfettamente valido.
#
# 25 e non 14: il numero di filtri del report puo' cambiare, e con lui la
# lunghezza del preambolo. Alzare il limite non rende l'euristica piu' incerta —
# quella distingue un titolo da un'intestazione per LARGHEZZA, e le righe di
# preambolo hanno una cella ciascuna contro le sedici della riga 14.
LIMITE_SCANSIONE = 25


def find_header_row(sheet, limit: int = LIMITE_SCANSIONE) -> int:
    """La **prima** riga che sembra una riga di intestazioni.

    Gli export mettono le intestazioni in riga 1, ma alcuni ci infilano sopra un
    titolo o una riga di filtri. Due segnali distinguono un titolo da un'intestazione:

    - **la larghezza**: un titolo occupa una cella o due, un'intestazione e' larga
      quanto i dati sotto di lei;
    - **il tipo**: un'intestazione e' quasi tutta testo, una riga di dati no.

    Si prende la **prima** riga che soddisfa entrambi, non quella col punteggio
    piu' alto: se l'intestazione ha un nome vuoto (capita) mentre le righe dati
    sono piene, la riga piu' "ricca" sarebbe un record, e si perderebbe la prima
    riga di dati leggendola come intestazione.
    """
    scansione: list[tuple[int, int, int]] = []  # (riga, celle, celle di testo)
    for r in range(1, limit + 1):
        cells = [str(v).strip() for v in sheet.row(r).values() if str(v).strip()]
        if not cells:
            continue
        testuali = sum(1 for c in cells if not _pare_numero(c))
        scansione.append((r, len(cells), testuali))

    if scansione:
        larga = max(n for _r, n, _t in scansione)
        for r, n, testuali in scansione:
            # Larga almeno la meta' della riga piu' larga: scarta i titoli.
            if n * 2 < larga:
                continue
            if testuali >= n * 0.6:
                return r

    raise SourceError(
        f"{sheet.name}: nessuna riga di intestazione riconoscibile nelle prime "
        f"{limit} righe.\n"
        f"  Cercavo una riga larga quanto i dati e in prevalenza testo.\n"
        f"  Il foglio e' vuoto, le intestazioni sono piu' in basso, oppure non "
        f"ci sono affatto.\n"
        f"  Si puo' indicare la riga a mano con header_row."
    )


def _pare_numero(s: str) -> bool:
    try:
        float(s.replace(",", "."))
        return True
    except ValueError:
        return False


def read_table(
    path: str | Path,
    *,
    sheet_name: str | None = None,
    header_row: int | None = None,
) -> TableSource:
    """Legge un dataset tabellare da un foglio.

    `sheet_name=None` prende il primo foglio: gli export hanno un foglio solo.
    Se ce ne sono piu' di uno e non si dice quale, si prende il primo e lo si
    dichiara nel preflight — meglio che indovinare in silenzio.
    """
    path = Path(path)
    names = read_sheet_names(path)
    if sheet_name is None:
        sheet_name = names[0] if names else None
    if sheet_name is None:
        raise SourceError(f"{path.name}: nessun foglio nel file.")

    sheet = read_sheet(path, sheet_name)
    if not sheet.rows:
        raise SourceError(f"{path.name}, foglio {sheet_name!r}: nessuna riga.")

    hrow = header_row or find_header_row(sheet)
    hdr = sheet.row(hrow)
    if not hdr:
        raise SourceError(
            f"{path.name}, foglio {sheet_name!r}: riga {hrow} vuota, "
            f"non contiene intestazioni."
        )

    # Le colonne nell'ordine del foglio, dalla prima all'ultima popolata
    # nell'intestazione. Le colonne senza intestazione restano nel blocco (come
    # in un CSV con un header vuoto) perche' il matcher lavora per nome: una
    # colonna senza nome semplicemente non verra' agganciata.
    cols = sorted(hdr, key=col_to_index)
    first, last = col_to_index(cols[0]), col_to_index(cols[-1])
    from .xlsxsource import index_to_col

    tutte = [index_to_col(i) for i in range(first, last + 1)]
    headers = [str(hdr.get(c, "")).strip() for c in tutte]

    # Il preambolo si tiene com'e', righe vuote comprese e alla loro posizione:
    # e' un blocco di testo da ricopiare, non una tabella da ripulire, e gli
    # spazi bianchi sono parte di come si legge. Per questo si scorre l'intervallo
    # 1..hrow-1 e non `sheet.rows`: una riga vuota non esiste nell'XML, e
    # saltarla farebbe scalare di uno tutte quelle sotto.
    preamble: list[list] = []
    for rownum in range(1, hrow):
        row = sheet.row(rownum) if rownum in sheet.rows else {}
        preamble.append([row.get(c) for c in tutte])

    data: list[list] = []
    for rownum in sorted(sheet.rows):
        if rownum <= hrow:
            continue
        row = sheet.row(rownum)
        valori = [row.get(c) for c in tutte]
        # Righe completamente vuote: coda del foglio, non dati.
        if not any(v is not None and str(v).strip() for v in valori):
            continue
        data.append(valori)

    return TableSource(
        path=path, headers=headers, data=data, sheet=sheet_name,
        preamble=preamble, header_row=hrow,
        delimiter=f"(foglio {sheet_name!r})" if len(names) > 1 else "(foglio)",
    )
