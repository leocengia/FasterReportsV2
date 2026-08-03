"""Adattatori per le sorgenti WFM: da matrice larga a righe tidy.

I 4 CSV dell'Omni Report sono già tabelle: una riga per record. `Turni` e
`Slot Only Cases` invece nascono da fogli **larghi** — un agente per riga, una
colonna per giorno — e vanno riportati in forma lunga prima di poter usare il
resto della pipeline.

Questi lettori restituiscono la stessa coppia `(headers, rows)` di
`csvsource.read_csv`. Da lì in poi matcher, coercizioni, preflight e writer non
cambiano: è la stessa astrazione prevista per il futuro SQL.

Le colonne si trovano **per nome di intestazione**, mai per posizione — vale qui
come per i CSV. Nel roster è ancora più necessario: il foglio contiene due
blocchi affiancati con popolazioni diverse, e le colonne-data sono 494.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from .coerce import EXCEL_EPOCH, to_datetime
from .errors import SourceError
from .names import full_name, has_marker, normalize_name, normalize_skill
from .shifts import (
    KIND_ASSENTE,
    KIND_LAVORA,
    KIND_SLOT,
    KIND_SLOT_NO_BOT,
    ShiftParseError,
    parse_shift,
    parse_slot,
)
from .xlsxsource import Sheet, col_to_index, index_to_col, read_sheet

# ---------------------------------------------------------------------------
# Roster turni
# ---------------------------------------------------------------------------

ROSTER_SHEET = "publish"
# Intestazioni fisse che aprono ogni blocco. `Agent` è il segnaposto di inizio.
ROSTER_KEYS = ("Agent", "Expected hours", "Zona", "Skill", "Name", "Surname")

# Intestazioni emesse in uscita: coincidono con i nomi canonici del contratto,
# così il matcher le aggancia per nome esatto.
ROSTER_HEADERS = (
    "Nome agente",
    "Team/Skill",
    "Contratto",
    "Ore/gg",
    "Data",
    "Stato",
    "Inizio turno",
    "Fine turno",
    "chiave",
)

BACKOFFICE_SHEET = "Only Cases Shifts"
BACKOFFICE_HEADERS = (
    "chiave (nome BO norm.)",
    "Data",
    "Slot inizio",
    "Slot fine",
    "Stato BO",
)
# Valori della colonna A che introducono righe non-agente (totali, etichette).
BACKOFFICE_NON_AGENT = (
    "total bo hrs",
    "optimized psp bo hrs",
    "psp bo hrs",
    "bo hrs variation",
)

_DATE_TEXT = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})$")
_SERIAL = re.compile(r"^\d+(\.0+)?$")


@dataclass
class SourceNotes:
    """Cosa il lettore ha visto e non ha buttato via.

    Non sono errori: sono fatti che il preflight deve riportare perché
    nasconderli è il modo in cui un report smette di corrispondere alla realtà.
    """

    skills_seen: dict[str, int] = field(default_factory=dict)
    skills_with_marker: dict[str, int] = field(default_factory=dict)
    shift_markers: dict[str, int] = field(default_factory=dict)
    non_shift_states: dict[str, list[str]] = field(default_factory=dict)
    unscheduled: list[str] = field(default_factory=list)
    blocks: list[str] = field(default_factory=list)
    rows_skipped: list[str] = field(default_factory=list)
    dates_available: list[date] = field(default_factory=list)
    slot_requests: list[str] = field(default_factory=list)
    stale_cache: bool = False
    # chiave -> skill, per tutti gli agenti la cui skill somiglia a una di
    # quelle richieste (esatta o marcata). Serve a filtrare il back office e ai
    # controlli di coerenza.
    target_agents: dict[str, str] = field(default_factory=dict)
    # chiave -> blocchi in cui l'agente compare con una skill richiesta.
    # Se un agente fosse "HPO" in due blocchi produrrebbe righe doppie.
    agents_in_blocks: dict[str, list[str]] = field(default_factory=dict)
    aliases_applied: dict[str, str] = field(default_factory=dict)
    keys_not_allowed: dict[str, int] = field(default_factory=dict)

    def _bump(self, d: dict[str, int], key: str) -> None:
        d[key] = d.get(key, 0) + 1


@dataclass
class TidySource:
    """Sostituto di `CsvSource` per le sorgenti larghe."""

    path: Path
    headers: list[str]
    data: list[list]
    notes: SourceNotes
    encoding: str = "xlsx"
    delimiter: str = "(matrice)"

    def rows(self):
        return iter(self.data)

    @property
    def n_cols(self) -> int:
        return len(self.headers)


def parse_header_date(value) -> date | None:
    """Data di un'intestazione di colonna, come testo `gg/mm/aaaa` o seriale."""
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    m = _DATE_TEXT.match(s)
    if m:
        d, mo, y = (int(g) for g in m.groups())
        try:
            return date(y, mo, d)
        except ValueError:
            return None
    if _SERIAL.match(s):
        n = int(float(s))
        # Sotto 40000 (~2009) non è una data di questo dominio: è un conteggio.
        if 40000 <= n <= 60000:
            return (EXCEL_EPOCH.date().toordinal() + n) and _from_serial(n)
    return None


def _from_serial(n: int) -> date:
    from datetime import timedelta

    return (EXCEL_EPOCH + timedelta(days=n)).date()


def _to_serial(d: date) -> int:
    return (datetime(d.year, d.month, d.day) - EXCEL_EPOCH).days


def find_header_row(sheet: Sheet, keys: tuple[str, ...], limit: int = 12) -> int:
    """Riga delle intestazioni: la prima che contiene almeno metà delle chiavi."""
    wanted = {k.lower() for k in keys}
    best, best_hits = 0, 0
    for r in range(1, limit + 1):
        vals = {str(v).strip().lower() for v in sheet.row(r).values()}
        hits = len(wanted & vals)
        if hits > best_hits:
            best, best_hits = r, hits
    if best_hits < max(2, len(keys) // 2):
        raise SourceError(
            f"{sheet.name}: riga di intestazione non riconosciuta nelle prime "
            f"{limit} righe.\n"
            f"  Cercavo almeno {max(2, len(keys)//2)} di: {', '.join(keys)}\n"
            f"  Trovate al massimo {best_hits} in una riga."
        )
    return best


@dataclass
class RosterBlock:
    fixed: dict[str, str]  # nome intestazione -> lettera colonna
    dates: list[tuple[str, date]]  # (lettera, data)
    start_index: int

    @property
    def label(self) -> str:
        cols = sorted(self.fixed.values(), key=col_to_index)
        span = f"{cols[0]}–{cols[-1]}" if cols else "?"
        if self.dates:
            span += f", date {self.dates[0][0]}–{self.dates[-1][0]}"
        return span


def find_roster_blocks(sheet: Sheet, header_row: int) -> list[RosterBlock]:
    """Individua i blocchi affiancati.

    Un blocco è un gruppo di colonne fisse consecutive (le `ROSTER_KEYS`,
    riconosciute **per nome**) seguito dalle sue colonne-giorno. Dove ricomincia
    un gruppo di colonne fisse, inizia il blocco successivo.

    Non si assume che `Agent` sia la colonna più a sinistra del blocco: se
    l'export riordina `Skill` e `Name`, il blocco va riconosciuto comunque —
    è lo stesso principio del contratto colonne, per nome e non per posizione.
    """
    hdr = sheet.row(header_row)
    kinds: dict[str, tuple[str, object]] = {}
    for c in sorted(hdr, key=col_to_index):
        raw = str(hdr[c]).strip()
        key = next((k for k in ROSTER_KEYS if k.lower() == raw.lower()), None)
        if key:
            kinds[c] = ("fixed", key)
            continue
        d = parse_header_date(hdr[c])
        if d:
            kinds[c] = ("date", d)

    ordered = sorted(kinds, key=col_to_index)
    if not any(k[0] == "fixed" for k in kinds.values()):
        raise SourceError(
            f"{sheet.name}: nessuna colonna del roster riconosciuta nella riga "
            f"{header_row}.\n"
            f"  Cercavo: {', '.join(ROSTER_KEYS)}\n"
            f"  Trovate: {', '.join(repr(str(hdr[c])) for c in ordered[:12])}"
        )

    blocks: list[RosterBlock] = []
    i = 0
    while i < len(ordered):
        if kinds[ordered[i]][0] != "fixed":
            i += 1
            continue
        fixed: dict[str, str] = {}
        first = ordered[i]
        while i < len(ordered) and kinds[ordered[i]][0] == "fixed":
            fixed.setdefault(kinds[ordered[i]][1], ordered[i])
            i += 1
        dates: list[tuple[str, date]] = []
        while i < len(ordered) and kinds[ordered[i]][0] == "date":
            dates.append((ordered[i], kinds[ordered[i]][1]))
            i += 1

        missing = [k for k in ROSTER_KEYS if k not in fixed]
        if missing:
            raise SourceError(
                f"{sheet.name}: il blocco che inizia in {first} non ha le colonne "
                f"{', '.join(missing)}.\n"
                f"  Intestazioni trovate nel blocco: {', '.join(sorted(fixed))}\n"
                f"  Senza di esse non si può costruire ne' il nome dell'agente ne' "
                f"il filtro sulla skill."
            )
        blocks.append(
            RosterBlock(
                fixed=fixed,
                dates=dates,
                start_index=col_to_index(min(fixed.values(), key=col_to_index)),
            )
        )
    return blocks


def read_roster(
    path: str | Path,
    *,
    week: tuple[date, date] | None = None,
    skills: tuple[str, ...] = ("HPO",),
    include_marked: bool = False,
    contratti: dict[str, str] | None = None,
    sheet_name: str | None = None,
) -> TidySource:
    """Roster largo -> una riga per (agente, giorno).

    `skills` filtra per **uguaglianza esatta**, come fa il FILTER di
    `Helper Turni` (`Turni!$B="HPO"`): è il comportamento fedele a oggi.

    `include_marked=False` è quindi il default, ma le skill che somigliano a una
    richiesta senza esserlo — `HPO␣␣␣*`, `hpo` — non vengono ignorate in
    silenzio: finiscono in `notes.skills_with_marker` e il controllo di coerenza
    **blocca**, perché è lì che si nasconde l'agente che oggi entra in metà
    delle regole e non nell'altra metà (vedi docs/audit-workbook-W30.md).
    Con `include_marked=True` quegli agenti entrano — cambia i numeri, quindi è
    una scelta da fare consapevolmente.

    In `notes.target_agents` finiscono **tutti** gli agenti con skill richiesta,
    marcati compresi: serve per filtrare il back office, che nel W30 include
    anche il marcato.
    """
    path = Path(path)
    sheet = read_sheet(path, sheet_name or _pick_sheet(path, ROSTER_SHEET))
    notes = SourceNotes(stale_cache=bool(sheet.n_formulas and not sheet.calc_chain_present))
    header_row = find_header_row(sheet, ROSTER_KEYS)
    blocks = find_roster_blocks(sheet, header_row)
    contratti = contratti or {}
    wanted = {s for s in skills}
    wanted_norm = {normalize_skill(s) for s in skills}

    all_dates: set[date] = set()
    out: list[list] = []

    for block in blocks:
        notes.blocks.append(block.label)
        in_week = [
            (c, d) for c, d in block.dates
            if week is None or (week[0] <= d <= week[1])
        ]
        all_dates.update(d for _, d in block.dates)

        for rownum in sorted(sheet.rows):
            if rownum <= header_row:
                continue
            row = sheet.row(rownum)
            skill_raw = row.get(block.fixed["Skill"], "")
            nome = full_name(row.get(block.fixed["Name"]), row.get(block.fixed["Surname"]))
            if not nome and not str(skill_raw).strip():
                continue

            skill = str(skill_raw).strip()
            notes._bump(notes.skills_seen, skill or "(vuota)")
            is_marked = has_marker(skill)
            if is_marked:
                notes._bump(notes.skills_with_marker, skill)

            # Somiglia a una skill richiesta? (esatta oppure marcata)
            near = normalize_skill(skill) in wanted_norm
            if near and nome:
                notes.target_agents.setdefault(normalize_name(nome), skill)
                notes.agents_in_blocks.setdefault(normalize_name(nome), []).append(
                    block.label
                )

            accepted = skill in wanted or (include_marked and near)
            if not accepted:
                continue
            if not nome:
                notes.rows_skipped.append(
                    f"{block.fixed['Name']}{rownum}: skill {skill!r} ma nome vuoto"
                )
                continue

            for col, d in in_week:
                cell = row.get(col)
                where = f"{col}{rownum} ({nome}, {d.isoformat()})"
                shift = parse_shift(cell, where=where)
                for mk in shift.markers:
                    notes._bump(notes.shift_markers, mk)

                if shift.kind == KIND_ASSENTE:
                    notes.unscheduled.append(where)
                    continue
                stato = shift.stato
                if stato is None:
                    notes.non_shift_states.setdefault(shift.kind, []).append(where)
                    continue

                out.append([
                    nome,
                    skill,
                    contratti.get(normalize_name(nome)),
                    shift.net_hours if shift.kind == KIND_LAVORA else None,
                    _to_serial(d),
                    stato,
                    shift.start,
                    shift.end,
                    normalize_name(nome) if shift.kind == KIND_LAVORA else None,
                ])

    notes.dates_available = sorted(all_dates)
    out.sort(key=lambda r: (normalize_name(r[0]), r[4]))
    return TidySource(path=path, headers=list(ROSTER_HEADERS), data=out, notes=notes)


# ---------------------------------------------------------------------------
# Back office / slot only cases
# ---------------------------------------------------------------------------

def read_alias_map(path: str | Path, sheet_name: str = "Helper Malpractice") -> dict[str, str]:
    """Tabella alias nomi da `Helper Malpractice!D:E` del workbook.

    Esiste perché il file back-office scrive alcuni nomi diversamente dal
    roster: `alessandro passierello` vs `passariello`, `asia chirrullo` vs
    `chirullo`, `kaotar garoui` vs `garaoui`, `victoriia lavrinets` vs
    `viktoriia`. Senza la traduzione la chiave dello slot non combacia con
    l'agente e il VBA non trova le sue finestre di back office: la regola
    "Available Cases fuori turno" non scatterebbe mai per lui, in silenzio.

    Si legge **dal workbook**, non da una copia in config: è la stessa tabella
    che usa `CreaMalpractice.LoadSlots`, e due copie divergono.
    """
    sheet = read_sheet(path, sheet_name)
    out: dict[str, str] = {}
    for rownum in sorted(sheet.rows):
        src, dst = sheet.cell("D", rownum), sheet.cell("E", rownum)
        if not src or not dst:
            continue
        k, v = normalize_name(src), normalize_name(dst)
        # La riga di intestazione ('Nome in Slot' -> 'Nome standard') si scarta
        # da sé: il VBA parte da riga 3, qui si riconosce che non è un alias
        # perché sorgente e destinazione non sono nomi di persona. Più semplice:
        # si salta l'identità e la riga 1/2.
        if rownum < 3 or not k or k == v:
            continue
        out[k] = v
    return out


def read_backoffice(
    path: str | Path,
    *,
    week: tuple[date, date] | None = None,
    sections: tuple[str, ...] | None = None,
    aliases: dict[str, str] | None = None,
    allowed_keys: set[str] | None = None,
    sheet_name: str | None = None,
) -> TidySource:
    """Foglio back-office largo -> una riga per (agente, giorno).

    Il foglio ha **sezioni** (`HPO`, `Part-Time 6h`, …) e righe di totale
    (`Total BO Hrs`, `PSP BO Hrs`, …), quindi non è una tabella piatta: la
    colonna A introduce le sezioni e le righe successive ereditano quella in
    corso. Le righe di totale e le etichette si saltano.

    `sections=None` prende tutte le sezioni di agenti: le sezioni part-time
    contengono agenti veri, non solo totali.

    `allowed_keys` limita alle chiavi passate — nella pratica gli agenti con
    skill richiesta nel roster. Serve perché il foglio back-office contiene
    anche agenti di altri team: nel W30 il risultato ha 37 agenti (i 36 `HPO`
    esatti **più** quello marcato `HPO␣␣␣*`), mentre il file ne ha 43.

    `aliases` traduce i nomi scritti diversamente fra le due sorgenti, come fa
    `CreaMalpractice.LoadSlots`. Vedi `read_alias_map`.
    """
    path = Path(path)
    sheet = read_sheet(path, sheet_name or _pick_sheet(path, BACKOFFICE_SHEET))
    notes = SourceNotes(stale_cache=bool(sheet.n_formulas and not sheet.calc_chain_present))

    header_row = _find_backoffice_header(sheet)
    hdr = sheet.row(header_row)
    datecols: list[tuple[str, date]] = []
    for c in sorted(hdr, key=col_to_index):
        d = parse_header_date(hdr[c])
        if d:
            datecols.append((c, d))
    if not datecols:
        raise SourceError(
            f"{sheet.name}: nessuna colonna-data nella riga {header_row}. "
            f"Attese date come seriali Excel."
        )
    notes.dates_available = sorted(d for _, d in datecols)
    in_week = [(c, d) for c, d in datecols if week is None or (week[0] <= d <= week[1])]

    # Le colonne fisse sono quelle prima della prima data: sezione, nome,
    # cognome. Vanno prese **per indice**, non fra le intestazioni popolate:
    # nel file reale `C1` è vuota (il cognome non ha intestazione), quindi
    # cercarle fra gli header ne troverebbe solo due.
    first_date_idx = col_to_index(datecols[0][0])
    if first_date_idx < 4:
        raise SourceError(
            f"{sheet.name}: la prima colonna-data è {datecols[0][0]}, ma servono "
            f"almeno 3 colonne prima (sezione, nome, cognome)."
        )
    col_section, col_first, col_last = (index_to_col(i) for i in (1, 2, 3))

    out: list[list] = []
    current_section = ""
    for rownum in sorted(sheet.rows):
        if rownum <= header_row:
            continue
        row = sheet.row(rownum)
        sec = str(row.get(col_section, "")).strip()
        if sec:
            current_section = sec
        low = current_section.lower()
        if low in BACKOFFICE_NON_AGENT or sec.lower() in BACKOFFICE_NON_AGENT:
            notes.rows_skipped.append(f"r{rownum}: sezione {current_section!r} (totali)")
            continue

        nome = full_name(row.get(col_first), row.get(col_last))
        if not nome:
            continue
        # Righe di servizio dentro una sezione agenti (es. 'intervals').
        if not str(row.get(col_first, "")).strip() or not str(row.get(col_last, "")).strip():
            notes.rows_skipped.append(f"r{rownum}: nome incompleto ({nome!r})")
            continue
        if _looks_like_label(nome):
            notes.rows_skipped.append(f"r{rownum}: etichetta {nome!r}")
            continue

        if sections is not None and current_section not in sections:
            continue

        key = _resolve_key(normalize_name(nome), aliases, allowed_keys, notes)
        if allowed_keys is not None and key not in allowed_keys:
            notes._bump(notes.keys_not_allowed, key)
            continue

        for col, d in in_week:
            cell = row.get(col)
            where = f"{col}{rownum} ({nome}, {d.isoformat()})"
            try:
                slot = parse_slot(cell, where=where)
            except ShiftParseError:
                notes.rows_skipped.append(f"{where}: valore non interpretabile, riga ignorata")
                break
            if slot.kind == KIND_SLOT:
                out.append([key, _to_serial(d), slot.start, slot.end, slot.stato_bo])
            elif slot.kind == KIND_SLOT_NO_BOT:
                out.append([key, _to_serial(d), None, None, slot.stato_bo])
            else:
                if slot.kind == "request":
                    notes.slot_requests.append(where)
                out.append([key, _to_serial(d), None, None, None])

    out.sort(key=lambda r: (r[0], r[1]))
    return TidySource(path=path, headers=list(BACKOFFICE_HEADERS), data=out, notes=notes)


def _resolve_key(
    key: str,
    aliases: dict[str, str] | None,
    allowed: set[str] | None,
    notes: SourceNotes,
) -> str:
    """Chiave dello slot nello stesso spazio di nomi del roster.

    La tabella `Helper Malpractice!D:E` mescola **due direzioni** — misurato sul
    W30:

      grafia back-office -> grafia roster   (4 voci: 'alessandro passierello'
          -> 'alessandro passariello', 'asia chirrullo' -> 'asia chirullo',
          'kaotar garoui' -> 'kaotar garaoui', 'victoriia' -> 'viktoriia')
      grafia roster -> grafia Salesforce    (3 voci: 'eleonora rosa sissa' ->
          'eleonora sissa', 'glenda martina medola' -> 'glenda medola',
          'nadia ariefieva' -> 'nadiia ariefieva')

    Il foglio `Slot Only Cases` usa la grafia del **roster** (verificato: le sue
    37 chiavi coincidono esattamente con gli agenti HPO del roster). Applicare
    l'alias alla cieca porterebbe quei 3 agenti *fuori* da quello spazio, e le
    loro righe sparirebbero.

    Quindi l'alias si usa per **raggiungere** lo spazio dei nomi del roster, non
    per lasciarlo: se la chiave è già valida si tiene, altrimenti si prova
    l'alias.
    """
    if allowed is None:
        # Senza insieme di riferimento non si può decidere: si applica l'alias
        # come fa il VBA.
        if aliases and key in aliases:
            notes.aliases_applied[key] = aliases[key]
            return aliases[key]
        return key
    if key in allowed:
        return key
    if aliases and key in aliases and aliases[key] in allowed:
        notes.aliases_applied[key] = aliases[key]
        return aliases[key]
    return key


def _looks_like_label(name: str) -> bool:
    """Righe come 'Agents in Only Cases per Interval' o 'intervals'."""
    low = name.lower()
    return low in ("intervals",) or len(name.split()) > 3


def _find_backoffice_header(sheet: Sheet) -> int:
    """La riga con più date: è l'intestazione del back office.

    Non una soglia fissa ma un massimo, così la funzione regge sia il file reale
    (150 colonne-giorno) sia un foglio ridotto.
    """
    best, best_n = 0, 0
    for r in range(1, 8):
        n = sum(1 for v in sheet.row(r).values() if parse_header_date(v))
        if n > best_n:
            best, best_n = r, n
    if best_n < 2:
        raise SourceError(
            f"{sheet.name}: riga di intestazione non trovata nelle prime 7 righe.\n"
            f"  Cercavo una riga con almeno 2 date (come seriali Excel o gg/mm/aaaa);\n"
            f"  al massimo ne ho trovate {best_n}."
        )
    return best


def _pick_sheet(path: Path, preferred: str) -> str:
    from .xlsxsource import read_sheet_names

    names = read_sheet_names(path)
    if preferred in names:
        return preferred
    return names[0]


def week_bounds(monday_serial: int) -> tuple[date, date]:
    """Intervallo lunedì–domenica a partire dal seriale del lunedì."""
    from datetime import timedelta

    start = _from_serial(monday_serial)
    return start, start + timedelta(days=6)


def week_from_iso(year: int, week: int) -> tuple[date, date]:
    """Lunedì–domenica della settimana ISO indicata.

    È così che si determina la settimana: dal numero che l'utente passa a
    `--week`, non dall'intervallo dei dati. Verificato sul W30: la ISO week 30
    del 2026 va dal 20 al 26 luglio, esattamente i sette giorni che il workbook
    ha in `Turni` e `Slot Only Cases`.
    """
    from datetime import timedelta

    try:
        monday = date.fromisocalendar(year, week, 1)
    except ValueError as exc:
        raise SourceError(
            f"Settimana ISO non valida: anno {year}, settimana {week} ({exc})."
        ) from None
    return monday, monday + timedelta(days=6)


def infer_year(values) -> int | None:
    """Anno prevalente fra le date passate.

    Si prende il piu' frequente e non il primo: gli export sbordano di qualche
    ora oltre i bordi della settimana, e a cavallo di capodanno il primo valore
    potrebbe essere dell'anno sbagliato.
    """
    from collections import Counter

    years: Counter[int] = Counter()
    for v in values:
        d = _as_datetime(v)
        if d:
            years[d.year] += 1
    if not years:
        return None
    return years.most_common(1)[0][0]


def week_from_dates(values) -> tuple[date, date] | None:
    """Intervallo coperto dalle date passate: min e max, così come sono.

    NON e' la settimana: l'export di `AT_DATASET` e' per data Seattle, e
    convertito in ora di Milano sborda oltre i sette giorni (nel W30 arriva al
    27 luglio con 3 righe). Serve solo per dire *quanto* i dati escono dalla
    settimana scelta, non per sceglierla.
    """
    seen: list[date] = []
    for v in values:
        d = _as_datetime(v)
        if d:
            seen.append(d.date())
    if not seen:
        return None
    return min(seen), max(seen)


def _as_datetime(value):
    """`to_datetime` che ignora cio' che non e' una data invece di sollevare.

    Qui si sta ragionando *su* un insieme di date per capire settimana e anno:
    un valore sporco va saltato, non deve far fallire il ragionamento. La
    validazione seria dei tipi l'ha gia' fatta `build_block`, che blocca se
    troppi valori di una colonna datetime non sono convertibili.
    """
    from .coerce import Uncoercible

    try:
        return to_datetime(value)
    except Uncoercible:
        return None
