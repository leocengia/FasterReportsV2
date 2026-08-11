"""Preflight: leggere il report e capire cosa la pipeline ha deciso.

Viene scritto SEMPRE, anche quando tutto va bene: e' la tracciabilita' delle
mappature scelte. Se fra sei mesi un numero e' strano, il preflight di quella
settimana dice a quale colonna del CSV era agganciata ogni formula.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from .contract import Contract, Dataset, col_to_index
from .csvsource import CsvSource
from .errors import PipelineError
from .matcher import DatasetMapping, resolve_dataset
from .transform import Block


@dataclass
class DatasetReport:
    name: str
    source: str
    encoding: str = ""
    delimiter: str = ""
    n_source_cols: int = 0
    mapping: DatasetMapping | None = None
    block: Block | None = None
    error: str = ""
    # Osservazioni del lettore (solo sorgenti WFM): skill viste, marcatori,
    # righe saltate. Non sono errori, ma vanno riportate.
    notes: object | None = None
    # Popolato SOLO per un dataset `optional: true` la cui fonte non c'e'. A
    # differenza di `error`, non fa fallire il preflight: il foglio viene
    # comunque scritto, vuoto. Va comunque mostrato in chiaro — un'assenza
    # taciuta e' quanto di piu' lontano dal principio del progetto.
    skipped_reason: str = ""

    @property
    def ok(self) -> bool:
        return not self.error


@dataclass
class PreflightReport:
    datasets: list[DatasetReport] = field(default_factory=list)
    generated_at: str = ""
    coherence: object | None = None
    # La settimana ricavata dai dati e i sette giorni che copre. Va scritta
    # sempre, anche quando le fonti WFM mancano e la sezione coerenza non gira:
    # e' il numero su cui si ritagliano turni e slot, e chi legge il rapporto
    # deve poterlo confrontare con quello che si aspettava.
    week: tuple[int, int] | None = None
    week_bounds: tuple[object, object] | None = None

    @property
    def ok(self) -> bool:
        if any(not d.ok for d in self.datasets):
            return False
        if self.coherence is not None and not self.coherence.ok:
            return False
        return True

    @property
    def errors(self) -> list[str]:
        return [d.error for d in self.datasets if d.error]

    def render(self) -> str:
        lines: list[str] = []
        stamp = self.generated_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        lines.append("=" * 78)
        lines.append("PREFLIGHT — mappatura colonne CSV -> workbook")
        lines.append(f"Generato: {stamp}")
        lines.append(f"Esito complessivo: {'OK' if self.ok else 'BLOCCATO'}")
        if self.week:
            anno, num = self.week
            riga = f"Settimana dai dati: W{num:02d} {anno}"
            if self.week_bounds:
                lo, hi = self.week_bounds
                riga += f" ({_dmy(lo)} → {_dmy(hi)}, 7 giorni)"
            lines.append(riga)
        lines.append("=" * 78)

        for d in self.datasets:
            lines.append("")
            lines.append(f"### {d.name}")
            lines.append(f"  sorgente: {d.source}")
            if d.encoding:
                delim = {",": "virgola", ";": "punto e virgola", "\t": "tab", "|": "pipe"}.get(
                    d.delimiter, repr(d.delimiter)
                )
                lines.append(
                    f"  encoding: {d.encoding} · separatore: {delim} · "
                    f"colonne nel sorgente: {d.n_source_cols}"
                )
            if d.error:
                lines.append("  STATO: BLOCCATO")
                for line in d.error.splitlines():
                    lines.append(f"  | {line}")
                continue

            if d.skipped_reason:
                lines.append("  STATO: SALTATO (fonte opzionale, assente questa settimana)")
                for line in d.skipped_reason.splitlines():
                    lines.append(f"  | {line}")
                lines.append(
                    "  Il foglio viene scritto VUOTO: nessun dato di settimane "
                    "precedenti resta dentro."
                )
                continue

            lines.append("  STATO: OK")
            assert d.mapping is not None
            rows = [("campo canonico", "col.", "colonna sorgente", "via", "note")]
            # Per indice, non per lettera: in ordine alfabetico 'U' finirebbe
            # dopo 'EI' e la tabella non seguirebbe l'ordine del foglio.
            for r in sorted(d.mapping.resolutions, key=lambda r: col_to_index(r.target_col)):
                note = ""
                if r.shadowed:
                    note = "disambiguata da match esatto; scartate: " + ", ".join(
                        repr(s) for s in r.shadowed
                    )
                if d.block:
                    st = d.block.stats.get(r.canonical)
                    if st:
                        bits = []
                        if st.blank:
                            bits.append(f"{st.blank} vuote")
                        if st.bad:
                            bits.append(f"{st.bad} non convertibili")
                        if bits:
                            note = "; ".join(filter(None, [note, " / ".join(bits)]))
                rows.append((r.canonical, r.target_col, r.source_header, r.via, note))
            lines.extend(_table(rows, indent="  "))

            if d.block:
                lines.append(
                    f"  righe scritte: {d.block.n_rows} · "
                    f"intervallo colonne: {d.block.start_col}:{d.block.end_col}"
                )
            unused = d.mapping.unused_headers
            lines.append(f"  colonne sorgente non usate dal motore: {len(unused)}")
            lines.extend(_render_notes(d.notes))

        if self.coherence is not None:
            lines.append("")
            lines.append("### COERENZA FRA LE FONTI")
            lines.extend(self.coherence.render(indent="  "))

        if not self.ok:
            lines.append("")
            lines.append("-" * 78)
            lines.append("La pipeline si e' fermata: nessun workbook e' stato prodotto.")
            lines.append("Correggi i punti BLOCCATO qui sopra e rilancia.")
        return "\n".join(lines) + "\n"

    def write(self, path: str | Path) -> Path:
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(self.render(), encoding="utf-8")
        return path


def _dmy(d) -> str:
    """Data in giorno/mese/anno: e' come la legge chi usa il report."""
    return d.strftime("%d/%m/%Y") if hasattr(d, "strftime") else str(d)


def _render_notes(notes) -> list[str]:
    """Cosa il lettore WFM ha visto: conta perché la trasformazione scarta righe."""
    if notes is None:
        return []
    out: list[str] = []
    skipped = getattr(notes, "rows_skipped", None)
    if skipped:
        out.append(f"  righe sorgente saltate: {len(skipped)}")
        for s in skipped[:4]:
            out.append(f"      {s}")
        if len(skipped) > 4:
            out.append(f"      ... e altre {len(skipped) - 4}")
    unsched = getattr(notes, "unscheduled", None)
    if unsched:
        out.append(f"  celle senza schedulazione: {len(unsched)}")
    excluded = getattr(notes, "keys_not_allowed", None)
    if excluded:
        out.append(
            f"  agenti della sorgente esclusi (fuori dal target): "
            f"{', '.join(sorted(excluded))}"
        )
    applied = getattr(notes, "aliases_applied", None)
    if applied:
        out.append(f"  alias nomi applicati: {len(applied)}")
        for k, v in sorted(applied.items()):
            out.append(f"      {k!r} -> {v!r}")
    # Va detto, perche' e' l'unico posto in cui il valore scritto non e' quello
    # letto: senza la forma canonica il FILTER di 'Helper Turni' scarterebbe
    # l'agente, e includerlo in 'Turni' non servirebbe a niente.
    normalizzate = getattr(notes, "skills_normalized", None)
    if normalizzate:
        out.append("  skill riscritte alla forma canonica (per il FILTER esatto):")
        for k, v in sorted(normalizzate.items()):
            out.append(f"      {k}  ({v} righe)")
    return out


def _table(rows: list[tuple[str, ...]], indent: str = "") -> list[str]:
    """Tabella a larghezza fissa. L'ultima colonna non viene paddata."""
    if not rows:
        return []
    n = len(rows[0])
    widths = [max(len(str(r[i])) for r in rows) for i in range(n)]
    out = []
    for j, r in enumerate(rows):
        cells = [
            str(r[i]).ljust(widths[i]) if i < n - 1 else str(r[i])
            for i in range(n)
        ]
        out.append(indent + " | ".join(cells).rstrip())
        if j == 0:
            out.append(indent + "-+-".join("-" * w for w in widths))
    return out


def check_dataset(
    contract: Contract,
    dataset: Dataset,
    source: CsvSource,
) -> DatasetReport:
    """Risolve gli header di un dataset e cattura l'errore invece di propagarlo.

    Cattura, non ignora: si vuole il quadro completo di tutti i dataset in un
    solo run, e poi si blocca. Fermarsi al primo errore costringerebbe a
    quattro giri per sistemare quattro CSV.
    """
    report = DatasetReport(
        name=dataset.name,
        source=str(source.path),
        encoding=source.encoding,
        delimiter=source.delimiter,
        n_source_cols=source.n_cols,
    )
    try:
        report.mapping = resolve_dataset(dataset, list(source.headers), contract.normalize)
    except PipelineError as exc:
        report.error = str(exc)
    return report
