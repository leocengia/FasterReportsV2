"""Risoluzione header sorgente -> colonna canonica.

E' il pezzo anti-fragilita': se l'export sposta o rinomina leggermente una
colonna, qui si assorbe; se il nome e' sparito o e' ambiguo, qui si sbatte la
porta. Mai indovinare in silenzio.

Priorita' di match (in ordine, il primo che risolve vince):

  1. `exact`      nome canonico == header grezzo (a meno di BOM/spazi ai bordi)
  2. `normalized` nome canonico == header normalizzato
  3. `alias_exact`      alias == header grezzo
  4. `alias_normalized` alias == header normalizzato

Perche' l'esatto viene prima del normalizzato — misurato su PSAT_DATASET:
`Agent Name` (col. I) e `agent_name` (col. BO) normalizzano entrambe su
"agent name". Se si partisse dal normalizzato sarebbero ambigue e la pipeline
si fermerebbe su un file perfettamente valido. Col match esatto la colonna
giusta si aggancia e l'altra resta dove sta.

Se piu' header collassano sullo stesso livello di match e nessuno e' esatto ->
AmbiguousColumnError. La sola ambiguita' tollerata e' quella risolta da un
livello superiore.
"""

from __future__ import annotations

from dataclasses import dataclass

from .contract import Dataset, Field
from .errors import AmbiguousColumnError, MissingColumnError
from .normalize import NormalizeRules, normalize, normalize_raw

VIA_EXACT = "esatto"
VIA_NORMALIZED = "normalizzato"
VIA_ALIAS_EXACT = "alias esatto"
VIA_ALIAS_NORMALIZED = "alias normalizzato"


@dataclass(frozen=True)
class Resolution:
    """Come un campo del contratto e' stato agganciato al sorgente."""

    canonical: str
    target_col: str
    source_header: str
    source_index: int  # 0-based, posizione nel CSV
    via: str
    matched_on: str  # il testo (canonico o alias) che ha fatto match
    shadowed: tuple[str, ...] = ()  # header scartati perche' risolti da un livello superiore

    @property
    def was_disambiguated(self) -> bool:
        return bool(self.shadowed)


@dataclass(frozen=True)
class DatasetMapping:
    dataset: str
    resolutions: tuple[Resolution, ...]
    source_headers: tuple[str, ...]
    unused_headers: tuple[str, ...]

    def by_canonical(self) -> dict[str, Resolution]:
        return {r.canonical: r for r in self.resolutions}


def resolve_field(
    dataset_name: str,
    fld: Field,
    source_headers: list[str],
    rules: NormalizeRules,
) -> Resolution:
    """Aggancia un singolo campo. Solleva se manca o e' ambiguo."""
    raw = [normalize_raw(h) for h in source_headers]
    norm = [normalize(h, rules) for h in source_headers]

    exact_only = fld.match == "exact"

    # --- livello 1: nome canonico, confronto grezzo -------------------------
    hits = _indices(raw, normalize_raw(fld.canonical))
    if len(hits) == 1:
        return _make(fld, source_headers, hits[0], VIA_EXACT, fld.canonical, [])
    if len(hits) > 1:
        # Due colonne con header identico: nessun criterio per scegliere.
        raise AmbiguousColumnError(
            dataset_name, fld.canonical, [source_headers[i] for i in hits], VIA_EXACT
        )

    if exact_only:
        raise MissingColumnError(
            dataset_name, fld.canonical, list(fld.aliases), source_headers, exact_only=True
        )

    # --- livello 2: nome canonico, confronto normalizzato -------------------
    target = normalize(fld.canonical, rules)
    hits = _indices(norm, target)
    if len(hits) == 1:
        return _make(fld, source_headers, hits[0], VIA_NORMALIZED, fld.canonical, [])
    if len(hits) > 1:
        raise AmbiguousColumnError(
            dataset_name, fld.canonical, [source_headers[i] for i in hits], VIA_NORMALIZED
        )

    # --- livelli 3 e 4: alias ---------------------------------------------
    for via, table, transform in (
        (VIA_ALIAS_EXACT, raw, normalize_raw),
        (VIA_ALIAS_NORMALIZED, norm, lambda a: normalize(a, rules)),
    ):
        found: list[tuple[int, str]] = []
        for alias in fld.aliases:
            for i in _indices(table, transform(alias)):
                found.append((i, alias))
        # Piu' alias possono puntare allo stesso header: e' un match unico.
        distinct = {i for i, _ in found}
        if len(distinct) == 1:
            i, alias = found[0]
            return _make(fld, source_headers, i, via, alias, [])
        if len(distinct) > 1:
            raise AmbiguousColumnError(
                dataset_name, fld.canonical, [source_headers[i] for i in sorted(distinct)], via
            )

    raise MissingColumnError(dataset_name, fld.canonical, list(fld.aliases), source_headers)


def resolve_dataset(
    dataset: Dataset,
    source_headers: list[str],
    rules: NormalizeRules,
) -> DatasetMapping:
    """Aggancia tutti i campi input di un dataset.

    Raccoglie *tutti* gli errori prima di sollevare: correggere un CSV una
    colonna per volta, con un run per errore, e' tempo buttato.
    """
    resolutions: list[Resolution] = []
    problems: list[Exception] = []

    for fld in dataset.input_fields:
        try:
            resolutions.append(resolve_field(dataset.name, fld, source_headers, rules))
        except (MissingColumnError, AmbiguousColumnError) as exc:
            problems.append(exc)

    if problems:
        raise _combine(dataset.name, problems)

    # Due campi diversi agganciati alla stessa colonna sorgente: il contratto
    # e' ambiguo rispetto a questo export, non e' un caso da far passare.
    by_index: dict[int, list[Resolution]] = {}
    for r in resolutions:
        by_index.setdefault(r.source_index, []).append(r)
    for idx, rs in by_index.items():
        if len(rs) > 1:
            raise AmbiguousColumnError(
                dataset.name,
                " / ".join(r.canonical for r in rs),
                [source_headers[idx]],
                "collisione: campi diversi agganciati alla stessa colonna sorgente",
            )

    # Annota gli header scartati: un header che normalizza come uno agganciato
    # e' esattamente il caso Agent Name / agent_name. Va tracciato nel
    # preflight, cosi' si vede che la disambiguazione e' avvenuta.
    used = {r.source_index for r in resolutions}
    annotated: list[Resolution] = []
    for r in resolutions:
        tgt = normalize(r.source_header, rules)
        shadowed = tuple(
            h
            for i, h in enumerate(source_headers)
            if i not in used and normalize(h, rules) == tgt
        )
        annotated.append(
            Resolution(
                canonical=r.canonical,
                target_col=r.target_col,
                source_header=r.source_header,
                source_index=r.source_index,
                via=r.via,
                matched_on=r.matched_on,
                shadowed=shadowed,
            )
        )

    return DatasetMapping(
        dataset=dataset.name,
        resolutions=tuple(annotated),
        source_headers=tuple(source_headers),
        unused_headers=tuple(h for i, h in enumerate(source_headers) if i not in used),
    )


def _indices(table: list[str], needle: str) -> list[int]:
    return [i for i, v in enumerate(table) if v == needle]


def _make(
    fld: Field,
    headers: list[str],
    index: int,
    via: str,
    matched_on: str,
    shadowed: list[str],
) -> Resolution:
    return Resolution(
        canonical=fld.canonical,
        target_col=fld.target_col,
        source_header=headers[index],
        source_index=index,
        via=via,
        matched_on=matched_on,
        shadowed=tuple(shadowed),
    )


def _combine(dataset: str, problems: list[Exception]) -> Exception:
    if len(problems) == 1:
        return problems[0]
    joined = "\n\n".join(str(p) for p in problems)
    kind = (
        AmbiguousColumnError
        if all(isinstance(p, AmbiguousColumnError) for p in problems)
        else MissingColumnError
    )
    # Costruito a mano: i costruttori specializzati non reggono un aggregato.
    exc = kind.__new__(kind)
    Exception.__init__(
        exc,
        f"Dataset {dataset}: {len(problems)} colonne non risolte.\n\n{joined}",
    )
    exc.dataset = dataset
    exc.canonical = "(piu' campi)"
    return exc
