"""Caricamento e validazione del contratto colonne (config/columns.yml).

Il contratto e' l'unico posto dove sta scritto "questa colonna del CSV va in
questa lettera del foglio". Se e' incoerente si ferma qui, prima di aprire Excel.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml

from .errors import ContractError, DuplicateTargetError
from .normalize import NormalizeRules

MatchMode = str  # "exact_first" | "exact"
Role = str  # "input" | "derived"


def col_to_index(letters: str) -> int:
    """'A' -> 1, 'AA' -> 27. Indice 1-based, come Excel."""
    s = str(letters).strip().upper()
    if not s or not s.isalpha():
        raise ContractError(f"Lettera di colonna non valida: {letters!r}")
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - 64)
    return n


def index_to_col(index: int) -> str:
    """1 -> 'A', 27 -> 'AA'."""
    if index < 1:
        raise ContractError(f"Indice di colonna non valido: {index}")
    out = ""
    while index:
        index, rem = divmod(index - 1, 26)
        out = chr(65 + rem) + out
    return out


@dataclass(frozen=True)
class Field:
    canonical: str
    target_col: str
    role: Role
    dtype: str = "str"
    match: MatchMode = "exact_first"
    # Formato `strptime` con cui leggere la data, per i campi in cui indovinare
    # sarebbe un lancio di moneta (`8/10/2026`: 10 agosto o 8 ottobre?). Vuoto =
    # si prova la lista di formati noti, che va bene per le date non ambigue.
    date_format: str = ""
    aliases: tuple[str, ...] = ()
    consumers: tuple[str, ...] = ()
    source: str | None = None  # per role=derived: da quale campo input discende
    formula: str | None = None
    notes: str | None = None

    @property
    def target_index(self) -> int:
        return col_to_index(self.target_col)

    @property
    def is_input(self) -> bool:
        return self.role == "input"


@dataclass(frozen=True)
class Dataset:
    name: str
    sheet: str
    header_row: int
    data_start_col: str
    fields: tuple[Field, ...]
    data_end_col: str | None = None
    list_object: str | None = None
    max_template_row: int | None = None
    derive_offset_from: str | None = None
    # Da quale lettore arriva il dataset. `csv` (default) legge da input/*.csv;
    # `wfm_roster` e `wfm_backoffice` da matrici larghe .xlsx che vanno prima
    # portate in forma lunga (core/wfmsource.py).
    reader: str = "csv"
    # Se la fonte manca, il preflight SEGNALA invece di BLOCCARE, e il build
    # scrive il foglio VUOTO (non lo lascia stare: un residuo della settimana
    # prima nel template sarebbe un report sbagliato, non uno che manca).
    # Va messo a True solo se nessuna regola di malpractice dipende dal
    # dataset — verificalo in `consumers` prima di cambiarlo: se anche un solo
    # campo e' letto da 'VBA:...', la fonte non e' opzionale.
    optional: bool = False
    # Il campo che rende una riga un RECORD. Se e' vuoto, la riga non e' un dato:
    # e' il preambolo, una riga di separazione, o la coda dei totali. Serve agli
    # export che non sono tabelle pulite — il report Salesforce formattato
    # finisce con `Total | Sum | 72,13` e `Count | 178`, che senza questo
    # diventerebbero due agenti di nome 'Total' e ''.
    key_field: str | None = None
    # Valori del `key_field` che CHIUDONO la tabella: dalla prima riga che ne
    # contiene uno, si smette di leggere. `Total` e' l'etichetta del totale
    # generale dei report SF.
    stop_values: tuple[str, ...] = ()
    # Il foglio consumatore legge questo dataset CELLA PER CELLA, non per
    # intervalli: `Duplicates Helper!A2` e' `DUP_DATASET!B15`, `A3` e' `B16`, e
    # cosi' via. Il limite di riga sta quindi nel riferimento puntuale piu' alto,
    # e `templatescan.scan_row_limits` — che cerca intervalli — non lo vede.
    # Con questo flag lo cerca anche `scan_cell_refs`.
    #
    # NON metterlo su un dataset che non e' letto cosi'. Una formula che punta a
    # UNA riga scelta a mano — il caso classico e' un commento della settimana
    # preso da una riga fissa — non e' un limite di capienza, e' una selezione:
    # letta come limite bloccherebbe ogni settimana con piu' righe di quella.
    # (Nel template di oggi non ce ne sono; il test
    # `test_un_riferimento_a_riga_fissa_diventerebbe_un_limite_falso` costruisce
    # il caso, perche' la proprieta' da difendere e' dello scanner.)
    read_by_row: bool = False
    # I fogli che vivono di questo dataset. Se la fonte e' `optional` e manca, le
    # loro celle di errore sono ATTESE (una divisione per un conteggio a zero) e
    # non devono far dichiarare fallito un build che invece e' andato come
    # doveva. Le celle di errore di tutti gli altri fogli continuano a contare.
    dependent_sheets: tuple[str, ...] = ()

    @property
    def input_fields(self) -> tuple[Field, ...]:
        return tuple(f for f in self.fields if f.is_input)

    @property
    def derived_fields(self) -> tuple[Field, ...]:
        return tuple(f for f in self.fields if not f.is_input)

    @property
    def start_index(self) -> int:
        return col_to_index(self.data_start_col)

    @property
    def last_input_index(self) -> int:
        """Ultima colonna in cui la pipeline scrive valori."""
        return max(f.target_index for f in self.input_fields)


@dataclass(frozen=True)
class Contract:
    normalize: NormalizeRules
    datasets: dict[str, Dataset] = field(default_factory=dict)

    def dataset(self, name: str) -> Dataset:
        try:
            return self.datasets[name]
        except KeyError:
            raise ContractError(
                f"Dataset {name!r} non presente nel contratto. "
                f"Disponibili: {', '.join(sorted(self.datasets))}"
            ) from None


_VALID_DTYPES = {"str", "float", "int", "datetime"}
_VALID_ROLES = {"input", "derived"}
_VALID_MATCH = {"exact_first", "exact"}
# `csv` e `table` sono sinonimi: un dataset tabellare, che arrivi come CSV o
# come foglio Excel. Il formato lo decide l'estensione del file, non questo.
_VALID_READERS = {"csv", "table", "wfm_roster", "wfm_backoffice"}


def load_contract(path: str | Path) -> Contract:
    path = Path(path)
    if not path.is_file():
        raise ContractError(f"Contratto non trovato: {path}")
    raw = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    return parse_contract(raw)


def parse_contract(raw: dict) -> Contract:
    rules = NormalizeRules.from_dict(raw.get("normalize"))
    ds_raw = raw.get("datasets")
    if not isinstance(ds_raw, dict) or not ds_raw:
        raise ContractError("Il contratto non contiene alcun dataset.")

    datasets: dict[str, Dataset] = {}
    for name, body in ds_raw.items():
        if not isinstance(body, dict):
            raise ContractError(f"Dataset {name}: definizione non valida.")
        for key in ("sheet", "header_row", "data_start_col", "fields"):
            if key not in body:
                raise ContractError(f"Dataset {name}: manca la chiave obbligatoria {key!r}.")

        fields: list[Field] = []
        seen_target: dict[int, str] = {}
        seen_canonical: set[str] = set()
        for fraw in body["fields"]:
            f = _parse_field(name, fraw)
            if f.canonical in seen_canonical:
                raise ContractError(
                    f"Dataset {name}: campo canonico {f.canonical!r} definito due volte."
                )
            seen_canonical.add(f.canonical)
            idx = f.target_index
            if idx in seen_target:
                raise DuplicateTargetError(
                    f"Dataset {name}: colonna target {f.target_col} assegnata sia a "
                    f"{seen_target[idx]!r} sia a {f.canonical!r}. "
                    f"Due campi non possono scrivere nella stessa cella."
                )
            seen_target[idx] = f.canonical
            fields.append(f)

        if not any(f.is_input for f in fields):
            raise ContractError(f"Dataset {name}: nessun campo con role=input.")

        start = col_to_index(body["data_start_col"])
        for f in fields:
            if f.is_input and f.target_index < start:
                raise ContractError(
                    f"Dataset {name}: il campo {f.canonical!r} punta a {f.target_col}, "
                    f"prima di data_start_col={body['data_start_col']}."
                )

        # I campi derivati devono discendere da un input dichiarato: e' cio' che
        # rende verificabile la nota P/Q del piano.
        inputs = {f.canonical for f in fields if f.is_input}
        for f in fields:
            if not f.is_input and f.source and f.source not in inputs:
                raise ContractError(
                    f"Dataset {name}: il campo derivato {f.canonical!r} dichiara "
                    f"source={f.source!r}, che non e' un campo input di questo dataset."
                )

        reader = str(body.get("reader", "csv"))
        if reader not in _VALID_READERS:
            raise ContractError(
                f"Dataset {name}: reader={reader!r} non valido "
                f"(ammessi: {', '.join(sorted(_VALID_READERS))})."
            )

        optional = bool(body.get("optional", False))
        if optional:
            # Difesa contro l'errore piu' costoso possibile qui: marcare come
            # opzionale una fonte che il motore usa davvero. Se manca, il VBA
            # non se ne accorgerebbe: i numeri sarebbero sbagliati, non assenti.
            consumatori_vba = sorted({
                c for f in fields for c in f.consumers if c.startswith("VBA:")
            })
            if consumatori_vba:
                raise ContractError(
                    f"Dataset {name}: optional=true ma e' consumato dal VBA "
                    f"({', '.join(consumatori_vba)}). Una fonte che alimenta il "
                    f"motore di malpractice non puo' essere opzionale: se "
                    f"manca, quei calcoli sarebbero sbagliati in silenzio, non "
                    f"solo assenti."
                )

        # `key_field` deve essere un campo input dichiarato: se punta a un nome
        # che non esiste, ogni riga risulterebbe senza chiave e il dataset
        # uscirebbe VUOTO — un export intero scartato in silenzio.
        key_field = body.get("key_field")
        if key_field is not None:
            key_field = str(key_field)
            if key_field not in inputs:
                raise ContractError(
                    f"Dataset {name}: key_field={key_field!r} non e' un campo "
                    f"input di questo dataset.\n"
                    f"  Campi disponibili: {', '.join(sorted(inputs))}\n"
                    f"  Con una chiave che non esiste ogni riga risulterebbe "
                    f"senza chiave, e il dataset uscirebbe vuoto senza errori."
                )
        stop_values = tuple(str(v) for v in (body.get("stop_values") or []))
        if stop_values and not key_field:
            raise ContractError(
                f"Dataset {name}: stop_values e' impostato ma key_field no. "
                f"I valori che chiudono la tabella si cercano NELLA chiave: "
                f"senza key_field non si sa dove guardarli."
            )

        datasets[name] = Dataset(
            name=name,
            reader=reader,
            optional=optional,
            key_field=key_field,
            stop_values=stop_values,
            read_by_row=bool(body.get("read_by_row", False)),
            dependent_sheets=tuple(
                str(s) for s in (body.get("dependent_sheets") or [])
            ),
            sheet=str(body["sheet"]),
            header_row=int(body["header_row"]),
            data_start_col=str(body["data_start_col"]).upper(),
            data_end_col=(str(body["data_end_col"]).upper() if body.get("data_end_col") else None),
            list_object=body.get("list_object"),
            max_template_row=(int(body["max_template_row"]) if body.get("max_template_row") else None),
            derive_offset_from=body.get("derive_offset_from"),
            fields=tuple(fields),
        )

    return Contract(normalize=rules, datasets=datasets)


def _parse_field(dataset: str, fraw: dict) -> Field:
    if not isinstance(fraw, dict):
        raise ContractError(f"Dataset {dataset}: campo non valido: {fraw!r}")
    for key in ("canonical", "target_col", "role"):
        if key not in fraw:
            raise ContractError(f"Dataset {dataset}: campo senza {key!r}: {fraw!r}")

    role = str(fraw["role"])
    if role not in _VALID_ROLES:
        raise ContractError(
            f"Dataset {dataset}, campo {fraw['canonical']!r}: role={role!r} non valido "
            f"(ammessi: {', '.join(sorted(_VALID_ROLES))})."
        )
    dtype = str(fraw.get("dtype", "str"))
    if dtype not in _VALID_DTYPES:
        raise ContractError(
            f"Dataset {dataset}, campo {fraw['canonical']!r}: dtype={dtype!r} non valido "
            f"(ammessi: {', '.join(sorted(_VALID_DTYPES))})."
        )
    match = str(fraw.get("match", "exact_first"))
    if match not in _VALID_MATCH:
        raise ContractError(
            f"Dataset {dataset}, campo {fraw['canonical']!r}: match={match!r} non valido "
            f"(ammessi: {', '.join(sorted(_VALID_MATCH))})."
        )

    date_format = str(fraw.get("date_format") or "")
    if date_format and dtype != "datetime":
        raise ContractError(
            f"Dataset {dataset}, campo {fraw['canonical']!r}: date_format vale solo "
            f"per dtype=datetime, qui il dtype e' {dtype!r}."
        )
    if date_format and "%" not in date_format:
        raise ContractError(
            f"Dataset {dataset}, campo {fraw['canonical']!r}: date_format="
            f"{date_format!r} non contiene direttive strptime.\n"
            f"  Esempio per un export americano: \"%m/%d/%Y %I:%M:%S %p\"."
        )

    col_to_index(fraw["target_col"])  # valida subito la lettera

    return Field(
        canonical=str(fraw["canonical"]),
        target_col=str(fraw["target_col"]).upper(),
        role=role,
        dtype=dtype,
        match=match,
        date_format=date_format,
        aliases=tuple(str(a) for a in (fraw.get("aliases") or [])),
        consumers=tuple(str(c) for c in (fraw.get("consumers") or [])),
        source=(str(fraw["source"]) if fraw.get("source") else None),
        formula=(str(fraw["formula"]) if fraw.get("formula") else None),
        notes=(str(fraw["notes"]) if fraw.get("notes") else None),
    )
