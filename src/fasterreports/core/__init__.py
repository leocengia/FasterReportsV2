"""Core condiviso: ingestione, matching header, coercizioni, preflight.

Non dipende da Excel ne' da xlwings: gira e si testa su qualunque macchina.
E' il punto di contatto fra Omni Report e WOW AHT — entrambi partono da export
caso-per-caso con lo stesso problema di colonne che si spostano
(vedi docs/contesto-wow-aht.md §7, "Un solo ingest, molti report").
"""

from .coherence import BLOCCA, SEGNALA, CoherenceReport, Finding, check_sources
from .contract import Contract, Dataset, Field, load_contract, parse_contract
from .csvsource import CsvSource, read_csv, read_csv_text
from .errors import (
    AmbiguousColumnError,
    CoercionError,
    ContractError,
    EmptyDatasetError,
    MissingColumnError,
    PipelineError,
    SourceError,
)
from .matcher import DatasetMapping, Resolution, resolve_dataset, resolve_field
from .names import full_name, normalize_name, normalize_skill
from .preflight import DatasetReport, PreflightReport, check_dataset
from .shifts import Shift, ShiftParseError, Slot, parse_shift, parse_slot
from .tablesource import TableSource, is_excel, read_table
from .transform import Block, add_derived, build_block
from .wfmsource import (
    SourceNotes,
    TidySource,
    read_alias_map,
    read_backoffice,
    read_roster,
    week_bounds,
    week_from_dates,
)
from .xlsxsource import read_sheet, read_sheet_names

__all__ = [
    "BLOCCA",
    "SEGNALA",
    "AmbiguousColumnError",
    "Block",
    "CoercionError",
    "CoherenceReport",
    "Contract",
    "ContractError",
    "CsvSource",
    "Dataset",
    "DatasetMapping",
    "DatasetReport",
    "EmptyDatasetError",
    "Field",
    "Finding",
    "MissingColumnError",
    "PipelineError",
    "PreflightReport",
    "Resolution",
    "Shift",
    "ShiftParseError",
    "Slot",
    "SourceError",
    "SourceNotes",
    "TableSource",
    "TidySource",
    "add_derived",
    "build_block",
    "check_dataset",
    "check_sources",
    "full_name",
    "load_contract",
    "normalize_name",
    "normalize_skill",
    "parse_contract",
    "parse_shift",
    "parse_slot",
    "read_alias_map",
    "read_backoffice",
    "read_csv",
    "read_csv_text",
    "read_roster",
    "read_sheet",
    "read_table",
    "read_sheet_names",
    "is_excel",
    "resolve_dataset",
    "resolve_field",
    "week_bounds",
    "week_from_dates",
]
