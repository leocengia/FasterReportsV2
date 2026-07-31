"""Core condiviso: ingestione, matching header, coercizioni, preflight.

Non dipende da Excel ne' da xlwings: gira e si testa su qualunque macchina.
E' il punto di contatto fra Omni Report e WOW AHT — entrambi partono da export
caso-per-caso con lo stesso problema di colonne che si spostano
(vedi docs/contesto-wow-aht.md §7, "Un solo ingest, molti report").
"""

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
from .preflight import DatasetReport, PreflightReport, check_dataset
from .transform import Block, add_derived, build_block

__all__ = [
    "AmbiguousColumnError",
    "Block",
    "CoercionError",
    "Contract",
    "ContractError",
    "CsvSource",
    "Dataset",
    "DatasetMapping",
    "DatasetReport",
    "EmptyDatasetError",
    "Field",
    "MissingColumnError",
    "PipelineError",
    "PreflightReport",
    "Resolution",
    "SourceError",
    "add_derived",
    "build_block",
    "check_dataset",
    "load_contract",
    "parse_contract",
    "read_csv",
    "read_csv_text",
    "resolve_dataset",
    "resolve_field",
]
