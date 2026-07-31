"""Caricamento di config/settings.yml."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml

from ..core.errors import ContractError


@dataclass(frozen=True)
class ExcelSettings:
    visible: bool = False
    macro: str = "Refresh_Dettaglio_Malpractice"
    silent_mode_flag: str = "SilentMode"
    full_rebuild: bool = True


@dataclass(frozen=True)
class ValidationSettings:
    max_uncoercible_ratio: float = 0.02
    min_rows: int = 1


@dataclass(frozen=True)
class Settings:
    root: Path
    source: str = "csv"
    input_dir: Path = Path("input")
    output_dir: Path = Path("output")
    template: Path = Path("template/Omni_Report_TEMPLATE.xlsm")
    input_files: dict[str, str] = field(default_factory=dict)
    workbook_name: str = "Omni_Report_W{week}.xlsm"
    preflight_name: str = "preflight_W{week}.txt"
    excel: ExcelSettings = field(default_factory=ExcelSettings)
    validation: ValidationSettings = field(default_factory=ValidationSettings)
    derived_mode: str = "formula"

    def workbook_path(self, week: str) -> Path:
        return self.output_dir / self.workbook_name.format(week=week)

    def preflight_path(self, week: str) -> Path:
        return self.output_dir / self.preflight_name.format(week=week)

    def input_path(self, dataset: str) -> Path:
        try:
            return self.input_dir / self.input_files[dataset]
        except KeyError:
            raise ContractError(
                f"settings.yml: manca il nome file per il dataset {dataset!r} "
                f"in input_files. Presenti: {', '.join(sorted(self.input_files))}"
            ) from None


def load_settings(path: str | Path, *, root: Path | None = None) -> Settings:
    path = Path(path)
    if not path.is_file():
        raise ContractError(f"settings.yml non trovato: {path}")
    raw = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    root = root or path.parent.parent

    paths = raw.get("paths") or {}
    out = raw.get("output") or {}
    xl = raw.get("excel") or {}
    val = raw.get("validation") or {}
    derived = raw.get("derived") or {}

    source = str(raw.get("source", "csv"))
    if source not in ("csv", "sql"):
        raise ContractError(f"settings.yml: source={source!r} non valido (csv|sql).")
    if source == "sql":
        raise ContractError(
            "settings.yml: source=sql non e' ancora implementato (piano §13).\n"
            "  Il contratto colonne non cambia: serve un sqlsource.py che restituisca\n"
            "  gli stessi header+righe di csvsource. Vedi docs/architettura.md §7."
        )

    mode = str(derived.get("mode", "formula"))
    if mode not in ("formula", "python"):
        raise ContractError(f"settings.yml: derived.mode={mode!r} non valido (formula|python).")

    def resolve(p: str | Path) -> Path:
        p = Path(p)
        return p if p.is_absolute() else root / p

    return Settings(
        root=root,
        source=source,
        input_dir=resolve(paths.get("input", "input")),
        output_dir=resolve(paths.get("output", "output")),
        template=resolve(paths.get("template", "template/Omni_Report_TEMPLATE.xlsm")),
        input_files=dict(raw.get("input_files") or {}),
        workbook_name=str(out.get("workbook_name", "Omni_Report_W{week}.xlsm")),
        preflight_name=str(out.get("preflight_name", "preflight_W{week}.txt")),
        excel=ExcelSettings(
            visible=bool(xl.get("visible", False)),
            macro=str(xl.get("macro", "Refresh_Dettaglio_Malpractice")),
            silent_mode_flag=str(xl.get("silent_mode_flag", "SilentMode")),
            full_rebuild=bool(xl.get("full_rebuild", True)),
        ),
        validation=ValidationSettings(
            max_uncoercible_ratio=float(val.get("max_uncoercible_ratio", 0.02)),
            min_rows=int(val.get("min_rows", 1)),
        ),
        derived_mode=mode,
    )
