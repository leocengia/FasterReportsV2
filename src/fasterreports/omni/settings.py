"""Caricamento di config/settings.yml."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml

from ..core.errors import ContractError, SourceError


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
class SourceSettings:
    """Parametri delle sorgenti WFM (roster turni e back office)."""

    skills: tuple[str, ...] = ("HPO",)
    include_marked_skills: bool = False
    roster_sheet: str | None = "publish"
    backoffice_sheet: str | None = "Only Cases Shifts"
    backoffice_sections: tuple[str, ...] | None = None
    monday_serial: int | None = None


@dataclass(frozen=True)
class Settings:
    root: Path
    source: str = "csv"
    input_dir: Path = Path("input")
    output_dir: Path = Path("output")
    template: Path = Path("template/Omni_Report_TEMPLATE.xlsm")
    # Lo storico settimanale di volume/AHT per case type. NON sta in output/:
    # quella cartella e' usa-e-getta (ed e' fuori da git), mentre questo file e'
    # la memoria del trend — l'unica cosa del progetto che non si puo'
    # ricostruire rilanciando il programma, perche' gli export delle settimane
    # passate non li abbiamo piu'. Va versionato.
    aht_history: Path = Path("data/aht_history.csv")
    # I case type che si VEDONO nelle heat map di 'AHT Trend WoW'. Vuoto = non
    # filtrare. Vedi il commento in settings.yml: la lista e' misurata sulle heat
    # map del file legacy e sulle 11 settimane di storico curate a mano, non
    # scelta a tavolino.
    casetype_heatmap: tuple[str, ...] = ()
    # I case type che nell'archivio ci sono e nelle heat map si e' deciso di NON
    # far entrare. Non filtra niente in produzione: e' la meta' che manca alla
    # domanda «ogni case type dello storico e' stato guardato?», e senza di lei
    # l'unico modo di far tacere il test sarebbe far entrare tutto.
    casetype_heatmap_esclusi: tuple[str, ...] = ()
    input_files: dict[str, str] = field(default_factory=dict)
    workbook_name: str = "Omni_Report_W{week}.xlsm"
    preflight_name: str = "preflight_W{week}.txt"
    excel: ExcelSettings = field(default_factory=ExcelSettings)
    validation: ValidationSettings = field(default_factory=ValidationSettings)
    sources: SourceSettings = field(default_factory=SourceSettings)
    derived_mode: str = "formula"
    contratti: dict[str, str] = field(default_factory=dict)

    def workbook_path(self, week: str) -> Path:
        return self.output_dir / self.workbook_name.format(week=week)

    def preflight_path(self, week: str) -> Path:
        return self.output_dir / self.preflight_name.format(week=week)

    def input_path(self, dataset: str) -> Path:
        """Il file della sorgente, trovato per **pattern** e non per nome esatto.

        Gli export reali si chiamano `AT DATASET W31.xlsx`, `SF DATABASE W31.csv`:
        il numero di settimana e' nel nome e cambia ogni volta, e l'estensione
        dipende da chi produce l'export. Pretendere `AT.csv` costringerebbe a
        rinominare quattro file a mano ogni settimana — cioe' a reintrodurre
        proprio il passaggio manuale che si vuole togliere.

        Il pattern deve corrispondere a **un solo** file: zero o molti sono
        entrambi errori, e si dice quali file c'erano.
        """
        try:
            pattern = self.input_files[dataset]
        except KeyError:
            raise ContractError(
                f"settings.yml: manca il pattern per il dataset {dataset!r} "
                f"in input_files. Presenti: {', '.join(sorted(self.input_files))}"
            ) from None

        # Un pattern senza caratteri jolly resta un nome esatto: comodo per
        # chi preferisce nomi fissi.
        if not any(ch in pattern for ch in "*?["):
            return self.input_dir / pattern

        if not self.input_dir.is_dir():
            raise SourceError(
                f"Cartella delle sorgenti inesistente: {self.input_dir}"
            )
        trovati = sorted(
            p for p in self.input_dir.glob(pattern)
            if p.is_file() and not p.name.startswith("~$")
        )
        if len(trovati) == 1:
            return trovati[0]

        presenti = sorted(
            p.name for p in self.input_dir.iterdir()
            if p.is_file() and p.name != ".gitkeep"
        )
        if not trovati:
            raise SourceError(
                f"{dataset}: nessun file corrisponde a {pattern!r} in "
                f"{self.input_dir}.\n"
                f"  File presenti: "
                f"{', '.join(repr(n) for n in presenti) if presenti else '(nessuno)'}\n"
                f"  Il pattern si cambia in config/settings.yml -> input_files."
            )
        raise SourceError(
            f"{dataset}: {len(trovati)} file corrispondono a {pattern!r}, "
            f"non so quale usare.\n"
            f"  {', '.join(repr(p.name) for p in trovati)}\n"
            f"  Togli dalla cartella quelli della settimana vecchia, oppure "
            f"restringi il pattern."
        )


def _str_tuple(value, campo: str) -> tuple[str, ...] | None:
    """Una lista di stringhe da YAML — mai una stringa nuda.

    Bug reale, non ipotetico: una stringa e' iterabile carattere per carattere,
    quindi `tuple("HPO")` da' `('H', 'P', 'O')` senza nessun errore. Chi scrive
    `skills: HPO` invece di `skills: ["HPO"]` otterrebbe un filtro completamente
    sbagliato — e nessun messaggio direbbe perche', perche' tecnicamente non e'
    successo nessun errore Python.
    """
    if value is None:
        return None
    if isinstance(value, str):
        raise ContractError(
            f"settings.yml: {campo} deve essere una LISTA, non una stringa "
            f"({value!r}).\n"
            f"  Scritto cosi':      {campo}: [\"{value}\"]\n"
            f"  non cosi':          {campo}: {value}\n"
            f"  Una stringa e' iterabile carattere per carattere: senza questo "
            f"controllo il filtro diventerebbe sbagliato senza che nessun "
            f"errore lo segnali."
        )
    return tuple(str(v) for v in value)


def _int_or_none(value, campo: str) -> int | None:
    """Come `int(...)`, ma con un `ContractError` invece di un `ValueError` nudo."""
    if value in (None, ""):
        return None
    try:
        return int(value)
    except (TypeError, ValueError):
        raise ContractError(
            f"settings.yml: {campo} deve essere un numero intero (il seriale "
            f"Excel del lunedi' della settimana), non {value!r}."
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

    # La chiave vecchia non si ignora: significava il CONTRARIO di quella nuova,
    # e lasciarla in un file di config senza che nessuno la legga vorrebbe dire
    # che chi la modifica non ottiene niente e non lo sa.
    if "casetype_esclusi" in (raw.get("aht_history") or {}):
        raise ContractError(
            "settings.yml: `aht_history.casetype_esclusi` non esiste piu'.\n"
            "  Era una lista di case type da TENERE FUORI dal trend. Al suo posto\n"
            "  c'e' `aht_history.casetype_heatmap`, che elenca quelli da FARCI\n"
            "  ENTRARE — cioe' l'opposto.\n"
            "  Il motivo del cambio: con la sola lista degli esclusi, un case type\n"
            "  mai visto prima entrava nelle heat map da se', perche'\n"
            "  'AHT Trend WoW'!A5 mostra tutto quello che trova nello storico.\n"
            "  Con la lista degli inclusi le heat map hanno un numero di righe\n"
            "  fisso, e un case type nuovo viene SEGNALATO invece che aggiunto.\n"
            "  Togli la chiave vecchia; se ti serve la lista, e' nel repository."
        )

    # Un nome in tutte e due le liste non e' una svista da correggere in silenzio:
    # e' una decisione che si contraddice, e le due liste servono proprio a dire
    # che la decisione e' stata presa. Chi legge il file non saprebbe quale vale.
    _dentro = set(
        _str_tuple(
            (raw.get("aht_history") or {}).get("casetype_heatmap"),
            "aht_history.casetype_heatmap",
        )
        or ()
    )
    _fuori = set(
        _str_tuple(
            (raw.get("aht_history") or {}).get("casetype_heatmap_esclusi"),
            "aht_history.casetype_heatmap_esclusi",
        )
        or ()
    )
    if _dentro & _fuori:
        raise ContractError(
            "settings.yml: questi case type sono sia in `casetype_heatmap` sia in "
            "`casetype_heatmap_esclusi`:\n"
            + "".join(f"    {c}\n" for c in sorted(_dentro & _fuori))
            + "  Le due liste dicono il contrario l'una dell'altra: la prima cosa\n"
            "  si vede nelle heat map, la seconda cosa si e' deciso di tenerne\n"
            "  fuori. Tienilo in una sola."
        )

    mode = str(derived.get("mode", "formula"))
    if mode not in ("formula", "python"):
        raise ContractError(f"settings.yml: derived.mode={mode!r} non valido (formula|python).")

    def resolve(p: str | Path) -> Path:
        p = Path(p)
        return p if p.is_absolute() else root / p

    src = raw.get("sources") or {}
    sources = SourceSettings(
        skills=_str_tuple(src.get("skills"), "sources.skills") or ("HPO",),
        include_marked_skills=bool(src.get("include_marked_skills", False)),
        roster_sheet=src.get("roster_sheet") or None,
        backoffice_sheet=src.get("backoffice_sheet") or None,
        backoffice_sections=_str_tuple(
            src.get("backoffice_sections"), "sources.backoffice_sections"
        ),
        monday_serial=_int_or_none(src.get("monday_serial"), "sources.monday_serial"),
    )

    # La mappa dei contratti è un file a parte: è dato HR, cambia con altri
    # tempi rispetto ai percorsi e alle scelte di esecuzione.
    contratti: dict[str, str] = {}
    contratti_path = path.parent / "contratti.yml"
    if contratti_path.is_file():
        body = yaml.safe_load(contratti_path.read_text(encoding="utf-8")) or {}
        raw_map = body.get("contratti") or {}
        if not isinstance(raw_map, dict):
            raise ContractError(
                f"{contratti_path.name}: 'contratti' deve essere una mappa "
                f"nome -> contratto."
            )
        from ..core.names import normalize_name

        contratti = {normalize_name(k): str(v) for k, v in raw_map.items() if v}

    return Settings(
        root=root,
        source=source,
        sources=sources,
        contratti=contratti,
        input_dir=resolve(paths.get("input", "input")),
        output_dir=resolve(paths.get("output", "output")),
        template=resolve(paths.get("template", "template/Omni_Report_TEMPLATE.xlsm")),
        aht_history=resolve(paths.get("aht_history", "data/aht_history.csv")),
        casetype_heatmap=_str_tuple(
            (raw.get("aht_history") or {}).get("casetype_heatmap"),
            "aht_history.casetype_heatmap",
        ) or (),
        casetype_heatmap_esclusi=_str_tuple(
            (raw.get("aht_history") or {}).get("casetype_heatmap_esclusi"),
            "aht_history.casetype_heatmap_esclusi",
        ) or (),
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
