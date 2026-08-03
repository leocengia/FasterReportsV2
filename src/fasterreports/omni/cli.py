"""Interfaccia a riga di comando dell'Omni Report.

  omni-report preflight --week 31      controlla i CSV, non apre Excel
  omni-report build --week 31          il "pulsante": produce il workbook
  omni-report contract                 stampa il contratto colonne risolto

Exit code: 0 = fatto, 1 = bloccato da un problema atteso (colonna mancante,
ambigua, tipo non coercibile), 2 = uso sbagliato della CLI.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from ..core.contract import load_contract
from ..core.errors import PipelineError
from .orchestrate import build, run_preflight
from .settings import load_settings

DEFAULT_ROOT = Path(__file__).resolve().parents[3]


def _add_common(p: argparse.ArgumentParser) -> None:
    p.add_argument("--week", required=True, help="numero settimana, es. 31")
    p.add_argument("--input", type=Path, help="cartella sorgenti (default: input/)")
    p.add_argument("--output", type=Path, help="cartella di output (default: output/)")
    p.add_argument("--config", type=Path, help="cartella config (default: config/)")
    p.add_argument(
        "--only", nargs="+", metavar="DATASET",
        help="limita a questi dataset (es. --only Turni 'Slot Only Cases'). "
             "Utile per ricaricare i soli turni dopo un cambio in corsa. "
             "AT_DATASET viene incluso comunque: da lui si ricava la settimana.",
    )


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="omni-report",
        description="Genera l'Omni Report settimanale dai CSV scaricati.",
    )
    sub = p.add_subparsers(dest="command", required=True)

    pf = sub.add_parser("preflight", help="valida i CSV senza aprire Excel")
    _add_common(pf)

    bd = sub.add_parser("build", help="genera il workbook (richiede Excel desktop)")
    _add_common(bd)
    bd.add_argument(
        "--visible", action="store_true", help="mostra Excel durante l'esecuzione"
    )

    ct = sub.add_parser("contract", help="stampa il contratto colonne caricato")
    ct.add_argument("--config", type=Path, help="cartella config (default: config/)")

    ck = sub.add_parser(
        "check",
        help="verifica ambiente e template, senza provare un build",
    )
    ck.add_argument("--config", type=Path, help="cartella config (default: config/)")
    ck.add_argument(
        "--no-excel", action="store_true",
        help="salta la prova di apertura di Excel (utile su una macchina senza Excel)",
    )

    return p


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    config_dir = args.config or DEFAULT_ROOT / "config"

    try:
        contract = load_contract(config_dir / "columns.yml")

        if args.command == "contract":
            _print_contract(contract)
            return 0

        settings = load_settings(config_dir / "settings.yml", root=DEFAULT_ROOT)

        if args.command == "check":
            from .doctor import run_checks

            report = run_checks(contract, settings, try_excel=not args.no_excel)
            print(report.render())
            return 0 if report.ok else 1

        if args.input:
            settings = _replace(settings, input_dir=args.input)
        if args.output:
            settings = _replace(settings, output_dir=args.output)
        if getattr(args, "visible", False):
            settings = _replace(settings, excel=_replace(settings.excel, visible=True))

        only = _resolve_only(contract, getattr(args, "only", None))

        if args.command == "preflight":
            report, _ = run_preflight(contract, settings, only=only)
            path = report.write(settings.preflight_path(args.week))
            print(report.render())
            print(f"Report salvato in: {path}")
            return 0 if report.ok else 1

        if args.command == "build":
            result = build(contract, settings, args.week)
            if not result.ok:
                print(result.report.render(), file=sys.stderr)
                print(f"\nBLOCCATO. Dettagli in: {result.preflight}", file=sys.stderr)
                return 1
            print(f"Preflight: {result.preflight}")
            for w in result.writes:
                print(
                    f"  {w.dataset:14} {w.rows_written:>7} righe -> {w.range_written}"
                    + (f"  [tabella -> {w.table_resized}]" if w.table_resized else "")
                )
                for warn in w.warnings:
                    print(f"      ! {warn}")
            print(f"Macro {'eseguita' if result.macro_ran else 'NON eseguita'}.")
            print(f"\nFatto: {result.workbook}")
            return 0

    except PipelineError as exc:
        print(f"\nERRORE: {exc}\n", file=sys.stderr)
        return 1
    except KeyboardInterrupt:
        print("\nInterrotto.", file=sys.stderr)
        return 130

    return 2


def _replace(settings, **kw):
    from dataclasses import replace

    return replace(settings, **kw)


def _resolve_only(contract, names) -> set[str] | None:
    """Valida `--only` e aggiunge le dipendenze implicite.

    `AT_DATASET` entra sempre: la settimana a cui ritagliare le sorgenti WFM si
    ricava dalle sue date. E il roster entra se si chiede il back office, perché
    il filtro degli agenti viene da lui.
    """
    if not names:
        return None
    unknown = [n for n in names if n not in contract.datasets]
    if unknown:
        raise PipelineError(
            f"--only: dataset sconosciuti: {', '.join(repr(n) for n in unknown)}\n"
            f"  Disponibili: {', '.join(sorted(contract.datasets))}"
        )
    wanted = set(names)
    if any(contract.datasets[n].reader != "csv" for n in wanted):
        wanted.add("AT_DATASET")
    if "Slot Only Cases" in wanted:
        wanted.add("Turni")
    return wanted


def _print_contract(contract) -> None:
    for name, ds in contract.datasets.items():
        print(f"\n### {name}  (foglio {ds.sheet}, header riga {ds.header_row})")
        for f in ds.fields:
            role = "input " if f.is_input else "derivat"
            extra = f"  match={f.match}" if f.match != "exact_first" else ""
            print(f"  {f.target_col:>3}  {role}  {f.dtype:<8} {f.canonical}{extra}")
            if f.aliases:
                print(f"        alias: {', '.join(f.aliases)}")
            if f.consumers:
                print(f"        letta da: {', '.join(f.consumers)}")


if __name__ == "__main__":
    raise SystemExit(main())
