#!/usr/bin/env python3
"""Confronta `Turni` e `Slot Only Cases` ricostruiti dalle sorgenti con quelli
già presenti in un workbook chiuso.

È il test di regressione del piano: il risultato di ieri è la specifica. Le
sorgenti e il loro risultato sono entrambi nel repo, quindi non serve indovinare
la semantica della trasformazione — si verifica.

  python tools/golden_turni.py \\
      --workbook "samples/omni-report/Omni Report W30.xlsm" \\
      --roster   "samples/omni-report/sorgenti/Turni_W30.xlsx" \\
      --backoffice "samples/omni-report/sorgenti/Back_Office_Time_Final.xlsx" \\
      --monday 46223

Differenze **attese** (dichiarate, non scoperte a posteriori):
  - righe FERIE-OFF: celle vuote invece della spazzatura del processo manuale
    (`Ore/gg`=8, `Inizio turno`=1447) — nessuna formula le legge, filtrano su LAVORA
  - `Stato BO`: `NO BOT` esplicito dove il processo manuale lascia vuoto
  - con --include-marked: l'agente con skill marcata entra in `Turni`

Qualunque altra differenza è un bug del parser.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from fasterreports.core.wfmsource import (  # noqa: E402
    read_alias_map,
    read_backoffice,
    read_roster,
    week_bounds,
)
from fasterreports.core.xlsxsource import read_sheet  # noqa: E402

TOL = 1e-9


def _num(v):
    if v is None or str(v).strip() == "":
        return None
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        return str(v).strip()


def _blank(v) -> bool:
    """None e cella assente sono la stessa cosa: entrambe 'vuoto'."""
    return v is None or (isinstance(v, str) and not v.strip())


def _same(a, b) -> bool:
    if _blank(a) and _blank(b):
        return True
    if _blank(a) or _blank(b):
        return False
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) < TOL
    return str(a).strip() == str(b).strip()


def read_workbook_turni(path: Path) -> dict[tuple[str, int], dict]:
    sheet = read_sheet(path, "Turni")
    out = {}
    for r in sorted(sheet.rows):
        if r == 1:
            continue
        row = sheet.row(r)
        nome = str(row.get("A", "")).strip()
        if not nome:
            continue
        data = _num(row.get("E"))
        out[(nome, int(data))] = {
            "Team/Skill": str(row.get("B", "")).strip(),
            "Contratto": str(row.get("C", "")).strip(),
            "Ore/gg": _num(row.get("D")),
            "Stato": str(row.get("F", "")).strip(),
            "Inizio turno": _num(row.get("G")),
            "Fine turno": _num(row.get("H")),
            "chiave": str(row.get("I", "")).strip(),
        }
    return out


def read_workbook_slot(path: Path) -> dict[tuple[str, int], dict]:
    sheet = read_sheet(path, "Slot Only Cases")
    out = {}
    for r in sorted(sheet.rows):
        if r == 1:
            continue
        row = sheet.row(r)
        key = str(row.get("A", "")).strip()
        if not key:
            continue
        out[(key, int(_num(row.get("B"))))] = {
            "Slot inizio": _num(row.get("C")),
            "Slot fine": _num(row.get("D")),
            "Stato BO": str(row.get("E", "")).strip(),
        }
    return out


def compare_strict(name, mine, theirs, fields, allowed) -> int:
    """Confronto in cui l'attesa è un **predicato sulla riga**, non un campo.

    Bucketare per campo nasconde le sorprese: se 101 righe FERIE-OFF hanno
    `Ore/gg` diverso per un motivo noto e una 102ª riga LAVORA ce l'ha diverso
    per un bug, un conteggio per campo le mescola. Qui ogni differenza va
    giustificata dalla riga in cui si trova.

    `allowed[campo]` è `(descrizione, predicato(riga_ricostruita) -> bool)`.
    """
    print(f"\n{'='*72}\n### {name}")
    print(f"  ricostruite: {len(mine)}   nel workbook: {len(theirs)}")
    only_mine = sorted(set(mine) - set(theirs))
    only_theirs = sorted(set(theirs) - set(mine))
    if only_mine:
        print(f"  SOLO ricostruite: {len(only_mine)} righe, "
              f"agenti: {sorted({k[0] for k in only_mine})}")
    if only_theirs:
        print(f"  SOLO nel workbook: {len(only_theirs)} righe, "
              f"agenti: {sorted({k[0] for k in only_theirs})}")

    accounted: dict[str, int] = {}
    unexpected: list[tuple] = []
    common = sorted(set(mine) & set(theirs))
    for key in common:
        for f in fields:
            a, b = mine[key].get(f), theirs[key].get(f)
            if _same(a, b):
                continue
            rule = allowed.get(f)
            if rule and rule[1](mine[key]):
                accounted[f] = accounted.get(f, 0) + 1
            else:
                unexpected.append((f, key, a, b))

    print(f"  righe in comune: {len(common)}")
    for f, n in sorted(accounted.items()):
        print(f"  [attesa] {f}: {n} differenze — {allowed[f][0]}")
    if unexpected:
        print(f"  ** INATTESE **: {len(unexpected)}")
        for f, key, a, b in unexpected[:10]:
            print(f"      {f} {key}: ricostruito={a!r} workbook={b!r}")
    elif not accounted:
        print("  nessuna differenza")
    return len(unexpected) + len(only_mine) + len(only_theirs)


def compare(name, mine, theirs, fields, expected_diff) -> int:
    print(f"\n{'='*72}\n### {name}")
    print(f"  ricostruite: {len(mine)}   nel workbook: {len(theirs)}")
    only_mine = sorted(set(mine) - set(theirs))
    only_theirs = sorted(set(theirs) - set(mine))
    if only_mine:
        agents = sorted({k[0] for k in only_mine})
        print(f"  SOLO ricostruite: {len(only_mine)} righe, agenti: {agents}")
    if only_theirs:
        agents = sorted({k[0] for k in only_theirs})
        print(f"  SOLO nel workbook: {len(only_theirs)} righe, agenti: {agents}")

    diffs: dict[str, list] = {}
    for key in sorted(set(mine) & set(theirs)):
        for f in fields:
            a, b = mine[key].get(f), theirs[key].get(f)
            if not _same(a, b):
                diffs.setdefault(f, []).append((key, a, b))

    common = len(set(mine) & set(theirs))
    print(f"  righe in comune: {common}")
    unexpected = 0
    for f, items in sorted(diffs.items()):
        tag = expected_diff.get(f)
        n = len(items)
        pct = f"{n/common:.0%}" if common else "-"
        if tag:
            print(f"  [attesa] {f}: {n} differenze ({pct}) — {tag}")
            for key, a, b in items[:2]:
                print(f"      {key}: ricostruito={a!r} workbook={b!r}")
        else:
            unexpected += n
            print(f"  ** INATTESA ** {f}: {n} differenze ({pct})")
            for key, a, b in items[:6]:
                print(f"      {key}: ricostruito={a!r} workbook={b!r}")
    if not diffs:
        print("  nessuna differenza sui campi confrontati")
    return unexpected


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("--workbook", type=Path, required=True)
    ap.add_argument("--roster", type=Path, required=True)
    ap.add_argument("--backoffice", type=Path, required=True)
    ap.add_argument("--monday", type=int, required=True, help="seriale Excel del lunedì")
    ap.add_argument("--skills", nargs="*", default=["HPO"])
    ap.add_argument("--include-marked", action="store_true")
    args = ap.parse_args(argv)

    week = week_bounds(args.monday)
    print(f"Settimana: {week[0]} .. {week[1]}  (seriali {args.monday}..{args.monday+6})")

    aliases = read_alias_map(args.workbook)
    print(f"Alias nomi da 'Helper Malpractice'!D:E: {len(aliases)}")
    for k, v in sorted(aliases.items()):
        print(f"    {k!r} -> {v!r}")

    roster = read_roster(
        args.roster, week=week, skills=tuple(args.skills),
        include_marked=args.include_marked,
    )
    allowed = set(roster.notes.target_agents)
    back = read_backoffice(
        args.backoffice, week=week, aliases=aliases, allowed_keys=allowed,
    )

    print(f"\nAgenti con skill richiesta nel roster: {len(allowed)}")
    if roster.notes.skills_with_marker:
        print(f"  skill marcate: {roster.notes.skills_with_marker}")
    dup = {k: v for k, v in roster.notes.agents_in_blocks.items() if len(set(v)) > 1}
    if dup:
        print(f"  ! agenti con skill richiesta in più blocchi: {dup}")
    if back.notes.aliases_applied:
        print(f"  alias applicati: {back.notes.aliases_applied}")
    if back.notes.keys_not_allowed:
        print(f"  agenti del back office esclusi (non nel target): "
              f"{sorted(back.notes.keys_not_allowed)}")

    mine_t = {
        (r[0], r[4]): {
            "Team/Skill": r[1], "Contratto": r[2], "Ore/gg": r[3],
            "Stato": r[5], "Inizio turno": r[6], "Fine turno": r[7], "chiave": r[8],
        }
        for r in roster.data
    }
    mine_s = {
        (r[0], r[1]): {"Slot inizio": r[2], "Slot fine": r[3], "Stato BO": r[4]}
        for r in back.data
    }

    # Le differenze ammesse valgono SOLO sulle righe che le giustificano.
    non_lavora = ("solo sulle righe non-LAVORA: celle vuote invece della "
                  "spazzatura del processo manuale",
                  lambda r: r["Stato"] != "LAVORA")
    non_bot = ("solo sulle righe senza slot (NO BOT / REQUEST)",
               lambda r: r["Slot inizio"] is None)

    bad = 0
    bad += compare_strict(
        "Turni", mine_t, read_workbook_turni(args.workbook),
        ["Team/Skill", "Ore/gg", "Stato", "Inizio turno", "Fine turno", "chiave"],
        {"Ore/gg": non_lavora, "Inizio turno": non_lavora, "Fine turno": non_lavora,
         "chiave": non_lavora},
    )
    bad += compare_strict(
        "Slot Only Cases", mine_s, read_workbook_slot(args.workbook),
        ["Slot inizio", "Slot fine", "Stato BO"],
        {"Slot inizio": non_bot, "Slot fine": non_bot,
         "Stato BO": ("'NO BOT' esplicito dove il manuale lascia vuoto", non_bot[1])},
    )

    print(f"\n{'='*72}")
    if bad:
        print(f"ESITO: {bad} differenze INATTESE -> il parser non è fedele.")
        return 1
    print("ESITO: nessuna differenza inattesa. Trasformazione fedele al processo manuale.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
