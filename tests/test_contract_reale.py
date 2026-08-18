"""Il contratto contro il workbook vero.

Questi test sono l'auto-verifica del piano §11: invece di fidarsi della tabella
scritta nel documento, si legge il file e si controlla che le lettere combacino.
Quando il workbook evolve — e nel W30 e' successo, tre colonne in piu' rispetto
al piano — sono questi test a dirlo.

La fixture si rigenera con:
  python tools/audit_workbook.py template/Omni_Report_TEMPLATE.xlsm \\
      --headers --usage --tables --json > tests/fixtures/workbook_template.json
"""

from __future__ import annotations

import pytest

from fasterreports.core.matcher import VIA_EXACT, resolve_dataset

# Colonne che compaiono nelle formule ma non sono dati di input.
#
# Vuoto dal 2026-08-18: l'unica voce era SF_DATABASE!A, che 'Profilo Colonne SF'
# usava come angolo di un INDEX su tutto il foglio invece che come campo. Quel
# foglio e' stato cancellato dal template (diagnostico, nessun lettore, 143
# formule ad array a ogni ricalcolo), e nel frattempo la colonna A e' diventata
# un campo vero: 'Date Viewpoint'. Tenere l'esenzione la nasconderebbe.
IGNORE_USAGE: set[tuple[str, str]] = set()


def test_ogni_campo_aggancia_la_colonna_giusta(contract, real_headers):
    """Il cuore: risolvere gli header reali deve dare le lettere del contratto.

    Se qualcuno rinomina una colonna nell'export, o se il contratto punta alla
    colonna sbagliata, si vede qui e non in una media sbagliata a valle.
    """
    for name, dataset in contract.datasets.items():
        info = real_headers["headers"].get(name)
        assert info, f"fixture senza il foglio {name}"

        # Gli header del foglio, nell'ordine delle colonne: e' quello che
        # produrrebbe un export CSV coerente col workbook.
        cols = info["columns"]
        ordered = [cols[c] for c in sorted(cols, key=_col_index)]

        mapping = resolve_dataset(dataset, ordered, contract.normalize)

        # La colonna sorgente agganciata deve stare nella lettera attesa: il
        # foglio ha gli header nella loro posizione canonica.
        letters = sorted(cols, key=_col_index)
        for r in mapping.resolutions:
            letter = letters[r.source_index]
            assert letter == r.target_col, (
                f"{name}: {r.canonical!r} agganciata alla colonna {letter} del foglio, "
                f"ma il contratto dice {r.target_col}"
            )


def test_le_colonne_a_rischio_si_risolvono_esattamente(contract, real_headers):
    """Le collisioni note vanno risolte per nome esatto, non per fortuna."""
    a_rischio = [
        ("PSAT_DATASET", "Agent Name"),   # vs agent_name (BO)
        ("PSAT_DATASET", "psat_score"),   # vs PSAT Score (EK)
        ("SF_DATABASE", "Case Origin (group)"),  # vs Case Origin (T)
        ("SF_DATABASE", "Case AHT (mins)"),      # vs le altre colonne AHT
    ]
    for ds_name, canonical in a_rischio:
        dataset = contract.dataset(ds_name)
        cols = real_headers["headers"][ds_name]["columns"]
        ordered = [cols[c] for c in sorted(cols, key=_col_index)]
        mapping = resolve_dataset(dataset, ordered, contract.normalize)
        res = mapping.by_canonical()[canonical]
        assert res.via == VIA_EXACT, (
            f"{ds_name}.{canonical} risolta via {res.via}: per una colonna con "
            f"omonimi serve il match esatto."
        )


def test_ogni_colonna_consumata_da_formule_e_nel_contratto(contract, real_headers):
    """Se una formula legge una colonna che il contratto non copre, la pipeline
    la lascerebbe vuota e il numero a valle sarebbe sbagliato in silenzio."""
    mancanti: list[str] = []
    for ds_name, cols in real_headers["usage"].items():
        dataset = contract.datasets.get(ds_name)
        if dataset is None:
            continue
        coperte = {f.target_col for f in dataset.fields}
        for col, sheets in cols.items():
            if (ds_name, col) in IGNORE_USAGE:
                continue
            if col not in coperte:
                mancanti.append(f"{ds_name}!{col} (letta da: {', '.join(sheets)})")

    assert not mancanti, (
        "Colonne consumate da formule ma assenti da config/columns.yml:\n  "
        + "\n  ".join(mancanti)
        + "\n\nAggiungile al contratto, altrimenti la pipeline le lascia vuote."
    )


def test_le_tre_colonne_trovate_nellaudit_sono_nel_contratto(contract):
    """Regressione esplicita sul delta rispetto al piano §4.

    Trovate scansionando le formule del W30; se qualcuno le rimuove dal
    contratto pensando che il piano sia completo, questo test lo ferma.
    """
    attese = [
        ("AT_DATASET", "H", "Number of Active Contacts"),      # Verifica AHT
        ("ATwi_DATASET", "J", "Handle Time in seconds"),       # Verifica AHT
        ("PSAT_DATASET", "AU", "Survey Date (Exact)"),         # Recap PSAT Positive
        # Trovate il 2026-08-04, ma non da questo test: dal foglio 'AHT Outliers'
        # uscito VUOTO nel primo W31 generato. Le sue formule le citano come
        # riferimento strutturato (`AHT_Data[Case Type]`), e l'audit che genera
        # la fixture cercava solo `SF_DATABASE!$Z`: il test girava su un elenco
        # incompleto e quindi taceva. Ora l'audit risolve anche i riferimenti
        # strutturati, e questo test le vede.
        ("SF_DATABASE", "Z", "Case Type"),                     # AHT Outliers
        ("SF_DATABASE", "BV", "Primary Category"),              # AHT Outliers
    ]
    for ds_name, letter, canonical in attese:
        ds = contract.dataset(ds_name)
        f = next((x for x in ds.fields if x.canonical == canonical), None)
        assert f is not None, f"{ds_name}: manca il campo {canonical!r}"
        assert f.target_col == letter


def test_la_tabella_AHT_Data_e_dichiarata_nel_contratto(contract, real_headers):
    """SF_DATABASE ospita un ListObject: il writer deve saperlo per ridimensionarlo."""
    tabelle = {t["name"]: t for t in real_headers["tables"]}
    assert "AHT_Data" in tabelle
    assert tabelle["AHT_Data"]["sheet"] == "SF_DATABASE"
    assert contract.dataset("SF_DATABASE").list_object == "AHT_Data"


def test_i_campi_derivati_hanno_una_formula_documentata(contract):
    for ds in contract.datasets.values():
        for f in ds.derived_fields:
            assert f.formula, f"{ds.name}.{f.canonical}: campo derivato senza formula"
            assert f.source, f"{ds.name}.{f.canonical}: campo derivato senza source"


def _col_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n
