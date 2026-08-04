"""I limiti di riga scritti a mano nelle formule.

E' l'unico guasto di questo progetto che arriva **da solo**: non serve che
qualcuno tocchi niente, basta che il volume dei dati cresca. Una formula che
legge `AT_DATASET!$P$2:$P$130000` funziona per anni e poi, la settimana in cui le
righe diventano 130001, smette di vedere l'eccedenza — senza un errore, con medie
e conteggi su un sottoinsieme.
"""

from __future__ import annotations

import pytest

from fasterreports.core.coherence import BLOCCA, SEGNALA, check_sources
from fasterreports.core.errors import SourceError
from fasterreports.core.templatescan import (
    RowLimit,
    binding_limits,
    scan_error_cells,
    scan_row_limits,
)
from xlsxbuild import make_workbook

DS = ("AT_DATASET", "SF_DATABASE", "PSAT_DATASET", "Turni")


def _finding(rep, frammento):
    for f in rep.findings:
        if frammento in f.check:
            return f
    return None


# ---------------------------------------------------------------------------
# La scansione
# ---------------------------------------------------------------------------


def test_trova_un_intervallo_con_riga_finale(tmp_path):
    p = make_workbook(
        tmp_path / "t.xlsx",
        [
            ("Report Agenti", {"A1": "=SUMIFS(AT_DATASET!$P$2:$P$130000,A2,1)"}),
            ("AT_DATASET", {"A1": "x"}),
        ],
    )
    (lim,) = scan_row_limits(p, DS)
    assert lim.dataset == "AT_DATASET"
    assert lim.max_row == 130000
    assert lim.sheet == "Report Agenti"
    assert lim.ref == "P2:P130000"


def test_colonna_intera_non_e_un_limite(tmp_path):
    """`AT_DATASET!$K:$K` non si ferma da nessuna parte: e' la forma robusta, e
    segnalarla sarebbe rumore."""
    p = make_workbook(
        tmp_path / "t.xlsx",
        [("Report Agenti", {"A1": "=SUMIFS(AT_DATASET!$K:$K,AT_DATASET!$B:$B,1)"})],
    )
    assert scan_row_limits(p, DS) == []


def test_riferimento_a_tabella_non_e_un_limite(tmp_path):
    """La tabella viene ridimensionata ai dati a ogni build."""
    p = make_workbook(
        tmp_path / "t.xlsx",
        [("Verifica AHT", {"A1": '=COUNTIFS(AHT_Data[Employee Name],"x")'})],
    )
    assert scan_row_limits(p, DS) == []


def test_nome_di_foglio_fra_apici(tmp_path):
    p = make_workbook(
        tmp_path / "t.xlsx",
        [("Helper", {"A1": "=COUNTA('Slot Only Cases'!$A$2:$A$1291)"})],
    )
    (lim,) = scan_row_limits(p, ("Slot Only Cases",))
    assert lim.max_row == 1291


def test_un_dataset_non_prende_il_limite_di_un_altro(tmp_path):
    """`AT_DATASET` e' sottostringa di `PSAT_DATASET`: senza confine a sinistra
    il limite di PSAT verrebbe attribuito anche ad AT."""
    p = make_workbook(
        tmp_path / "t.xlsx",
        [("R", {"A1": "=COUNTA(PSAT_DATASET!$I$2:$I$1000)"})],
    )
    lims = scan_row_limits(p, DS)
    assert [l.dataset for l in lims] == ["PSAT_DATASET"]


def test_stesso_riferimento_in_piu_formule_conta_una_volta(tmp_path):
    p = make_workbook(
        tmp_path / "t.xlsx",
        [
            ("R", {
                "A1": "=COUNTA(Turni!$A$2:$A$10000)",
                "A2": "=COUNTA(Turni!$A$2:$A$10000)",
                "A3": "=COUNTA(Turni!$B$2:$B$10000)",
            }),
        ],
    )
    refs = {l.ref for l in scan_row_limits(p, DS)}
    assert refs == {"A2:A10000", "B2:B10000"}


def test_template_assente(tmp_path):
    with pytest.raises(SourceError):
        scan_row_limits(tmp_path / "manca.xlsm", DS)


def test_binding_limits_prende_il_piu_basso():
    lims = [
        RowLimit("SF_DATABASE", 50000, "Anagrafica", "BB2:BB50000"),
        RowLimit("SF_DATABASE", 3389, "Profilo Colonne SF", "A2:EM3389"),
        RowLimit("Turni", 10000, "Helper Turni", "A2:A10000"),
    ]
    vinc = binding_limits(lims)
    assert vinc["SF_DATABASE"].max_row == 3389
    assert vinc["Turni"].max_row == 10000


# ---------------------------------------------------------------------------
# Il controllo
# ---------------------------------------------------------------------------


def test_dati_oltre_il_limite_bloccano():
    lims = [RowLimit("SF_DATABASE", 3389, "Profilo Colonne SF", "A2:EM3389")]
    rep = check_sources(row_limits=lims, last_rows={"SF_DATABASE": 4000})
    f = _finding(rep, "troppo corte")
    assert f is not None and f.level == BLOCCA
    assert "4000" in f.summary and "3389" in f.summary
    assert "Profilo Colonne SF" in f.details[0]
    assert "sottoinsieme" in f.hint
    assert not rep.ok


def test_dati_vicini_al_limite_segnalano():
    """Segnalare PRIMA e' il punto: quando mordera' non lo dira' nessuno."""
    lims = [RowLimit("SF_DATABASE", 3389, "Profilo Colonne SF", "A2:EM3389")]
    rep = check_sources(row_limits=lims, last_rows={"SF_DATABASE": 2945})
    f = _finding(rep, "vicine al limite")
    assert f is not None and f.level == SEGNALA
    assert "87%" in f.summary
    assert rep.ok


def test_margine_ampio_non_dice_nulla():
    lims = [RowLimit("PSAT_DATASET", 1000, "Recap PSAT Positive", "I2:I1000")]
    rep = check_sources(row_limits=lims, last_rows={"PSAT_DATASET": 23})
    assert _finding(rep, "limite") is None


def test_senza_template_il_controllo_si_salta():
    """Non si inventa un allarme su un dato che non si ha."""
    rep = check_sources(row_limits=None, last_rows={"SF_DATABASE": 999999})
    assert _finding(rep, "limite") is None
    assert _finding(rep, "troppo corte") is None


def test_dataset_senza_righe_non_segnala():
    lims = [RowLimit("Turni", 10000, "Helper Turni", "A2:A10000")]
    rep = check_sources(row_limits=lims, last_rows={"Turni": 0})
    assert _finding(rep, "limite") is None


# ---------------------------------------------------------------------------
# Il template vero
# ---------------------------------------------------------------------------


def test_il_template_reale_ha_i_limiti_che_conosciamo():
    """Fotografia dei limiti veri, così una modifica al template che li cambia si
    vede qui invece di essere scoperta da una settimana piena."""
    from pathlib import Path

    tpl = Path(__file__).resolve().parents[1] / "template" / "Omni_Report_TEMPLATE.xlsm"
    if not tpl.is_file():
        pytest.skip("template assente")

    lims = scan_row_limits(
        tpl,
        ("AT_DATASET", "ATwi_DATASET", "SF_DATABASE", "PSAT_DATASET",
         "Turni", "Slot Only Cases"),
    )
    vinc = {ds: l.max_row for ds, l in binding_limits(lims).items()}

    # ATwi e Slot Only Cases sono letti solo per colonna intera o dal VBA:
    # nessun limite, ed e' la situazione da preferire.
    assert "ATwi_DATASET" not in vinc
    assert "Slot Only Cases" not in vinc

    # Gli altri hanno margini ampi, tranne SF_DATABASE — che e' il motivo per cui
    # esiste tools/fix_profilo_colonne_sf.py. Quando quel fix sara' applicato al
    # template, SF_DATABASE salira' a 50000 (Anagrafica) e questa riga va aggiornata.
    assert vinc["AT_DATASET"] == 130000
    assert vinc["Turni"] == 10000
    assert vinc["PSAT_DATASET"] == 1000
    assert vinc["SF_DATABASE"] in (3389, 50000), (
        f"SF_DATABASE limitato a {vinc['SF_DATABASE']}: se hai cambiato le formule "
        f"di 'Profilo Colonne SF', aggiorna questo test."
    )


# ---------------------------------------------------------------------------
# Le celle di errore nel workbook prodotto
# ---------------------------------------------------------------------------


def _con_errori(tmp_path, celle):
    """Un .xlsx con celle di errore (t="e"), che il costruttore non produce."""
    import zipfile

    p = make_workbook(tmp_path / "e.xlsx", [("Foglio", {"A1": 1})])
    with zipfile.ZipFile(p) as z:
        parti = {n: z.read(n) for n in z.namelist()}
    corpo = "".join(
        f'<row r="{r}"><c r="A{r}" t="e"><v>{v}</v></c></row>'
        for r, v in celle
    )
    parti["xl/worksheets/sheet1.xml"] = (
        '<?xml version="1.0"?><worksheet><sheetData>' + corpo + "</sheetData></worksheet>"
    ).encode("utf-8")
    with zipfile.ZipFile(p, "w") as z:
        for n, data in parti.items():
            z.writestr(n, data)
    return p


def test_trova_le_celle_di_errore(tmp_path):
    p = _con_errori(tmp_path, [(1, "#SPILL!"), (2, "#SPILL!"), (3, "#REF!")])
    (err,) = scan_error_cells(p)
    assert err.sheet == "Foglio"
    assert err.kinds == {"#SPILL!": 2, "#REF!": 1}
    assert err.total == 3
    assert err.examples[0] == "A1"


def test_workbook_senza_errori(tmp_path):
    p = make_workbook(tmp_path / "ok.xlsx", [("Foglio", {"A1": 1, "B1": "x"})])
    assert scan_error_cells(p) == []


def test_ordinati_per_gravita(tmp_path):
    """Il foglio con piu' errori per primo: di solito e' la causa, gli altri sono
    conseguenze sue."""
    import zipfile

    p = make_workbook(
        tmp_path / "e.xlsx", [("Poco", {"A1": 1}), ("Molto", {"A1": 1})]
    )
    with zipfile.ZipFile(p) as z:
        parti = {n: z.read(n) for n in z.namelist()}
    parti["xl/worksheets/sheet1.xml"] = (
        '<?xml version="1.0"?><worksheet><sheetData>'
        '<row r="1"><c r="A1" t="e"><v>#REF!</v></c></row>'
        "</sheetData></worksheet>"
    ).encode("utf-8")
    parti["xl/worksheets/sheet2.xml"] = (
        '<?xml version="1.0"?><worksheet><sheetData>'
        + "".join(
            f'<row r="{r}"><c r="A{r}" t="e"><v>#VALUE!</v></c></row>'
            for r in range(1, 6)
        )
        + "</sheetData></worksheet>"
    ).encode("utf-8")
    with zipfile.ZipFile(p, "w") as z:
        for n, data in parti.items():
            z.writestr(n, data)

    err = scan_error_cells(p)
    assert [e.sheet for e in err] == ["Molto", "Poco"]


def test_workbook_assente(tmp_path):
    with pytest.raises(SourceError):
        scan_error_cells(tmp_path / "manca.xlsm")


def test_build_result_non_e_ok_con_celle_di_errore():
    """Il file c'e' ma i suoi numeri no: `ok` deve dirlo, e la CLI uscire != 0."""
    from fasterreports.core.preflight import PreflightReport
    from fasterreports.core.templatescan import ErrorCells
    from fasterreports.omni.orchestrate import BuildResult
    from pathlib import Path

    rep = PreflightReport()
    pulito = BuildResult(workbook=Path("x.xlsm"), preflight=Path("p.txt"), report=rep)
    assert pulito.ok and pulito.n_errors == 0

    sporco = BuildResult(
        workbook=Path("x.xlsm"),
        preflight=Path("p.txt"),
        report=rep,
        error_cells=[ErrorCells("AHT Outliers", {"#REF!": 73}, ("L3",))],
    )
    assert not sporco.ok
    assert sporco.n_errors == 73
