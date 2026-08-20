"""Lo strumento che patcha il template senza aprire Excel.

Va testato piu' di altri, per una ragione asimmetrica: qui non si legge un
workbook, si RISCRIVE — e un errore non produce un numero sbagliato, produce un
file che Excel non apre. Il difetto si scoprirebbe sulla macchina di Leonardo, a
giro iniziato.

Due proprieta' contano piu' delle altre:

- **quando non riconosce, non tocca.** La prima esecuzione vera ha rifiutato
  tutte e 34 le formule (il regex cercava `&quot;` dove l'XML ha `"`), e ha
  lasciato il file intatto dicendolo. E' andata come deve andare;
- **e' idempotente.** Rilanciarlo su un template gia' patchato non deve
  raddoppiare niente.
"""

from __future__ import annotations

import re
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

from patch_template_duplicates import (  # noqa: E402
    main,
    nuova_formula,
    patch_duplicates_helper,
    togli_external_link,
)
from xlsxbuild import Condivisa, make_workbook  # noqa: E402

TEMPLATE = ROOT / "template" / "Omni_Report_TEMPLATE.xlsm"

VECCHIA_M = (
    'IF($K{r}="","",DATE(VALUE(MID(K{r},FIND("/",K{r},FIND("/",K{r})+1)+1,4)),'
    'VALUE(LEFT(K{r},FIND("/",K{r})-1)),VALUE(MID(K{r},FIND("/",K{r})+1,2))))'
)


def _foglio(tmp_path, grid):
    p = make_workbook(tmp_path / "helper.xlsx", [("Duplicates Helper", grid)])
    with zipfile.ZipFile(p) as z:
        return p, z.read("xl/worksheets/sheet1.xml").decode("utf8")


def test_la_formula_nuova_ha_la_colonna_assoluta():
    """`$K2`, come le altre colonne del foglio (`O`,`P`,`Q` usano `$M3`)."""
    assert nuova_formula("M", 2) == 'IF($K2="","",$K2)'
    assert nuova_formula("N", 963) == 'IF($L963="","",$L963)'


def test_riscrive_il_master_e_lascia_stare_l_erede(tmp_path):
    """Le celle eredi non hanno testo: ereditano la formula nuova da se'.

    E' il motivo per cui le sostituzioni sono 34 e non 2000 — e il motivo per cui
    toccare gli eredi sarebbe un errore.
    """
    _p, raw = _foglio(tmp_path, {
        "M2": "=" + VECCHIA_M.format(r=2),
        "M3": Condivisa(si=0, formula=VECCHIA_M.format(r=3), ref="M3:M66"),
        "M4": Condivisa(si=0, valore=46240.5),
    })
    nuovo, fatte, saltate = patch_duplicates_helper(raw)

    assert fatte == ["M2", "M3"], fatte
    assert saltate == []
    assert 'IF($K2="","",$K2)' in nuovo
    assert 'IF($K3="","",$K3)' in nuovo
    assert 'FIND("/"' not in nuovo
    # L'erede resta un erede: tag auto-chiuso, nessun testo aggiunto.
    assert '<f t="shared" si="0"/>' in nuovo
    # E il gruppo conserva il suo intervallo.
    assert 'ref="M3:M66"' in nuovo


def test_il_gruppo_condiviso_usa_la_riga_del_MASTER(tmp_path):
    """`ref="M67:M130"` -> la formula nuova deve dire `$K67`, non `$K2`.

    I riferimenti relativi delle celle eredi sono ancorati alla prima cella del
    gruppo: sbagliare quella riga sposterebbe di 65 righe tutte le date di quel
    blocco. Numeri plausibili, della persona sbagliata.
    """
    _p, raw = _foglio(tmp_path, {
        "M67": Condivisa(si=7, formula=VECCHIA_M.format(r=67), ref="M67:M130"),
    })
    nuovo, fatte, _ = patch_duplicates_helper(raw)
    assert fatte == ["M67"]
    assert 'IF($K67="","",$K67)' in nuovo
    assert "$K2" not in nuovo


def test_una_formula_inattesa_non_viene_toccata(tmp_path):
    """Quando non riconosce, non tocca — e lo dice."""
    _p, raw = _foglio(tmp_path, {"M2": "=SOMMA(K2:K9)"})
    nuovo, fatte, saltate = patch_duplicates_helper(raw)
    assert fatte == []
    assert len(saltate) == 1 and "M2" in saltate[0]
    assert nuovo == raw, "il foglio e' stato modificato pur non riconoscendo la formula"


def test_una_formula_che_legge_la_colonna_sbagliata_non_viene_toccata(tmp_path):
    """`M` deve leggere `K`. Se leggesse `L`, il foglio non e' quello che credo."""
    _p, raw = _foglio(tmp_path, {"M2": "=" + VECCHIA_M.format(r=2).replace("K", "L")})
    _nuovo, fatte, saltate = patch_duplicates_helper(raw)
    assert fatte == []
    assert saltate and "atteso K2" in saltate[0]


def test_e_idempotente(tmp_path):
    _p, raw = _foglio(tmp_path, {"M2": "=" + VECCHIA_M.format(r=2)})
    primo, fatte1, _ = patch_duplicates_helper(raw)
    secondo, fatte2, saltate2 = patch_duplicates_helper(primo)
    assert fatte1 == ["M2"]
    assert fatte2 == [], "la seconda passata ha riscritto qualcosa"
    assert secondo == primo
    # La formula nuova non e' riconosciuta come vecchia, quindi nemmeno segnalata
    # come "inattesa": e' semplicemente gia' a posto.
    assert len(saltate2) == 1 and "M2" in saltate2[0]


def test_togli_external_link_cancella_tutte_e_quattro_le_tracce():
    """Se ne resta una, Excel non apre il file."""
    parti = {
        "[Content_Types].xml": (
            b'<Types><Override PartName="/xl/externalLinks/externalLink1.xml" '
            b'ContentType="x"/><Override PartName="/xl/workbook.xml" ContentType="y"/></Types>'
        ),
        "xl/workbook.xml": (
            b"<workbook><sheets/><externalReferences><externalReference r:id=\"rId30\"/>"
            b"</externalReferences></workbook>"
        ),
        "xl/_rels/workbook.xml.rels": (
            b'<Relationships><Relationship Id="rId30" Type="http://schemas.openxmlformats.org'
            b'/officeDocument/2006/relationships/externalLink" Target="externalLinks/externalLink1.xml"/>'
            b'<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
        ),
        "xl/externalLinks/externalLink1.xml": b"<externalLink/>",
        "xl/externalLinks/_rels/externalLink1.xml.rels": b"<Relationships/>",
    }
    tolte = togli_external_link(parti)
    assert len(tolte) == 5, tolte
    assert not [n for n in parti if "external" in n]
    assert b"externalLink" not in parti["[Content_Types].xml"]
    assert b"rId30" not in parti["xl/_rels/workbook.xml.rels"]
    assert b"externalReferences" not in parti["xl/workbook.xml"]
    # E la relazione del foglio non e' stata toccata.
    assert b'Id="rId1"' in parti["xl/_rels/workbook.xml.rels"]


def test_su_un_pacchetto_senza_collegamenti_non_fa_niente():
    parti = {"xl/workbook.xml": b"<workbook/>"}
    assert togli_external_link(parti) == []
    assert parti == {"xl/workbook.xml": b"<workbook/>"}


@pytest.mark.skipif(not TEMPLATE.is_file(), reason="template assente")
def test_il_template_tracciato_e_gia_patchato(tmp_path):
    """Dopo la patch, rilanciarla non trova niente da fare.

    E' il modo piu' economico di verificare che il template nel repo sia nello
    stato che il contratto si aspetta: se un domani qualcuno rigenerasse il
    template dal WIP dimenticando la patch, questo test lo direbbe.
    """
    import shutil

    copia = tmp_path / "t.xlsm"
    shutil.copy2(TEMPLATE, copia)
    assert main([str(copia), "--dry-run"]) == 0

    with zipfile.ZipFile(copia) as z:
        assert not [n for n in z.namelist() if "external" in n], (
            "il template ha ancora il collegamento esterno"
        )
        for n in z.namelist():
            if n.startswith("xl/worksheets/"):
                raw = z.read(n).decode("utf8", "replace")
                assert 'FIND("/"' not in raw, (
                    f"{n} contiene ancora il parsing del testo delle date: "
                    f"lancia  python tools/patch_template_duplicates.py "
                    f"template/Omni_Report_TEMPLATE.xlsm"
                )
