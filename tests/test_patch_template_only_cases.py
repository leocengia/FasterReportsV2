"""Lo strumento che rinomina il foglio e allarga il tetto, senza aprire Excel.

Stessa asimmetria dello strumento dei duplicati: qui non si legge un workbook, si
RISCRIVE — e un errore non produce un numero sbagliato, produce un file che Excel
non apre, scoperto sulla macchina con Excel a giro iniziato.

Le proprieta' che contano:

- **i due numeri della FILTER si muovono insieme.** `SEQUENCE(59999,1,2)` e
  `AT_DATASET!$F$2:$F$60000` sono la stessa cosa detta due volte. Allargare solo
  l'intervallo farebbe confrontare a `FILTER` due array di lunghezza diversa,
  cioe' `#VALUE!` al primo ricalcolo — e il foglio si svuoterebbe. Lo strumento
  rifiuta invece di scrivere;
- **e' idempotente**, come quello dei duplicati;
- **il rename e' completo**. Il nome del foglio sta in tre punti; dimenticarne
  uno lascia il pacchetto incoerente, e Excel se ne accorge prima di noi.
"""

from __future__ import annotations

import re
import shutil
import sys
import zipfile
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

from patch_template_only_cases import (  # noqa: E402
    NOME_NUOVO,
    NOME_VECCHIO,
    RIGA_NUOVA,
    RIGA_VECCHIA,
    allarga_tetto,
    main,
    rinomina_foglio,
)

TEMPLATE = ROOT / "template" / "Omni_Report_TEMPLATE.xlsm"

GUIDA = (
    "_xlfn.LET(_xlpm.stato,'{nome}'!$B$5,_xlpm.r,_xlfn.SEQUENCE({seq},1,2),"
    "_xlpm.s,AT_DATASET!$F$2:$F${riga},_xlfn._xlws.FILTER(_xlpm.r,_xlpm.s=_xlpm.stato,\"\"))"
)


# ---------------------------------------------------------------------------
# PATCH 2 — il tetto
# ---------------------------------------------------------------------------


def test_intervallo_e_sequenza_si_spostano_insieme():
    prima = GUIDA.format(nome=NOME_NUOVO, seq=RIGA_VECCHIA - 1, riga=RIGA_VECCHIA)
    dopo, n_int, n_seq = allarga_tetto(prima)
    assert (n_int, n_seq) == (1, 1)
    assert dopo == GUIDA.format(nome=NOME_NUOVO, seq=RIGA_NUOVA - 1, riga=RIGA_NUOVA)


def test_allarga_solo_gli_intervalli_che_partono_da_riga_2():
    """Un intervallo che parte altrove non e' "tutto il dataset": e' una scelta.

    `AT_DATASET!$P$500:$P$60000` sarebbe una selezione fatta a mano — allargarla
    ne cambierebbe il senso, non il margine.
    """
    testo = "MINIFS(AT_DATASET!$P$500:$P$60000,AT_DATASET!$F$2:$F$60000,$B$5)"
    dopo, n_int, _ = allarga_tetto(testo)
    assert n_int == 1
    assert "$P$500:$P$60000" in dopo
    assert f"$F$2:$F${RIGA_NUOVA}" in dopo


def test_non_tocca_gli_altri_dataset():
    """`Turni` e `Slot Only Cases` hanno limiti loro, e restano dove sono."""
    testo = "SUMIFS('Slot Only Cases'!$C$2:$C$2000,Turni!$E$2:$E$5000,$D5)"
    dopo, n_int, n_seq = allarga_tetto(testo)
    assert (dopo, n_int, n_seq) == (testo, 0, 0)


def test_su_un_testo_gia_patchato_non_fa_niente():
    gia = GUIDA.format(nome=NOME_NUOVO, seq=RIGA_NUOVA - 1, riga=RIGA_NUOVA)
    assert allarga_tetto(gia) == (gia, 0, 0)


# ---------------------------------------------------------------------------
# PATCH 1 — il rename
# ---------------------------------------------------------------------------


def _parti_finte() -> dict[str, bytes]:
    return {
        "xl/workbook.xml": (
            f'<workbook><sheets><sheet name="OC Eventi" sheetId="1" r:id="rId1"/>'
            f'<sheet name="{NOME_VECCHIO}" sheetId="2" r:id="rId2"/>'
            f"</sheets></workbook>"
        ).encode("utf8"),
        "xl/_rels/workbook.xml.rels": (
            b'<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/>'
            b'<Relationship Id="rId2" Target="worksheets/sheet2.xml"/></Relationships>'
        ),
        "xl/worksheets/sheet1.xml": (
            f"<worksheet><f>'{NOME_VECCHIO}'!$B$5</f></worksheet>"
        ).encode("utf8"),
        "xl/worksheets/sheet2.xml": b"<worksheet/>",
        "docProps/app.xml": (
            f"<Properties><vt:lpstr>{NOME_VECCHIO}</vt:lpstr></Properties>"
        ).encode("utf8"),
    }


def test_il_rename_tocca_tutti_e_tre_i_punti():
    parti = _parti_finte()
    fatte = rinomina_foglio(parti)
    assert len(fatte) == 3
    unito = b"".join(parti.values())
    assert NOME_VECCHIO.encode() not in unito
    assert unito.count(NOME_NUOVO.encode()) == 3


def test_il_rename_e_idempotente():
    parti = _parti_finte()
    rinomina_foglio(parti)
    assert rinomina_foglio(parti) == []


def test_senza_il_foglio_vecchio_non_inventa_niente():
    parti = _parti_finte()
    parti["xl/workbook.xml"] = b'<workbook><sheets><sheet name="Altro"/></sheets></workbook>'
    assert rinomina_foglio(parti) == []


# ---------------------------------------------------------------------------
# Sul template tracciato
# ---------------------------------------------------------------------------


@pytest.mark.skipif(not TEMPLATE.is_file(), reason="template assente")
def test_il_template_tracciato_e_gia_patchato(tmp_path):
    """Rilanciare la patch sul template del repo non trova niente da fare.

    E' il modo piu' economico di verificare che il file sia nello stato che il
    codice si aspetta: se un domani il template venisse rifatto a partire da un
    export nuovo, dimenticando le due patch, questo test lo direbbe — e con lui
    quello sul nome del foglio in test_only_cases.py.
    """
    copia = tmp_path / "t.xlsm"
    shutil.copy2(TEMPLATE, copia)
    assert main([str(copia), "--dry-run"]) == 0

    with zipfile.ZipFile(copia) as z:
        for n in z.namelist():
            raw = z.read(n)
            assert NOME_VECCHIO.encode() not in raw, f"{n} cita ancora il nome vecchio"
            if n.startswith("xl/worksheets/"):
                testo = raw.decode("utf8", "replace")
                assert not re.search(
                    rf"AT_DATASET!\$[A-Z]{{1,3}}\$2:\$[A-Z]{{1,3}}\${RIGA_VECCHIA}\b",
                    testo,
                ), (
                    f"{n} legge AT_DATASET solo fino a riga {RIGA_VECCHIA}: lancia\n"
                    f"  python tools/patch_template_only_cases.py "
                    f"template/Omni_Report_TEMPLATE.xlsm"
                )
