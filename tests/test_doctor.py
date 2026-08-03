"""`omni-report check`: i controlli su ambiente e template.

Esistono perché un template preparato male fallisce durante il build, a Excel
già aperto e magari appeso su un MsgBox invisibile. Questi test verificano che il
check lo dica prima, e che distingua ciò che **blocca** da ciò che è solo da
sapere.
"""

from __future__ import annotations

import zipfile
from dataclasses import replace
from pathlib import Path

import pytest

from fasterreports.omni.doctor import ATTENZIONE, MANCA, OK, run_checks
from fasterreports.omni.settings import load_settings

ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "samples" / "omni-report" / "Omni Report W30.xlsm"


@pytest.fixture
def settings():
    return load_settings(ROOT / "config" / "settings.yml", root=ROOT)


def _check(rep, frammento):
    for c in rep.checks:
        if frammento in c.name:
            return c
    return None


def _run(contract, settings, template):
    return run_checks(contract, replace(settings, template=template), try_excel=False)


# --- template assente o sbagliato ------------------------------------------

def test_template_assente_blocca(contract, settings, tmp_path):
    rep = _run(contract, settings, tmp_path / "non_esiste.xlsm")
    c = _check(rep, "template presente")
    assert c.status == MANCA
    assert "COME_PREPARARE" in c.hint
    assert not rep.ok


def test_template_xlsx_blocca(contract, settings, tmp_path):
    """Salvato senza macro: il VBA e' il motore, senza non si va da nessuna parte."""
    p = tmp_path / "Omni_Report_TEMPLATE.xlsx"
    p.write_bytes(b"PK\x03\x04finto")
    rep = _run(contract, settings, p)
    c = _check(rep, "template con macro")
    assert c.status == MANCA
    assert "perso tutte le macro" in c.hint


def test_template_non_zip_blocca(contract, settings, tmp_path):
    p = tmp_path / "Omni_Report_TEMPLATE.xlsm"
    p.write_text("non sono uno zip", encoding="utf-8")
    rep = _run(contract, settings, p)
    assert _check(rep, "template leggibile").status == MANCA


def test_template_senza_vba_blocca(contract, settings, tmp_path):
    """Un .xlsm senza vbaProject.bin: rinominato a mano da un .xlsx."""
    p = tmp_path / "Omni_Report_TEMPLATE.xlsm"
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("xl/workbook.xml", "<workbook/>")
    rep = _run(contract, settings, p)
    c = _check(rep, "macro nel template")
    assert c.status == MANCA
    assert "Dettaglio Malpractice" in c.hint


# --- template vero, patch VBA mancante -------------------------------------

@pytest.mark.skipif(not SAMPLE.is_file(), reason="campione W30 assente")
def test_template_reale_struttura_ok_ma_patch_mancante(contract, settings):
    """Il W30 non patchato: struttura valida, ma i MsgBox bloccherebbero il run."""
    rep = _run(contract, settings, SAMPLE)

    assert _check(rep, "template presente").status == OK
    assert _check(rep, "macro nel template").status == OK
    assert _check(rep, "fogli attesi").status == OK
    assert _check(rep, "tabella alias").status == OK
    assert _check(rep, "Email Agenti").status == OK

    c = _check(rep, "SetSilentMode")
    assert c.status == MANCA
    assert "Alt+F11" in c.hint
    assert not rep.ok


@pytest.mark.skipif(not SAMPLE.is_file(), reason="campione W30 assente")
def test_alias_e_email_letti_dal_template(contract, settings):
    rep = _run(contract, settings, SAMPLE)
    assert "7 alias" in _check(rep, "tabella alias").detail
    assert "agenti con email" in _check(rep, "Email Agenti").detail


# --- fonti in input/ non bloccano -----------------------------------------

def test_fonti_mancanti_non_bloccano(contract, settings, tmp_path):
    """Le fonti si mettono al momento del run: segnalare basta."""
    s = replace(settings, input_dir=tmp_path, template=tmp_path / "assente.xlsm")
    rep = run_checks(contract, s, try_excel=False)
    c = _check(rep, "fonti in input")
    assert c.status == ATTENZIONE
    assert "0/6" in c.detail
    # E non e' fra i motivi di blocco.
    assert not c.blocking


def test_fonti_presenti_ok(contract, settings, tmp_path):
    for nome in settings.input_files.values():
        (tmp_path / nome).write_text("x", encoding="utf-8")
    s = replace(settings, input_dir=tmp_path, template=tmp_path / "assente.xlsm")
    rep = run_checks(contract, s, try_excel=False)
    assert _check(rep, "fonti in input").status == OK


# --- report ----------------------------------------------------------------

def test_excel_saltato_senza_prova(contract, settings, tmp_path):
    rep = _run(contract, settings, tmp_path / "assente.xlsm")
    assert _check(rep, "Excel raggiungibile").detail == "richiesto --no-excel"


def test_render_elenca_i_da_sistemare(contract, settings, tmp_path):
    rep = _run(contract, settings, tmp_path / "assente.xlsm")
    text = rep.render()
    assert "MANCA" in text
    assert "controlli da sistemare" in text
    assert "COME_PREPARARE_IL_TEMPLATE.md" in text


def test_render_dice_quando_e_pronto(contract, settings):
    from fasterreports.omni.doctor import CheckReport

    rep = CheckReport()
    rep.add("finto", OK, "va bene")
    assert rep.ok
    assert "Tutto pronto per il build" in rep.render()
