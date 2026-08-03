"""Generazione della patch VBA.

Il modulo non modifica il workbook: produce il testo del modulo patchato, che
l'utente incolla nell'editor. La ragione sta nel docstring dello strumento — il
p-code compilato accanto al sorgente rende la modifica binaria un modo di
fallire silenzioso.

Qui si verifica che le tre modifiche siano quelle giuste, che non tocchino altro,
e che rieseguirle non faccia danni.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

from make_vba_patch import patch, verify  # noqa: E402

SAMPLE = ROOT / "samples" / "omni-report" / "Omni Report W30.xlsm"

MODULO_MINIMO = """Attribute VB_Name = "CreaMalpractice"
Option Explicit
Private mAHT As String

Public Sub Refresh_Dettaglio_Malpractice()
    On Error GoTo CleanFail
    Dim outRows As Collection
    MsgBox "Dettaglio Malpractice rigenerato: " & outRows.Count & " righe.", vbInformation
    Exit Sub
CleanFail:
    Application.ScreenUpdating = True
    MsgBox "Errore durante la rigenerazione: " & Err.Description, vbExclamation
End Sub
"""


def test_le_tre_modifiche():
    out, done = patch(MODULO_MINIMO)
    assert "Public SilentMode As Boolean" in out
    assert "Public Sub SetSilentMode(ByVal value As Boolean)" in out
    assert out.count("If Not SilentMode Then MsgBox") == 2
    assert "If SilentMode Then Err.Raise Err.Number, , Err.Description" in out
    assert len([d for d in done if not d.startswith("ATTENZIONE")]) == 3


def test_il_risultato_passa_i_controlli():
    out, _ = patch(MODULO_MINIMO)
    assert verify(out) == []


def test_dichiarazione_dopo_option_explicit():
    """Prima di Option Explicit sarebbe un errore di compilazione."""
    out, _ = patch(MODULO_MINIMO)
    lines = out.splitlines()
    assert lines.index("Option Explicit") < lines.index("Public SilentMode As Boolean")


def test_err_raise_dentro_cleanfail_prima_di_end_sub():
    out, _ = patch(MODULO_MINIMO)
    lines = [l.strip() for l in out.splitlines()]
    i_label = lines.index("CleanFail:")
    i_raise = next(i for i, l in enumerate(lines) if l.startswith("If SilentMode Then Err.Raise"))
    i_end = next(i for i in range(i_label, len(lines)) if lines[i] == "End Sub")
    assert i_label < i_raise < i_end


def test_indentazione_preservata():
    out, _ = patch(MODULO_MINIMO)
    for line in out.splitlines():
        if "If Not SilentMode Then MsgBox" in line:
            assert line.startswith("    "), f"indentazione persa: {line!r}"


def test_idempotente():
    p1, _ = patch(MODULO_MINIMO)
    p2, done = patch(p1)
    assert p1 == p2
    assert any("gia' presente" in d for d in done)


def test_non_tocca_le_altre_righe():
    out, _ = patch(MODULO_MINIMO)
    for line in ("Private mAHT As String", "Dim outRows As Collection", "Exit Sub"):
        assert line in out


def test_msgbox_commentato_ignorato():
    code = MODULO_MINIMO.replace(
        '    MsgBox "Errore', "    ' MsgBox \"Errore"
    )
    out, _ = patch(code)
    assert out.count("If Not SilentMode Then MsgBox") == 1
    assert "' MsgBox \"Errore" in out


def test_msgbox_come_funzione_segnalato_non_modificato():
    """`x = MsgBox(...)` e' una domanda, non un avviso: si segnala e non si tocca."""
    code = MODULO_MINIMO.replace(
        '    MsgBox "Errore durante la rigenerazione: " & Err.Description, vbExclamation',
        '    risposta = MsgBox("Continuare?", vbYesNo)',
    )
    out, done = patch(code)
    assert "risposta = MsgBox(" in out
    assert "If Not SilentMode Then risposta" not in out
    assert any(d.startswith("ATTENZIONE") and "non in prima posizione" in d for d in done)


def test_senza_cleanfail_segnala():
    code = MODULO_MINIMO.replace("CleanFail:\n", "").replace(
        "    On Error GoTo CleanFail\n", ""
    )
    out, done = patch(code)
    assert any("nessuna etichetta CleanFail" in d for d in done)
    assert "Err.Raise assente" in " ".join(verify(out))


def test_senza_option_explicit():
    code = MODULO_MINIMO.replace("Option Explicit\n", "")
    out, _ = patch(code)
    lines = out.splitlines()
    # La dichiarazione va dopo l'ultimo Attribute, non in cima al file.
    assert lines[0].startswith("Attribute VB_Name")
    assert verify(out) == []


def test_flag_personalizzato():
    out, _ = patch(MODULO_MINIMO, flag="ModoZitto")
    assert "Public ModoZitto As Boolean" in out
    assert "Public Sub SetModoZitto" in out
    assert verify(out, flag="ModoZitto") == []


# --- sul modulo vero -------------------------------------------------------

@pytest.mark.skipif(not SAMPLE.is_file(), reason="campione W30 assente")
def test_sul_modulo_reale():
    pytest.importorskip("oletools", reason="serve oletools per leggere il VBA")
    from make_vba_patch import read_module

    original = read_module(SAMPLE, "CreaMalpractice")
    assert "Refresh_Dettaglio_Malpractice" in original
    assert original.count("MsgBox") == 2  # uno di successo, uno d'errore

    out, done = patch(original)
    assert verify(out) == []
    assert not [d for d in done if d.startswith("ATTENZIONE")]

    # Il resto del modulo non e' stato toccato: solo 3 punti cambiano.
    import difflib

    diff = [
        l for l in difflib.unified_diff(
            original.replace("\r\n", "\n").splitlines(), out.splitlines(), n=0
        )
        if l.startswith("-") and not l.startswith("---")
    ]
    assert len(diff) == 2  # le due righe MsgBox riscritte, nient'altro rimosso


@pytest.mark.skipif(not SAMPLE.is_file(), reason="campione W30 assente")
def test_il_patchato_soddisfa_il_check_del_doctor():
    """Coerenza fra chi genera la patch e chi la verifica."""
    pytest.importorskip("oletools")
    from make_vba_patch import read_module

    out, _ = patch(read_module(SAMPLE, "CreaMalpractice"))
    # Gli stessi criteri di omni/doctor.py::_check_vba
    assert "SetSilentMode" in out
    nudi = [
        l.strip() for l in out.splitlines()
        if "MsgBox" in l and "SilentMode" not in l and not l.strip().startswith("'")
    ]
    assert not nudi
    assert "Err.Raise" in out
