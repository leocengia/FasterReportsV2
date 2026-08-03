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


# --- codifiche: il pezzo che puo' corrompere NormKey in silenzio ----------

ACCENTATO = 'Attribute VB_Name = "M"\nOption Explicit\nSub S()\n' \
            '    s = Replace(s, "à", "a")   \' 100° caso\n' \
            '    MsgBox "fine"\nEnd Sub\n'


def test_bas_per_import_in_ansi_senza_bom(tmp_path):
    """L'editor VBA esporta e importa in ANSI: UTF-8 gli corromperebbe gli accenti."""
    from make_vba_patch import _write

    p = tmp_path / "m.bas"
    _write(p, ACCENTATO, for_import_route=True)
    raw = p.read_bytes()
    assert not raw.startswith(b"\xef\xbb\xbf")
    assert '"à"' in raw.decode("cp1252")


def test_vb_per_incollare_in_utf8_con_bom(tmp_path):
    """Con il BOM gli editor di Windows non tirano a indovinare ANSI."""
    from make_vba_patch import _write

    p = tmp_path / "m.vb"
    _write(p, ACCENTATO, for_import_route=False)
    raw = p.read_bytes()
    assert raw.startswith(b"\xef\xbb\xbf")
    assert '"à"' in raw.decode("utf-8-sig")


def test_senza_bom_il_vb_letto_come_ansi_corrompe_normkey(tmp_path):
    """La ragione per cui il BOM c'e': senza, l'accento si rompe e nessuno lo vede."""
    from make_vba_patch import _write

    p = tmp_path / "m.vb"
    _write(p, ACCENTATO, for_import_route=False)
    corrotto = p.read_bytes().decode("cp1252", "replace")
    assert '"à"' not in corrotto  # e' proprio questo che il BOM previene


def test_crlf_in_entrambi(tmp_path):
    from make_vba_patch import _write

    for suffix, route in ((".bas", True), (".vb", False)):
        p = tmp_path / f"m{suffix}"
        _write(p, ACCENTATO, for_import_route=route)
        assert b"\r\n" in p.read_bytes()
        assert b"\n\n" not in p.read_bytes().replace(b"\r\n", b"\n").replace(b"\n\n", b"")


def test_carattere_non_rappresentabile_in_ansi_blocca(tmp_path):
    """Meglio fermarsi che scrivere un .bas che l'import corromperebbe."""
    from make_vba_patch import PatchError, _write

    with pytest.raises(PatchError) as e:
        _write(tmp_path / "m.bas", 'Sub S()\n  x = "中文"\nEnd Sub\n',
               for_import_route=True)
    assert "ANSI" in str(e.value)


@pytest.mark.skipif(not SAMPLE.is_file(), reason="campione W30 assente")
def test_accenti_di_normkey_sopravvivono_sul_modulo_reale(tmp_path):
    """NormKey del W30 mappa a-grave/e-grave/...: devono restare leggibili."""
    pytest.importorskip("oletools")
    from make_vba_patch import _write, read_module

    out, _ = patch(read_module(SAMPLE, "CreaMalpractice"))
    bas = tmp_path / "m.bas"
    vb = tmp_path / "m.vb"
    _write(bas, out, for_import_route=True)
    _write(vb, out, for_import_route=False)

    for testo in (bas.read_bytes().decode("cp1252"),
                  vb.read_bytes().decode("utf-8-sig")):
        for accento in ("à", "è", "é", "ì", "ò", "ù"):
            assert f'Replace(s, "{accento}"' in testo, f"{accento} corrotto"
        assert "° pct" in testo


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


# --- applicazione via xlwings (percorsi d'errore, senza Excel) --------------

def test_apply_via_xlwings_template_assente(tmp_path):
    from fasterreports.omni.vbapatch import PatchError, apply_via_xlwings

    with pytest.raises(PatchError) as e:
        apply_via_xlwings(tmp_path / "non_esiste.xlsm")
    assert "non trovato" in str(e.value)


def test_il_tool_riusa_la_logica_del_package():
    """Una regola sola: il tool non duplica la patch, la importa."""
    from fasterreports.omni import vbapatch
    import make_vba_patch

    assert make_vba_patch.patch is vbapatch.patch
    assert make_vba_patch.verify is vbapatch.verify
    assert make_vba_patch._write is vbapatch.write_encoded


def test_suggerimento_del_doctor_cita_patch_template(contract):
    """Chi legge l'errore deve trovare la via piu' rapida per primo."""
    from dataclasses import replace

    from fasterreports.omni.doctor import run_checks
    from fasterreports.omni.settings import load_settings

    if not SAMPLE.is_file():
        pytest.skip("campione W30 assente")
    settings = load_settings(ROOT / "config" / "settings.yml", root=ROOT)
    rep = run_checks(contract, replace(settings, template=SAMPLE), try_excel=False)
    hint = next(c for c in rep.checks if "SetSilentMode" in c.name).hint
    assert "omni-report patch-template" in hint
    assert "make_vba_patch" in hint


# ---------------------------------------------------------------------------
# L'ordine delle dichiarazioni: il modulo deve COMPILARE
# ---------------------------------------------------------------------------

MODULO_REALE = '''Attribute VB_Name = "CreaMalpractice"
Option Explicit
' ==== Etichette categoria: unica fonte di verita' ====
Private mAHT As String, mFast As String, mBreakDay As String
Private mBreakSim As String, mLogin As String, mAC As String

Public Sub Refresh_Dettaglio_Malpractice()
    On Error GoTo CleanFail
    MsgBox "fatto", vbInformation
    Exit Sub
CleanFail:
    MsgBox "errore: " & Err.Description, vbExclamation
End Sub

Private Function NormKey(v As Variant) As String
    NormKey = LCase(Trim(CStr(v)))
End Function
'''


def test_il_blocco_va_sotto_le_dichiarazioni_di_modulo():
    """Il guasto costato un run: il blocco veniva infilato dopo Option Explicit,
    quindi le `Private` del modulo finivano DOPO un End Sub. VBA rifiuta:

        Errore di compilazione: dopo End Sub, End Function o End Property
        sono ammessi solo commenti

    E non si vedeva prima, perche' Excel compila solo quando serve: il file si
    salvava senza un lamento e l'errore usciva quando la pipeline lanciava la
    macro, come dialogo modale che in automazione nessuno chiude.
    """
    from fasterreports.omni.vbapatch import check_declaration_order, patch

    nuovo, _ = patch(MODULO_REALE)
    assert check_declaration_order(nuovo) == []

    righe = nuovo.split("\n")
    decl = righe.index("Public SilentMode As Boolean")
    private = next(i for i, l in enumerate(righe) if l.startswith("Private mAHT"))
    prima_proc = next(i for i, l in enumerate(righe) if "Sub SetSilentMode" in l)
    assert private < decl < prima_proc


def test_check_declaration_order_trova_il_modulo_rotto():
    from fasterreports.omni.vbapatch import check_declaration_order

    rotto = (
        "Option Explicit\n"
        "Public SilentMode As Boolean\n"
        "\n"
        "Public Sub SetSilentMode(ByVal value As Boolean)\n"
        "    SilentMode = value\n"
        "End Sub\n"
        "' ==== Etichette ====\n"
        "Private mAHT As String, mFast As String\n"
        "\n"
        "Public Sub Refresh()\n"
        "    Dim wb As Workbook\n"
        "End Sub\n"
    )
    problemi = check_declaration_order(rotto)
    assert len(problemi) == 1
    assert "riga 8" in problemi[0]
    assert "mAHT" in problemi[0]


def test_dim_dentro_una_procedura_non_e_un_problema():
    """`Dim` locale e' la cosa piu' normale del mondo: un falso positivo qui
    renderebbe il controllo inutilizzabile."""
    from fasterreports.omni.vbapatch import check_declaration_order

    ok = (
        "Option Explicit\n"
        "Private m As String\n"
        "\n"
        "Public Sub Uno()\n"
        "    Dim wb As Workbook\n"
        "    Dim r As Long, c As Long\n"
        "    Set wb = ThisWorkbook\n"
        "End Sub\n"
        "\n"
        "Private Function Due(ByVal x As Long) As String\n"
        "    Dim s As String\n"
        "    Due = s\n"
        "End Function\n"
    )
    assert check_declaration_order(ok) == []


def test_verify_bocciava_un_modulo_che_non_compila():
    """`verify` guardava solo la presenza delle cose, non l'ordine: diceva
    'nessun problema' su un modulo che VBA rifiuta."""
    from fasterreports.omni.vbapatch import verify

    rotto = (
        "Option Explicit\n"
        "Public SilentMode As Boolean\n"
        "Public Sub SetSilentMode(ByVal value As Boolean)\n"
        "    SilentMode = value\n"
        "End Sub\n"
        "Private mAHT As String\n"
        "Public Sub Refresh()\n"
        "    If Not SilentMode Then MsgBox \"x\"\n"
        "    If SilentMode Then Err.Raise Err.Number\n"
        "End Sub\n"
    )
    problemi = verify(rotto)
    assert any("dopo una procedura" in p for p in problemi)


def test_il_commento_resta_attaccato_alla_sua_procedura():
    """Infilarsi fra il commento e la Sub che descrive li separerebbe."""
    from fasterreports.omni.vbapatch import patch

    code = (
        "Option Explicit\n"
        "Private m As String\n"
        "\n"
        "' Questo commento descrive Refresh, non il blocco nuovo\n"
        "Public Sub Refresh()\n"
        "    MsgBox \"x\"\n"
        "End Sub\n"
    )
    nuovo, _ = patch(code)
    righe = nuovo.split("\n")
    commento = next(i for i, l in enumerate(righe) if "descrive Refresh" in l)
    refresh = next(i for i, l in enumerate(righe) if "Sub Refresh()" in l)
    assert refresh == commento + 1


def test_modulo_senza_procedure_ripiega_su_option_explicit():
    from fasterreports.omni.vbapatch import check_declaration_order, patch

    code = 'Attribute VB_Name = "X"\nOption Explicit\nPrivate m As String\n'
    nuovo, _ = patch(code)
    assert "Public SilentMode As Boolean" in nuovo
    assert check_declaration_order(nuovo) == []
