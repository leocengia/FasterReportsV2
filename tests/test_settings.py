"""Come si trovano i file delle sorgenti in `input/`.

Gli export reali portano la settimana nel nome (`AT DATASET W31.xlsx`) e
l'estensione dipende da chi li produce. I nomi si riconoscono per **pattern**:
pretendere `AT.csv` vorrebbe dire rinominare quattro file a mano ogni settimana.

Il rischio del pattern e' l'ambiguita' — due settimane nella stessa cartella — e
questi test fissano che in quel caso si **blocchi** dicendo quali file c'erano,
invece di scegliere per conto proprio.
"""

from __future__ import annotations

import pytest

from fasterreports.core.errors import ContractError, SourceError
from fasterreports.omni.orchestrate import _source_label
from fasterreports.omni.settings import Settings


def make_settings(tmp_path, input_files, *, subdir="input") -> Settings:
    d = tmp_path / subdir
    d.mkdir(exist_ok=True)
    return Settings(root=tmp_path, input_dir=d, input_files=input_files)


def test_pattern_con_un_solo_file(tmp_path):
    s = make_settings(tmp_path, {"AT_DATASET": "AT DATASET*"})
    atteso = s.input_dir / "AT DATASET W31.xlsx"
    atteso.touch()
    assert s.input_path("AT_DATASET") == atteso


def test_pattern_indifferente_all_estensione(tmp_path):
    """Lo stesso pattern deve prendere il CSV o l'XLSX: il formato lo decide
    l'estensione al momento della lettura, non la configurazione."""
    s = make_settings(tmp_path, {"SF_DATABASE": "SF DATABASE*"})
    (s.input_dir / "SF DATABASE W31.csv").touch()
    assert s.input_path("SF_DATABASE").suffix == ".csv"


def test_nessun_file_elenca_quelli_presenti(tmp_path):
    s = make_settings(tmp_path, {"Turni": "Turni*"})
    (s.input_dir / "AT DATASET W31.xlsx").touch()
    with pytest.raises(SourceError) as e:
        s.input_path("Turni")
    msg = str(e.value)
    assert "'Turni*'" in msg                    # cosa cercava
    assert "AT DATASET W31.xlsx" in msg         # cosa c'era davvero
    assert "settings.yml" in msg                # dove si cambia


def test_due_settimane_nella_cartella_blocca(tmp_path):
    """Il guasto vero: la settimana vecchia rimasta in cartella. Scegliere il
    file "piu' recente" produrrebbe un report giusto per caso e sbagliato in
    silenzio la volta che il timestamp inganna."""
    s = make_settings(tmp_path, {"AT_DATASET": "AT DATASET*"})
    (s.input_dir / "AT DATASET W30.xlsx").touch()
    (s.input_dir / "AT DATASET W31.xlsx").touch()
    with pytest.raises(SourceError) as e:
        s.input_path("AT_DATASET")
    msg = str(e.value)
    assert "W30" in msg and "W31" in msg
    assert "settimana vecchia" in msg


def test_ignora_i_file_di_lavoro_di_excel(tmp_path):
    """Un file aperto in Excel genera `~$nome.xlsx` accanto: non e' una sorgente,
    e senza escluderlo il pattern diventerebbe ambiguo proprio mentre si guarda
    il file."""
    s = make_settings(tmp_path, {"AT_DATASET": "AT DATASET*"})
    vero = s.input_dir / "AT DATASET W31.xlsx"
    vero.touch()
    (s.input_dir / "~$AT DATASET W31.xlsx").touch()
    assert s.input_path("AT_DATASET") == vero


def test_ignora_le_cartelle(tmp_path):
    s = make_settings(tmp_path, {"AT_DATASET": "AT*"})
    (s.input_dir / "AT vecchi").mkdir()
    vero = s.input_dir / "AT DATASET W31.xlsx"
    vero.touch()
    assert s.input_path("AT_DATASET") == vero


def test_nome_esatto_resta_nome_esatto(tmp_path):
    """Chi preferisce nomi fissi non deve essere costretto ai pattern."""
    s = make_settings(tmp_path, {"AT_DATASET": "AT.csv"})
    assert s.input_path("AT_DATASET") == s.input_dir / "AT.csv"


def test_cartella_input_inesistente(tmp_path):
    s = Settings(
        root=tmp_path,
        input_dir=tmp_path / "non_esiste",
        input_files={"AT_DATASET": "AT*"},
    )
    with pytest.raises(SourceError) as e:
        s.input_path("AT_DATASET")
    assert "non_esiste" in str(e.value)


def test_dataset_senza_pattern(tmp_path):
    s = make_settings(tmp_path, {"AT_DATASET": "AT*"})
    with pytest.raises(ContractError) as e:
        s.input_path("Turni")
    assert "AT_DATASET" in str(e.value)   # dice quali sono configurati


# ---------------------------------------------------------------------------
# Il rapporto deve sopravvivere al guasto che sta descrivendo
# ---------------------------------------------------------------------------


def test_source_label_non_solleva_se_il_file_manca(tmp_path):
    """`input_path` solleva quando la sorgente manca, ed e' giusto. Ma il ramo che
    *scrive* l'errore nel preflight la chiamava di nuovo: la prima fonte assente
    faceva morire tutto il rapporto invece di elencare le sei righe."""
    s = make_settings(tmp_path, {"Turni": "Turni*"})
    label = _source_label(s, "Turni")
    assert "Turni*" in label
    assert "nessun file" in label


def test_source_label_pattern_non_configurato(tmp_path):
    s = make_settings(tmp_path, {})
    assert "non configurato" in _source_label(s, "Turni")


def test_source_label_file_presente(tmp_path):
    s = make_settings(tmp_path, {"AT_DATASET": "AT DATASET*"})
    (s.input_dir / "AT DATASET W31.xlsx").touch()
    assert _source_label(s, "AT_DATASET").endswith("AT DATASET W31.xlsx")
