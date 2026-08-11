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
from fasterreports.omni.settings import Settings, load_settings


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


# ---------------------------------------------------------------------------
# load_settings: gli stessi errori parlanti anche sugli YAML malformati
#
# `sources.skills` e `sources.backoffice_sections` finivano dentro `tuple(...)`
# senza controllare che il valore fosse gia' una lista. Una stringa e'
# iterabile carattere per carattere: `skills: "HPO"` invece di
# `skills: ["HPO"]` produceva `('H', 'P', 'O')` — un filtro completamente
# sbagliato, senza che nessun errore lo segnalasse. Bug reale, non ipotetico.
# ---------------------------------------------------------------------------


def _scrivi_settings(tmp_path, corpo_sources: str) -> Path:
    p = tmp_path / "settings.yml"
    p.write_text(f"sources:\n{corpo_sources}\n", encoding="utf-8")
    return p


def test_skills_stringa_nuda_blocca_con_lerrore_giusto(tmp_path):
    p = _scrivi_settings(tmp_path, '  skills: "HPO"')
    with pytest.raises(ContractError) as e:
        load_settings(p, root=tmp_path)
    msg = str(e.value)
    assert "sources.skills" in msg
    assert '["HPO"]' in msg  # mostra la forma corretta, non solo l'errore


def test_skills_lista_funziona_come_sempre(tmp_path):
    p = _scrivi_settings(tmp_path, '  skills: ["HPO", "RETAIL"]')
    s = load_settings(p, root=tmp_path)
    assert s.sources.skills == ("HPO", "RETAIL")


def test_skills_assente_usa_il_default(tmp_path):
    p = tmp_path / "settings.yml"
    p.write_text("{}\n", encoding="utf-8")
    s = load_settings(p, root=tmp_path)
    assert s.sources.skills == ("HPO",)


def test_backoffice_sections_stringa_nuda_blocca(tmp_path):
    p = _scrivi_settings(tmp_path, '  backoffice_sections: "HPO"')
    with pytest.raises(ContractError) as e:
        load_settings(p, root=tmp_path)
    assert "sources.backoffice_sections" in str(e.value)


def test_backoffice_sections_lista_funziona_come_sempre(tmp_path):
    p = _scrivi_settings(tmp_path, '  backoffice_sections: ["HPO", "Part-Time 6h"]')
    s = load_settings(p, root=tmp_path)
    assert s.sources.backoffice_sections == ("HPO", "Part-Time 6h")


def test_monday_serial_non_numerico_da_contracterror_non_valueerror(tmp_path):
    p = _scrivi_settings(tmp_path, "  monday_serial: non-un-numero")
    with pytest.raises(ContractError) as e:
        load_settings(p, root=tmp_path)
    assert "sources.monday_serial" in str(e.value)


def test_monday_serial_numerico_funziona(tmp_path):
    p = _scrivi_settings(tmp_path, "  monday_serial: 46223")
    s = load_settings(p, root=tmp_path)
    assert s.sources.monday_serial == 46223


def test_monday_serial_assente_e_none(tmp_path):
    p = tmp_path / "settings.yml"
    p.write_text("{}\n", encoding="utf-8")
    s = load_settings(p, root=tmp_path)
    assert s.sources.monday_serial is None
