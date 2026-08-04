from __future__ import annotations

import pytest

from fasterreports.core.contract import Field, parse_contract
from fasterreports.core.errors import AmbiguousColumnError, MissingColumnError
from fasterreports.core.matcher import (
    VIA_ALIAS_EXACT,
    VIA_ALIAS_NORMALIZED,
    VIA_EXACT,
    VIA_NORMALIZED,
    resolve_dataset,
    resolve_field,
)
from fasterreports.core.normalize import NormalizeRules

RULES = NormalizeRules()


def fld(canonical: str, *, aliases=(), match="exact_first", target="B") -> Field:
    return Field(
        canonical=canonical,
        target_col=target,
        role="input",
        dtype="str",
        match=match,
        aliases=tuple(aliases),
    )


# --- priorita' dei livelli di match -----------------------------------------

def test_match_esatto():
    r = resolve_field("D", fld("Agent Email"), ["x", "Agent Email", "y"], RULES)
    assert (r.source_index, r.via) == (1, VIA_EXACT)


def test_match_normalizzato_quando_lesatto_manca():
    r = resolve_field("D", fld("Agent Email"), ["agent_email"], RULES)
    assert (r.source_index, r.via) == (0, VIA_NORMALIZED)


def test_match_alias_esatto():
    r = resolve_field("D", fld("Agent Email", aliases=["Agent E-mail"]), ["Agent E-mail"], RULES)
    assert (r.source_index, r.via, r.matched_on) == (0, VIA_ALIAS_EXACT, "Agent E-mail")


def test_match_alias_normalizzato():
    f = fld("Wrap-up Time in seconds", aliases=["wrap_up_time_seconds"])
    r = resolve_field("D", f, ["Wrap Up Time Seconds"], RULES)
    assert r.via == VIA_ALIAS_NORMALIZED


def test_lesatto_vince_sul_normalizzato():
    # Il caso reale PSAT: 'Agent Name' (I) e 'agent_name' (BO) coesistono.
    # Senza priorita' all'esatto questo file sarebbe ambiguo, pur essendo valido.
    headers = ["agent_name", "Agent Name"]
    r = resolve_field("PSAT_DATASET", fld("Agent Name"), headers, RULES)
    assert (r.source_index, r.via) == (1, VIA_EXACT)


def test_il_canonico_vince_sullalias():
    # Un alias non deve mai scavalcare il nome canonico presente nel file.
    headers = ["case_no", "case_number"]
    r = resolve_field("SF", fld("case_number", aliases=["case_no"]), headers, RULES)
    assert (r.source_index, r.via) == (1, VIA_EXACT)


def test_spazio_finale_non_rompe_lesatto():
    r = resolve_field("D", fld("Case Closure SLA"), ["Case Closure SLA "], RULES)
    assert r.via == VIA_EXACT


# --- errori -----------------------------------------------------------------

def test_colonna_mancante():
    with pytest.raises(MissingColumnError) as e:
        resolve_field("SF", fld("Case AHT (mins)", aliases=["case_aht_mins"]), ["Foo", "Bar"], RULES)
    msg = str(e.value)
    assert "Case AHT (mins)" in msg
    assert "case_aht_mins" in msg  # gli alias provati sono nel messaggio
    assert "Foo" in msg  # e anche gli header trovati


def test_colonna_mancante_suggerisce_il_nome_simile():
    with pytest.raises(MissingColumnError) as e:
        resolve_field("SF", fld("Employee Name"), ["Employe Name", "Zzz"], RULES)
    assert "Forse intendevi" in str(e.value)


def test_ambiguita_su_normalizzato():
    # Nessun match esatto: 'psat score' e 'psat_score' sono indistinguibili.
    with pytest.raises(AmbiguousColumnError) as e:
        resolve_field("PSAT", fld("psat score"), ["psat_score", "PSAT Score"], RULES)
    assert "ambigua" in str(e.value)
    assert "columns.yml" in str(e.value)  # l'errore dice come si rimedia


def test_ambiguita_su_header_duplicati():
    with pytest.raises(AmbiguousColumnError):
        resolve_field("D", fld("Agent Email"), ["Agent Email", "Agent Email"], RULES)


def test_match_exact_non_ripiega():
    # match=exact: se il nome esatto non c'e', si deve fermare, non agganciare
    # la colonna 'quasi giusta'.
    f = fld("Agent Name", match="exact")
    with pytest.raises(MissingColumnError) as e:
        resolve_field("PSAT", f, ["agent_name"], RULES)
    assert "solo nome esatto" in str(e.value)


def test_match_exact_ignora_gli_alias():
    f = fld("psat_score", match="exact", aliases=["PSAT Score"])
    with pytest.raises(MissingColumnError):
        resolve_field("PSAT", f, ["PSAT Score"], RULES)


# --- risoluzione di un dataset intero ---------------------------------------

MINI = {
    "normalize": {},
    "datasets": {
        "D": {
            "sheet": "D",
            "header_row": 1,
            "data_start_col": "A",
            "fields": [
                {"canonical": "Alpha", "target_col": "A", "role": "input", "dtype": "str"},
                {"canonical": "Beta", "target_col": "B", "role": "input", "dtype": "float"},
                {"canonical": "Gamma", "target_col": "C", "role": "input", "dtype": "str",
                 "aliases": ["gamma_alias"]},
            ],
        }
    },
}


@pytest.fixture
def mini():
    return parse_contract(MINI)


def test_dataset_ordine_colonne_irrilevante(mini):
    ds = mini.dataset("D")
    m1 = resolve_dataset(ds, ["Alpha", "Beta", "Gamma"], mini.normalize)
    m2 = resolve_dataset(ds, ["Gamma", "Alpha", "Beta"], mini.normalize)
    # Le posizioni sorgente cambiano, la mappa canonico->target no.
    assert {r.canonical: r.target_col for r in m1.resolutions} == {
        r.canonical: r.target_col for r in m2.resolutions
    }
    assert m2.by_canonical()["Gamma"].source_index == 0


def test_dataset_raccoglie_tutti_gli_errori(mini):
    ds = mini.dataset("D")
    with pytest.raises(MissingColumnError) as e:
        resolve_dataset(ds, ["Alpha"], mini.normalize)
    msg = str(e.value)
    # Due colonne mancanti, un solo run: si sistemano entrambe insieme.
    assert "2 colonne non risolte" in msg
    assert "Beta" in msg and "Gamma" in msg


def test_dataset_segnala_header_inutilizzati(mini):
    ds = mini.dataset("D")
    m = resolve_dataset(ds, ["Alpha", "Beta", "Gamma", "Extra1", "Extra2"], mini.normalize)
    assert m.unused_headers == ("Extra1", "Extra2")


def test_dataset_rifiuta_due_campi_sulla_stessa_colonna():
    raw = {
        "normalize": {},
        "datasets": {
            "D": {
                "sheet": "D", "header_row": 1, "data_start_col": "A",
                "fields": [
                    {"canonical": "Alpha", "target_col": "A", "role": "input"},
                    # 'alpha' normalizza come 'Alpha': entrambi aggancerebbero
                    # la stessa colonna sorgente.
                    {"canonical": "alpha", "target_col": "B", "role": "input"},
                ],
            }
        },
    }
    c = parse_contract(raw)
    with pytest.raises(AmbiguousColumnError) as e:
        resolve_dataset(c.dataset("D"), ["Alpha"], c.normalize)
    assert "stessa colonna sorgente" in str(e.value)


def test_shadowed_registra_la_disambiguazione(mini):
    ds = mini.dataset("D")
    m = resolve_dataset(ds, ["Alpha", "Beta", "Gamma", "gamma"], mini.normalize)
    gamma = m.by_canonical()["Gamma"]
    # 'gamma' e' stato scartato: va scritto nel preflight, non nascosto.
    assert gamma.shadowed == ("gamma",)
    assert gamma.was_disambiguated
