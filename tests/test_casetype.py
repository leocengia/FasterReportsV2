"""I case type nuovi finiscono in 'Helper CaseType' senza spostare niente.

Il difetto che questo pezzo elimina, misurato sull'export della W33: la lista
scritta a mano nel template copriva 31 case type, i dati ne avevano 32, e quattro
non erano in lista. Uno di essi ('Call Assignment', 30 casi) superava perfino la
soglia di volume del deepdive — un risultato vero che non compariva da nessuna
parte, senza nessun errore.

Il vincolo da rispettare mentre si rimedia: 'CaseType Deepdive' punta alle righe
di 'Helper CaseType' per POSIZIONE e porta formattazione fatta a mano. Muovere
una riga esistente metterebbe i case type sotto le etichette sbagliate.
"""

from __future__ import annotations

from fasterreports.core.casetype import coppie_dai_dati, da_appendere


def test_coppie_dai_dati_traduce_il_canale():
    """L'export dice Phone/Other, l'helper vuole Phone/Non-live."""
    assert coppie_dai_dati([("Phone", "EVC"), ("Other", "EVC")]) == [
        ("Non-live", "EVC"),
        ("Phone", "EVC"),
    ]


def test_coppie_dai_dati_deduplica():
    assert coppie_dai_dati([("Phone", "EVC")] * 50) == [("Phone", "EVC")]


def test_coppie_dai_dati_scarta_i_case_type_vuoti():
    assert coppie_dai_dati([("Phone", ""), ("Phone", None), ("Phone", " ")]) == []


def test_coppie_dai_dati_pulisce_gli_spazi():
    assert coppie_dai_dati([("Phone", "  EVC  ")]) == [("Phone", "EVC")]


# ---------------------------------------------------------------------------


def test_appende_solo_quello_che_manca():
    esistenti = [("Phone", "EVC"), ("Non-live", "EVC")]
    presenti = [("Phone", "EVC"), ("Phone", "Call Assignment"), ("Non-live", "EVC")]
    assert da_appendere(esistenti, presenti) == [("Phone", "Call Assignment")]


def test_niente_da_fare_se_la_lista_e_completa():
    coppie = [("Phone", "EVC"), ("Non-live", "Alfa")]
    assert da_appendere(coppie, coppie) == []


def test_non_propone_mai_di_rimuovere():
    """Un case type senza casi questa settimana resta in lista con volume zero:
    togliendolo si sposterebbero tutte le righe sotto di lui, che e' esattamente
    cio' che 'CaseType Deepdive' non puo' sopportare."""
    esistenti = [("Phone", "Contract Update"), ("Phone", "EVC")]
    assert da_appendere(esistenti, [("Phone", "EVC")]) == []


def test_lo_stesso_case_type_su_due_canali_sono_due_righe():
    """Nella W33 'Specialty Functions' mancava su entrambi i canali: sono due
    combinazioni distinte, non una."""
    nuovi = da_appendere([], [("Phone", "Specialty Functions"), ("Non-live", "Specialty Functions")])
    assert len(nuovi) == 2


def test_ordine_deterministico():
    """L'ordine di scoperta nell'export non deve entrare nel foglio: due giri
    sugli stessi dati devono appendere nello stesso ordine."""
    presenti = [("Phone", "Zeta"), ("Non-live", "Alfa"), ("Phone", "Alfa")]
    assert da_appendere([], presenti) == da_appendere([], list(reversed(presenti)))
    assert da_appendere([], presenti) == [
        ("Non-live", "Alfa"), ("Phone", "Alfa"), ("Phone", "Zeta")
    ]


def test_confronto_insensibile_agli_spazi_intorno():
    """Un ' EVC' nella lista scritta a mano non deve far riappendere 'EVC'."""
    assert da_appendere([(" Phone ", " EVC ")], [("Phone", "EVC")]) == []
