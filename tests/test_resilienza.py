"""Resilienza ai cambi di export — piano §11.

Tre scenari, uno per esito:
  (a) colonne mescolate   -> output identico
  (b) colonne rinominate  -> output identico, se il nome nuovo e' un alias
  (c) colonna rimossa     -> preflight fallito con messaggio chiaro

E' il motivo per cui esiste tutto il layer di ingestione: se questi tre test
passano, un export che cambia ordine non produce piu' numeri falsi.
"""

from __future__ import annotations

import pytest

from fasterreports.core.csvsource import read_csv_text
from fasterreports.core.errors import MissingColumnError
from fasterreports.core.matcher import resolve_dataset
from fasterreports.core.transform import build_block

# Un CSV ATwi minimo ma realistico: 5 campi del contratto piu' colonne di rumore.
BASE_HEADERS = [
    "Workitem ID",
    "Agent Email",
    "Agent Business Location",
    "Talk Time in seconds",
    "Wrap-up Time in seconds",
    "Handle Time in seconds",
    "Initiation Method",
]
BASE_ROWS = [
    ["w1", "mario@x.it", "Milano", "120", "45", "165", "OUTBOUND"],
    ["w2", "lucia@x.it", "Roma", "300", "60", "360", "INBOUND"],
]


def _csv(headers: list[str], rows: list[list[str]]) -> str:
    lines = [",".join(headers)]
    lines += [",".join(r) for r in rows]
    return "\n".join(lines) + "\n"


def _build(contract, headers, rows):
    ds = contract.dataset("ATwi_DATASET")
    src = read_csv_text(_csv(headers, rows))
    mapping = resolve_dataset(ds, list(src.headers), contract.normalize)
    return build_block(ds, mapping, src.rows())


@pytest.fixture
def baseline(contract):
    return _build(contract, BASE_HEADERS, BASE_ROWS)


def test_baseline_scrive_nelle_lettere_canoniche(baseline):
    # Blocco da B a M (data_start_col=B, ultimo input=M).
    assert (baseline.start_col, baseline.end_col) == ("B", "M")
    riga = baseline.rows[0]
    # offset 0 = colonna B, quindi C=1, H=6, I=7, J=8, M=11
    assert riga[1] == "mario@x.it"   # C
    assert riga[6] == 120.0          # H
    assert riga[7] == 45.0           # I
    assert riga[8] == 165.0          # J
    assert riga[11] == "OUTBOUND"    # M
    # Le colonne non mappate restano vuote: il motore non le legge.
    assert riga[0] is None and riga[2] is None


def test_a_colonne_mescolate_output_identico(contract, baseline):
    order = [4, 0, 6, 2, 5, 1, 3]  # permutazione arbitraria
    headers = [BASE_HEADERS[i] for i in order]
    rows = [[r[i] for i in order] for r in BASE_ROWS]
    assert _build(contract, headers, rows).rows == baseline.rows


def test_a_colonne_mescolate_con_colonne_nuove_in_mezzo(contract, baseline):
    # Caso realistico: l'export aggiunge campi e sposta i vecchi.
    headers = ["Nuova A"] + BASE_HEADERS[3:] + ["Nuova B"] + BASE_HEADERS[:3]
    rows = [
        ["?"] + r[3:] + ["?"] + r[:3]
        for r in BASE_ROWS
    ]
    assert _build(contract, headers, rows).rows == baseline.rows


def test_b_colonne_rinominate_negli_alias_output_identico(contract, baseline):
    headers = list(BASE_HEADERS)
    headers[1] = "agent_email"              # alias di Agent Email
    headers[3] = "talk_time_seconds"        # alias
    headers[4] = "wrap_up_time_seconds"     # alias
    headers[5] = "handle_time_in_seconds"   # alias
    headers[6] = "initiation_method"        # alias
    assert _build(contract, headers, BASE_ROWS).rows == baseline.rows


def test_b_rinomina_solo_di_stile_output_identico(contract, baseline):
    # Maiuscole, underscore e trattini diversi: assorbiti dalla normalizzazione
    # anche senza un alias dedicato.
    headers = list(BASE_HEADERS)
    headers[4] = "WRAP UP TIME IN SECONDS"
    headers[5] = "Handle_Time_In_Seconds"
    assert _build(contract, headers, BASE_ROWS).rows == baseline.rows


def test_c_colonna_rimossa_fallisce_con_messaggio_chiaro(contract):
    headers = [h for h in BASE_HEADERS if h != "Talk Time in seconds"]
    rows = [[v for i, v in enumerate(r) if i != 3] for r in BASE_ROWS]
    with pytest.raises(MissingColumnError) as e:
        _build(contract, headers, rows)
    msg = str(e.value)
    assert "Talk Time in seconds" in msg
    assert "ATwi_DATASET" in msg
    assert "talk_time_seconds" in msg          # alias provati
    assert "Initiation Method" in msg          # header effettivamente presenti


def test_c_rinomina_non_prevista_fallisce_invece_di_indovinare(contract):
    headers = list(BASE_HEADERS)
    headers[3] = "TalkSecs"  # nome nuovo, non e' un alias
    with pytest.raises(MissingColumnError):
        _build(contract, headers, BASE_ROWS)


def test_separatore_punto_e_virgola(contract, baseline):
    text = ";".join(BASE_HEADERS) + "\n" + "\n".join(";".join(r) for r in BASE_ROWS) + "\n"
    ds = contract.dataset("ATwi_DATASET")
    src = read_csv_text(text, name="semi.csv")
    assert src.delimiter == ";"
    mapping = resolve_dataset(ds, list(src.headers), contract.normalize)
    assert build_block(ds, mapping, src.rows()).rows == baseline.rows


def test_virgole_nei_verbatim_non_confondono_il_separatore(contract):
    # I commenti PSAT contengono virgole e punti e virgola: il conteggio del
    # separatore deve ignorare cio' che sta fra virgolette.
    ds = contract.dataset("PSAT_DATASET")
    headers = ["Agent Name", "Survey Date (Exact)", "Case Number", "psat_score", "response_text"]
    text = (
        ";".join(headers)
        + "\n"
        + 'Mario Rossi;2026-07-28;C1;1;"Ottimo, davvero; molto disponibile"\n'
    )
    src = read_csv_text(text, name="psat.csv")
    assert src.delimiter == ";"
    mapping = resolve_dataset(ds, list(src.headers), contract.normalize)
    block = build_block(ds, mapping, src.rows())
    # DQ = response_text, offset DQ-A = 120
    assert block.rows[0][120] == "Ottimo, davvero; molto disponibile"
