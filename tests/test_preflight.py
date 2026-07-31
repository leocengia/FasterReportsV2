from __future__ import annotations

from fasterreports.core.csvsource import read_csv_text
from fasterreports.core.preflight import DatasetReport, PreflightReport, check_dataset

HEADERS = "Agent Email,Talk Time in seconds,Wrap-up Time in seconds,Handle Time in seconds,Initiation Method"


def _source(text: str):
    return read_csv_text(text, name="ATwi.csv")


def test_report_ok_mostra_la_mappatura(contract):
    ds = contract.dataset("ATwi_DATASET")
    src = _source(HEADERS + "\na@x.it,1,2,3,OUTBOUND\n")
    report = PreflightReport(generated_at="2026-07-31 10:00:00")
    report.datasets.append(check_dataset(contract, ds, src))

    assert report.ok
    text = report.render()
    assert "Esito complessivo: OK" in text
    assert "Talk Time in seconds" in text
    assert "esatto" in text
    # La lettera target va nel report: e' cio' che si va a controllare fra sei mesi.
    assert " H " in text or "| H " in text


def test_report_bloccato_spiega_e_non_solleva(contract):
    ds = contract.dataset("ATwi_DATASET")
    src = _source("Agent Email,Initiation Method\na@x.it,OUTBOUND\n")
    dr = check_dataset(contract, ds, src)

    # check_dataset cattura: si vuole il quadro di tutti i dataset in un run.
    assert not dr.ok
    report = PreflightReport()
    report.datasets.append(dr)
    text = report.render()
    assert "BLOCCATO" in text
    assert "Talk Time in seconds" in text
    assert "nessun workbook e' stato prodotto" in text


def test_report_riporta_encoding_e_separatore(contract):
    ds = contract.dataset("ATwi_DATASET")
    src = _source(HEADERS.replace(",", ";") + "\na@x.it;1;2;3;OUTBOUND\n")
    report = PreflightReport()
    report.datasets.append(check_dataset(contract, ds, src))
    text = report.render()
    assert "punto e virgola" in text


def test_report_segnala_le_colonne_scartate(contract):
    ds = contract.dataset("PSAT_DATASET")
    # 'agent_name' oltre ad 'Agent Name': la disambiguazione va tracciata.
    src = read_csv_text(
        "Agent Name,agent_name,Survey Date (Exact),Case Number,psat_score,response_text\n"
        "Mario,mario,2026-07-28,C1,1,ok\n",
        name="PSAT.csv",
    )
    report = PreflightReport()
    report.datasets.append(check_dataset(contract, ds, src))
    text = report.render()
    assert "disambiguata da match esatto" in text
    assert "'agent_name'" in text


def test_report_scritto_su_file(contract, tmp_path):
    ds = contract.dataset("ATwi_DATASET")
    src = _source(HEADERS + "\na@x.it,1,2,3,OUTBOUND\n")
    report = PreflightReport()
    report.datasets.append(check_dataset(contract, ds, src))
    path = report.write(tmp_path / "sub" / "preflight_W31.txt")
    assert path.is_file()
    assert "PREFLIGHT" in path.read_text(encoding="utf-8")


def test_report_conta_le_colonne_non_usate(contract):
    ds = contract.dataset("ATwi_DATASET")
    src = _source(HEADERS + ",Extra1,Extra2\na@x.it,1,2,3,OUTBOUND,x,y\n")
    report = PreflightReport()
    report.datasets.append(check_dataset(contract, ds, src))
    assert "non usate dal motore: 2" in report.render()


def test_errore_di_lettura_finisce_nel_report():
    report = PreflightReport()
    report.datasets.append(
        DatasetReport(name="AT_DATASET", source="input/AT.csv", error="File sorgente non trovato")
    )
    assert not report.ok
    assert "File sorgente non trovato" in report.render()
