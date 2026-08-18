from __future__ import annotations

import json
from pathlib import Path

import pytest

from fasterreports.core.contract import load_contract

ROOT = Path(__file__).resolve().parents[1]
FIXTURES = Path(__file__).parent / "fixtures"


@pytest.fixture(scope="session")
def contract():
    """Il contratto vero del repo: i test lo esercitano, non una copia."""
    return load_contract(ROOT / "config" / "columns.yml")


@pytest.fixture(scope="session")
def real_headers() -> dict:
    """Intestazioni e consumi estratti dal TEMPLATE tracciato nel repo.

    Prima veniva da 'samples/Omni Report W30.xlsm', un workbook consegnato. E'
    stata spostata sul template il 2026-08-18, per un motivo concreto: il
    template aveva guadagnato 'Helper CaseType', che legge SF_DATABASE!EJ, e
    quella colonna mancava dal contratto — ma la fixture, ferma al W30, non
    poteva accorgersene. Il template e' anche versionato qui, quindi la fixture
    e' rigenerabile da chiunque abbia il repo, senza bisogno di un file di
    esempio esterno.

    Rigenerabile con:
      python tools/audit_workbook.py template/Omni_Report_TEMPLATE.xlsm \\
          --headers --usage --tables --json > tests/fixtures/workbook_template.json
    """
    path = FIXTURES / "workbook_template.json"
    if not path.is_file():
        pytest.skip(f"fixture assente: {path}")
    return json.loads(path.read_text(encoding="utf-8"))
