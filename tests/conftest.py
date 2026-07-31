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
    """Intestazioni reali estratte da 'Omni Report W30.xlsm'.

    Rigenerabile con:
      python tools/audit_workbook.py "samples/Omni Report W30.xlsm" \\
          --headers --usage --tables --json > tests/fixtures/workbook_W30.json
    """
    path = FIXTURES / "workbook_W30.json"
    if not path.is_file():
        pytest.skip(f"fixture assente: {path}")
    return json.loads(path.read_text(encoding="utf-8"))
