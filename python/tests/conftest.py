import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

DATA_FILE = ROOT.parent / "data" / "Treasury_data_090426.xlsx"


@pytest.fixture(scope="session")
def data_path() -> Path:
    if not DATA_FILE.exists():
        pytest.skip("Treasury data workbook not available")
    return DATA_FILE
