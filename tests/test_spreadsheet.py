from pathlib import Path

import pytest
from openpyxl import load_workbook

from conference_agenda.spreadsheet import WorkbookValidationError, parse_workbook


ROOT = Path(__file__).resolve().parents[1]
SEED_WORKBOOK = ROOT / "seed_data" / "ceha-aes-2026.xlsx"


def test_seed_workbook_parses_with_expected_content():
    config = parse_workbook(SEED_WORKBOOK.read_bytes())

    assert config["conference"]["organization_slug"] == "ceha"
    assert config["conference"]["event_slug"] == "aes-2026"
    assert config["conference"]["start_date"] == "2026-04-20"
    assert config["conference"]["end_date"] == "2026-04-24"
    assert len(config["sessions"]) == 63
    assert any(
        session["session_name"] == "Practical AI Recipes for EH Professionals"
        for session in config["sessions"]
    )


def test_workbook_rejects_missing_required_session_column(tmp_path):
    path = tmp_path / "invalid.xlsx"
    workbook = load_workbook(SEED_WORKBOOK)
    sheet = workbook["Sessions"]
    for cell in sheet[1]:
        if cell.value == "end_time":
            cell.value = "removed_end_time"
    workbook.save(path)

    with pytest.raises(WorkbookValidationError) as error:
        parse_workbook(path.read_bytes())

    assert "missing required column end_time" in str(error.value)
