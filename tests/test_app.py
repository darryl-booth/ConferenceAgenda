from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import pytest

from conference_agenda import create_app
import conference_agenda.routes as routes
from conference_agenda.routes import (
    _create_or_update_event,
    _create_revision,
    _selected_event_day,
)
from conference_agenda.spreadsheet import parse_workbook


ROOT = Path(__file__).resolve().parents[1]
SEED_WORKBOOK = ROOT / "seed_data" / "ceha-aes-2026.xlsx"


@pytest.fixture()
def app(tmp_path):
    app = create_app(
        {
            "TESTING": True,
            "SECRET_KEY": "test-secret",
            "DATABASE": str(tmp_path / "test.sqlite3"),
            "ADMIN_USERNAME": "admin",
            "ADMIN_PASSWORD": "test-password",
        }
    )
    with app.app_context():
        workbook_bytes = SEED_WORKBOOK.read_bytes()
        config = parse_workbook(workbook_bytes)
        _create_or_update_event(
            config, workbook_bytes, SEED_WORKBOOK.name, "Initial test revision"
        )
    return app


@pytest.fixture()
def client(app):
    return app.test_client()


def test_public_agenda_and_info_are_accessible_without_login(client):
    agenda = client.get("/ceha/aes-2026?day=2026-04-23")
    info = client.get("/ceha/aes-2026/info")

    assert agenda.status_code == 200
    assert b"Practical AI Recipes for EH Professionals" in agenda.data
    assert b"Config v1" in agenda.data
    assert info.status_code == 200
    assert b"Long Beach" in info.data


def test_keyword_search_scans_all_days_and_labels_result_dates(client):
    response = client.get("/ceha/aes-2026?day=2026-04-23&q=Shannon")

    assert b"Introduction to Hazard Analysis Critical Control Point" in response.data
    assert b"Introduction to Risk Based Inspection" in response.data
    assert b"Monday, April 20" in response.data
    assert b"Tuesday, April 21" in response.data
    assert response.data.count(b'class="search-day-heading"') == 2
    assert b"Search results across all conference days" in response.data


def test_event_outside_date_range_defaults_to_first_day(client):
    response = client.get("/ceha/aes-2026")

    assert b"Monday, April 20" in response.data
    assert b"Pool Plan Check" in response.data
    assert b"Show finished sessions" not in response.data


def test_active_event_defaults_to_today():
    days = ["2026-04-20", "2026-04-21", "2026-04-22"]
    conference = {"start_date": "2026-04-20", "end_date": "2026-04-22"}

    assert _selected_event_day(days, "", "2026-04-21", conference) == "2026-04-21"
    assert _selected_event_day(days, "", "2026-04-25", conference) == "2026-04-20"


def test_active_event_renders_current_time_scroll_marker(client, monkeypatch):
    monkeypatch.setattr(
        routes,
        "_now_for_event",
        lambda _config: datetime(
            2026, 4, 23, 10, 55, tzinfo=ZoneInfo("America/Los_Angeles")
        ),
    )

    response = client.get("/ceha/aes-2026")

    assert b"Thursday, April 23" in response.data
    assert b'data-current-time-target' in response.data
    assert b"10:55 AM" in response.data
    assert b'scrollIntoView({ block: "center" })' in response.data


def test_restore_creates_a_new_revision(app):
    with app.app_context():
        from conference_agenda.db import get_db

        db = get_db()
        event = db.execute("SELECT * FROM events").fetchone()
        first = db.execute("SELECT * FROM revisions WHERE event_id = ?", (event["id"],)).fetchone()
        config = parse_workbook(first["workbook"])
        _create_revision(
            event["id"],
            config,
            first["workbook"],
            first["filename"],
            "Second revision",
        )
        new_version = _create_revision(
            event["id"],
            config,
            first["workbook"],
            first["filename"],
            "Restored from version 1",
            source_revision_id=first["id"],
        )
        revisions = db.execute(
            "SELECT version, source_revision_id FROM revisions WHERE event_id = ? ORDER BY version",
            (event["id"],),
        ).fetchall()

    assert new_version == 3
    assert [row["version"] for row in revisions] == [1, 2, 3]
    assert revisions[-1]["source_revision_id"] == first["id"]
