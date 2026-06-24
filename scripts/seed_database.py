from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = ROOT / "seed_data" / "ceha-aes-2026.xlsx"
sys.path.insert(0, str(ROOT))

from conference_agenda import create_app
from conference_agenda.routes import _create_or_update_event
from conference_agenda.spreadsheet import parse_workbook


def main():
    app = create_app()
    with app.app_context():
        workbook_bytes = WORKBOOK.read_bytes()
        config = parse_workbook(workbook_bytes)
        event = _create_or_update_event(
            config,
            workbook_bytes,
            WORKBOOK.name,
            "Initial schedule extracted from the 2026 AES program PDF.",
        )
        print(f"Seeded /{event['organization_slug']}/{event['event_slug']}")


if __name__ == "__main__":
    main()
