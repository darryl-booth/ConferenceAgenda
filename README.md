# Conference Agenda

A small Flask application for publishing a mobile-friendly conference agenda from an Excel workbook. Attendees do not need accounts. Organizers upload a replacement workbook from the admin page, and every accepted upload remains available in revision history.

## What it includes

- Readable public URLs such as `/ceha/aes-2026`
- Day navigation, Today shortcut, keyword search, automatic current-time positioning,
  and subdued styling for completed sessions
- Session details including number, description, presenters, room, capacity, resources, track, and CECH
- Branded conference information page
- Excel template download and current-configuration download
- Validated replacement uploads with immutable revision history and rollback
- SQLite storage
- Responsive Chrome, Safari, iPhone, and Android layout
- Web-app manifest for adding the site to a home screen
- App and configuration versions in every page footer

## Run locally

Activate the environment and install dependencies:

```bash
source .venv/bin/activate
pip install -r requirements.txt
```

Seed the local database with the PDF-derived workbook:

```bash
python scripts/seed_database.py
```

Set development credentials and start Flask:

```bash
export FLASK_SECRET_KEY="$(python -c 'import secrets; print(secrets.token_hex(32))')"
export FLASK_ADMIN_USERNAME="admin"
export FLASK_ADMIN_PASSWORD="choose-a-strong-password"
flask --app wsgi run --debug
```

Open `http://127.0.0.1:5000/ceha/aes-2026`. The admin area is at `http://127.0.0.1:5000/admin`.

## Workbook contract

The workbook has two required sheets:

- `Conference Info`: field names in column A and values in column B
- `Sessions`: one agenda activity per row

Required session columns are `session_name`, `date`, `start_time`, and `end_time`. The supplied workbook includes all supported columns, editing instructions, and source notes.

The public URL is built from `organization_slug` and `event_slug`. Once an event exists, those two values cannot be changed through a replacement upload; create a new event for a new URL.

## Tests

```bash
pytest
```

## PythonAnywhere

1. Upload or clone this project into your PythonAnywhere account.
2. Create a virtual environment and run `pip install -r requirements.txt`.
3. Run `python scripts/seed_database.py` once if you want the sample event.
4. Create a new PythonAnywhere web app using manual configuration.
5. Set the source directory to this project and the virtual environment path to its `.venv`.
6. In the WSGI configuration file, add the project directory to `sys.path`, then import `app` from `wsgi`.
7. Set `FLASK_SECRET_KEY`, `FLASK_ADMIN_USERNAME`, and `FLASK_ADMIN_PASSWORD` in the WSGI file before importing the app, or configure equivalent environment variables.
8. Reload the web app.

The SQLite database is stored in `instance/conference_agenda.sqlite3`. Include that file in backups. Only one organizer should upload at a time, which is a good fit for the intended lightweight use and PythonAnywhere deployment.
