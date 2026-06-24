import json
import secrets
from datetime import datetime
from functools import wraps
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo

from flask import (
    Blueprint,
    abort,
    current_app,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)

from .db import get_db
from .spreadsheet import WorkbookValidationError, parse_workbook

bp = Blueprint("main", __name__)


def _csrf_token():
    if "_csrf_token" not in session:
        session["_csrf_token"] = secrets.token_urlsafe(32)
    return session["_csrf_token"]


def _validate_csrf():
    expected = session.get("_csrf_token")
    supplied = request.form.get("_csrf_token", "")
    if not expected or not secrets.compare_digest(expected, supplied):
        abort(400, "Invalid form token. Refresh the page and try again.")


@bp.app_context_processor
def inject_csrf_token():
    return {"csrf_token": _csrf_token}


def admin_required(view):
    @wraps(view)
    def wrapped_view(**kwargs):
        if not session.get("admin_authenticated"):
            return redirect(url_for("main.admin_login", next=request.path))
        return view(**kwargs)

    return wrapped_view


def _event_row(org_slug, event_slug):
    return get_db().execute(
        """
        SELECT events.*, organizations.name AS organization_name,
               organizations.slug AS organization_slug
        FROM events
        JOIN organizations ON organizations.id = events.organization_id
        WHERE organizations.slug = ? AND events.slug = ?
        """,
        (org_slug, event_slug),
    ).fetchone()


def _event_config(event):
    if not event or not event["active_revision_id"]:
        return None, None
    revision = get_db().execute(
        "SELECT * FROM revisions WHERE id = ?", (event["active_revision_id"],)
    ).fetchone()
    return revision, json.loads(revision["config_json"])


def _public_event_or_404(org_slug, event_slug):
    event = _event_row(org_slug, event_slug)
    revision, config = _event_config(event)
    if not event or not revision or not config:
        abort(404)
    return event, revision, config


def _now_for_event(config):
    return datetime.now(ZoneInfo(config["conference"]["timezone"]))


def _selected_event_day(event_days, requested_day, today, conference):
    if not event_days:
        return ""
    if requested_day == "today" and today in event_days:
        return today
    if requested_day in event_days:
        return requested_day

    event_is_active = conference["start_date"] <= today <= conference["end_date"]
    if event_is_active and today in event_days:
        return today
    return event_days[0]


@bp.get("/")
def index():
    events = get_db().execute(
        """
        SELECT events.name, events.slug, organizations.name AS organization_name,
               organizations.slug AS organization_slug
        FROM events
        JOIN organizations ON organizations.id = events.organization_id
        WHERE events.active_revision_id IS NOT NULL
        ORDER BY organizations.name, events.name
        """
    ).fetchall()
    if len(events) == 1:
        event = events[0]
        return redirect(
            url_for(
                "main.agenda",
                org_slug=event["organization_slug"],
                event_slug=event["slug"],
            )
        )
    return render_template("event_list.html", events=events)


@bp.get("/<org_slug>/<event_slug>")
def agenda(org_slug, event_slug):
    event, revision, config = _public_event_or_404(org_slug, event_slug)
    conference = config["conference"]
    now = _now_for_event(config)
    requested_day = request.args.get("day", "").strip()
    query = request.args.get("q", "").strip()

    event_days = sorted({item["date"] for item in config["sessions"] if item["visible"]})
    today = now.date().isoformat()
    selected_day = _selected_event_day(
        event_days, requested_day, today, conference
    )

    sessions = []
    search_terms = query.casefold().split()
    for item in config["sessions"]:
        if not item["visible"] or item["date"] != selected_day:
            continue
        searchable = " ".join(
            [
                item["session_name"],
                item["session_number"],
                item["long_description"],
                item["presenters"],
                item["room"],
                item["track"],
                item["session_type"],
            ]
        ).casefold()
        if search_terms and not all(term in searchable for term in search_terms):
            continue

        end_at = datetime.fromisoformat(f"{item['date']}T{item['end_time']}").replace(
            tzinfo=now.tzinfo
        )
        item_copy = dict(item)
        item_copy["is_past"] = end_at < now
        sessions.append(item_copy)

    scroll_target_index = None
    if selected_day == today and not query:
        scroll_target_index = next(
            (
                index
                for index, item in enumerate(sessions)
                if datetime.fromisoformat(
                    f"{item['date']}T{item['start_time']}"
                ).replace(tzinfo=now.tzinfo)
                >= now
            ),
            len(sessions),
        )

    return render_template(
        "agenda.html",
        event=event,
        revision=revision,
        conference=conference,
        sessions=sessions,
        event_days=event_days,
        selected_day=selected_day,
        today=today,
        query=query,
        scroll_target_index=scroll_target_index,
        current_time_label=now.strftime("%-I:%M %p"),
        now=now,
    )


@bp.get("/<org_slug>/<event_slug>/info")
def event_info(org_slug, event_slug):
    event, revision, config = _public_event_or_404(org_slug, event_slug)
    return render_template(
        "event_info.html",
        event=event,
        revision=revision,
        conference=config["conference"],
    )


@bp.route("/admin/login", methods=("GET", "POST"))
def admin_login():
    if request.method == "POST":
        _validate_csrf()
        username_ok = secrets.compare_digest(
            request.form.get("username", ""), current_app.config["ADMIN_USERNAME"]
        )
        password_ok = secrets.compare_digest(
            request.form.get("password", ""), current_app.config["ADMIN_PASSWORD"]
        )
        if username_ok and password_ok:
            session.clear()
            session["admin_authenticated"] = True
            session["_csrf_token"] = secrets.token_urlsafe(32)
            target = request.args.get("next", "")
            if not target.startswith("/"):
                target = url_for("main.admin_dashboard")
            return redirect(target)
        flash("The username or password was not recognized.", "error")
    return render_template("admin/login.html")


@bp.post("/admin/logout")
@admin_required
def admin_logout():
    _validate_csrf()
    session.clear()
    return redirect(url_for("main.admin_login"))


@bp.get("/admin")
@admin_required
def admin_dashboard():
    events = get_db().execute(
        """
        SELECT events.*, organizations.name AS organization_name,
               organizations.slug AS organization_slug,
               revisions.version, revisions.created_at
        FROM events
        JOIN organizations ON organizations.id = events.organization_id
        LEFT JOIN revisions ON revisions.id = events.active_revision_id
        ORDER BY organizations.name, events.name
        """
    ).fetchall()
    return render_template("admin/dashboard.html", events=events)


@bp.get("/admin/template")
@admin_required
def admin_download_template():
    workbook_path = (
        Path(current_app.root_path).parent / "seed_data" / "ceha-aes-2026.xlsx"
    )
    if not workbook_path.exists():
        abort(404)
    return send_file(
        workbook_path,
        as_attachment=True,
        download_name="conference-agenda-template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@bp.route("/admin/new", methods=("GET", "POST"))
@admin_required
def admin_new_event():
    if request.method == "POST":
        _validate_csrf()
        upload = request.files.get("workbook")
        if not upload or not upload.filename:
            flash("Choose an Excel workbook to create the event.", "error")
            return render_template("admin/new_event.html")
        workbook_bytes = upload.read()
        try:
            config = parse_workbook(workbook_bytes)
            event = _create_or_update_event(
                config, workbook_bytes, upload.filename, request.form.get("note", "")
            )
        except WorkbookValidationError as exc:
            for error in exc.errors:
                flash(error, "error")
            return render_template("admin/new_event.html")
        except ValueError as exc:
            flash(str(exc), "error")
            return render_template("admin/new_event.html")
        flash("Event created and revision 1 is live.", "success")
        return redirect(
            url_for(
                "main.admin_event",
                org_slug=event["organization_slug"],
                event_slug=event["event_slug"],
            )
        )
    return render_template("admin/new_event.html")


@bp.get("/admin/<org_slug>/<event_slug>")
@admin_required
def admin_event(org_slug, event_slug):
    event = _event_row(org_slug, event_slug)
    if not event:
        abort(404)
    revision, config = _event_config(event)
    revisions = get_db().execute(
        "SELECT * FROM revisions WHERE event_id = ? ORDER BY version DESC", (event["id"],)
    ).fetchall()
    return render_template(
        "admin/event.html",
        event=event,
        revision=revision,
        config=config,
        revisions=revisions,
    )


@bp.post("/admin/<org_slug>/<event_slug>/upload")
@admin_required
def admin_upload(org_slug, event_slug):
    _validate_csrf()
    event = _event_row(org_slug, event_slug)
    if not event:
        abort(404)
    upload = request.files.get("workbook")
    if not upload or not upload.filename:
        flash("Choose an Excel workbook to upload.", "error")
        return redirect(url_for("main.admin_event", org_slug=org_slug, event_slug=event_slug))

    workbook_bytes = upload.read()
    try:
        config = parse_workbook(workbook_bytes)
        if (
            config["conference"]["organization_slug"] != org_slug
            or config["conference"]["event_slug"] != event_slug
        ):
            raise WorkbookValidationError(
                [
                    "The organization_slug and event_slug must match this event. "
                    "Create a new event if you need different URLs."
                ]
            )
        _create_revision(
            event["id"],
            config,
            workbook_bytes,
            upload.filename,
            request.form.get("note", ""),
        )
    except WorkbookValidationError as exc:
        for error in exc.errors:
            flash(error, "error")
        return redirect(url_for("main.admin_event", org_slug=org_slug, event_slug=event_slug))

    flash("The new revision is now live.", "success")
    return redirect(url_for("main.admin_event", org_slug=org_slug, event_slug=event_slug))


@bp.get("/admin/<org_slug>/<event_slug>/download")
@admin_required
def admin_download(org_slug, event_slug):
    event = _event_row(org_slug, event_slug)
    revision, _config = _event_config(event)
    if not event or not revision:
        abort(404)
    return send_file(
        BytesIO(revision["workbook"]),
        as_attachment=True,
        download_name=f"{org_slug}-{event_slug}-v{revision['version']}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@bp.get("/admin/<org_slug>/<event_slug>/revisions/<int:revision_id>/download")
@admin_required
def admin_download_revision(org_slug, event_slug, revision_id):
    event = _event_row(org_slug, event_slug)
    if not event:
        abort(404)
    revision = get_db().execute(
        "SELECT * FROM revisions WHERE id = ? AND event_id = ?",
        (revision_id, event["id"]),
    ).fetchone()
    if not revision:
        abort(404)
    return send_file(
        BytesIO(revision["workbook"]),
        as_attachment=True,
        download_name=f"{org_slug}-{event_slug}-v{revision['version']}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@bp.post("/admin/<org_slug>/<event_slug>/revisions/<int:revision_id>/restore")
@admin_required
def admin_restore_revision(org_slug, event_slug, revision_id):
    _validate_csrf()
    event = _event_row(org_slug, event_slug)
    if not event:
        abort(404)
    revision = get_db().execute(
        "SELECT * FROM revisions WHERE id = ? AND event_id = ?",
        (revision_id, event["id"]),
    ).fetchone()
    if not revision:
        abort(404)
    config = json.loads(revision["config_json"])
    new_version = _create_revision(
        event["id"],
        config,
        revision["workbook"],
        revision["filename"],
        f"Restored from version {revision['version']}",
        source_revision_id=revision["id"],
    )
    flash(f"Version {revision['version']} was restored as version {new_version}.", "success")
    return redirect(url_for("main.admin_event", org_slug=org_slug, event_slug=event_slug))


def _create_or_update_event(config, workbook_bytes, filename, note):
    db = get_db()
    conference = config["conference"]
    organization = db.execute(
        "SELECT * FROM organizations WHERE slug = ?",
        (conference["organization_slug"],),
    ).fetchone()
    if organization:
        organization_id = organization["id"]
    else:
        cursor = db.execute(
            "INSERT INTO organizations (name, slug) VALUES (?, ?)",
            (conference["organization_name"], conference["organization_slug"]),
        )
        organization_id = cursor.lastrowid

    existing = db.execute(
        "SELECT id FROM events WHERE organization_id = ? AND slug = ?",
        (organization_id, conference["event_slug"]),
    ).fetchone()
    if existing:
        raise ValueError(
            "That organization and event URL already exist. Upload from its admin page instead."
        )

    cursor = db.execute(
        "INSERT INTO events (organization_id, name, slug) VALUES (?, ?, ?)",
        (organization_id, conference["event_name"], conference["event_slug"]),
    )
    event_id = cursor.lastrowid
    db.commit()
    _create_revision(event_id, config, workbook_bytes, filename, note)
    return {
        "id": event_id,
        "organization_slug": conference["organization_slug"],
        "event_slug": conference["event_slug"],
    }


def _create_revision(
    event_id,
    config,
    workbook_bytes,
    filename,
    note,
    source_revision_id=None,
):
    db = get_db()
    row = db.execute(
        "SELECT COALESCE(MAX(version), 0) + 1 AS next_version FROM revisions WHERE event_id = ?",
        (event_id,),
    ).fetchone()
    version = row["next_version"]
    cursor = db.execute(
        """
        INSERT INTO revisions
            (event_id, version, filename, workbook, config_json, note, source_revision_id)
        VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
        (
            event_id,
            version,
            filename,
            workbook_bytes,
            json.dumps(config),
            note.strip(),
            source_revision_id,
        ),
    )
    revision_id = cursor.lastrowid
    db.execute(
        "UPDATE events SET active_revision_id = ?, name = ? WHERE id = ?",
        (revision_id, config["conference"]["event_name"], event_id),
    )
    db.execute(
        """
        UPDATE organizations
        SET name = ?
        WHERE id = (SELECT organization_id FROM events WHERE id = ?)
        """,
        (config["conference"]["organization_name"], event_id),
    )
    db.commit()
    return version
