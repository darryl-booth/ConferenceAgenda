from flask import Flask
from datetime import datetime

from .db import init_app

APP_VERSION = "0.3.0"


def create_app(test_config=None):
    app = Flask(__name__, instance_relative_config=True)
    app.config.from_mapping(
        SECRET_KEY="change-me-in-production",
        DATABASE=app.instance_path + "/conference_agenda.sqlite3",
        ADMIN_USERNAME="admin",
        ADMIN_PASSWORD="change-me",
        MAX_CONTENT_LENGTH=10 * 1024 * 1024,
    )
    app.config.from_prefixed_env()

    if test_config:
        app.config.update(test_config)

    import os

    os.makedirs(app.instance_path, exist_ok=True)
    init_app(app)

    from .routes import bp

    app.register_blueprint(bp)

    @app.context_processor
    def inject_global_values():
        return {"app_version": APP_VERSION}

    @app.template_filter("friendly_date")
    def friendly_date(value):
        if not value:
            return ""
        return datetime.fromisoformat(str(value)).strftime("%A, %B %-d")

    @app.template_filter("friendly_time")
    def friendly_time(value):
        if not value:
            return ""
        return datetime.strptime(str(value), "%H:%M").strftime("%-I:%M %p")

    return app
