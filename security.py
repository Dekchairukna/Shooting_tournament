"""Session and request protection shared by all application pages."""
import os
import secrets
from pathlib import Path

from flask import abort, jsonify, request, session


def configure_security(app):
    secret = os.environ.get("SECRET_KEY")
    if not secret:
        if os.environ.get("DATABASE_URL", "").startswith(("postgres",)):
            raise RuntimeError("Set SECRET_KEY before starting the PostgreSQL deployment")
        path = Path(app.instance_path) / ".session-secret"
        path.parent.mkdir(parents=True, exist_ok=True)
        try:
            fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        except FileExistsError:
            pass
        else:
            with os.fdopen(fd, "w") as stream:
                stream.write(secrets.token_hex(32))
        secret = path.read_text().strip()
    if not secret or secret == "dev-secret-change-me":
        raise RuntimeError("SECRET_KEY must be a private, non-default value")
    app.config.update(SECRET_KEY=secret, SESSION_COOKIE_HTTPONLY=True,
                      SESSION_COOKIE_SAMESITE="Lax",
                      SESSION_COOKIE_SECURE=os.environ.get("SESSION_COOKIE_SECURE") == "1",
                      MAX_CONTENT_LENGTH=16 * 1024 * 1024)

    def csrf_token():
        if "csrf_token" not in session:
            session["csrf_token"] = secrets.token_urlsafe(32)
        return session["csrf_token"]

    app.jinja_env.globals["csrf_token"] = csrf_token

    @app.before_request
    def protect_mutations():
        if request.method in {"POST", "PUT", "PATCH", "DELETE"}:
            expected = session.get("csrf_token", "")
            supplied = request.headers.get("X-CSRF-Token") or request.form.get("csrf_token", "")
            if not expected or not secrets.compare_digest(expected, supplied):
                if request.is_json:
                    return jsonify(ok=False, message="เซสชันหมดอายุ กรุณาโหลดหน้าใหม่ก่อนบันทึก"), 400
                abort(400, description="เซสชันหมดอายุ กรุณาโหลดหน้าใหม่ก่อนบันทึก")

    @app.after_request
    def security_headers(response):
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["Referrer-Policy"] = "same-origin"
        if request.endpoint != "static":
            response.headers["Cache-Control"] = "no-store"
        return response
