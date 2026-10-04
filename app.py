from __future__ import annotations

import json
import os
import time
from io import BytesIO
from datetime import date, datetime
from functools import wraps
from typing import Dict, List, Tuple
from types import SimpleNamespace
from flask import jsonify

from flask import (
    Flask,
    abort,
    flash,
    g,
    has_request_context,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from flask_login import (
    LoginManager,
    UserMixin,
    current_user,
    login_required,
    login_user,
    logout_user,
)
from flask_sqlalchemy import SQLAlchemy
from sqlalchemy import func
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename
from openpyxl import Workbook, load_workbook

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, "instance", "shooting.db")

# ให้ SQLite เขียนได้หลังแตกไฟล์ ZIP บน Mac/Windows
def ensure_sqlite_writable():
    if os.environ.get("DATABASE_URL"):
        return
    try:
        os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
        os.chmod(os.path.dirname(DB_PATH), 0o755)
        if os.path.exists(DB_PATH):
            os.chmod(DB_PATH, 0o666)
    except Exception:
        # ไม่ให้ระบบล่มเพราะ chmod บนบาง host ไม่รองรับ
        pass

ensure_sqlite_writable()

app = Flask(__name__)
def _load_secret_key() -> str:
    """ใช้ SECRET_KEY จาก environment ก่อน ถ้าไม่มีให้สร้างคีย์สุ่มเก็บไว้ใน instance/secret_key
    (ทุก gunicorn worker ในเครื่องเดียวกันอ่านไฟล์เดียวกัน session จึงไม่หลุดระหว่าง worker)
    ไม่ใช้ค่าตายตัวที่เดาได้อีกต่อไป เพราะคนที่รู้ค่าจะปลอม session เป็น superadmin ได้"""
    key = os.environ.get("SECRET_KEY", "").strip()
    if key and key != "dev-secret-change-me":
        return key
    import secrets
    path = os.path.join(BASE_DIR, "instance", "secret_key")
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        if os.path.exists(path):
            with open(path) as fh:
                stored = fh.read().strip()
            if stored:
                return stored
        new_key = secrets.token_hex(32)
        fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
        with os.fdopen(fd, "w") as fh:
            fh.write(new_key)
        return new_key
    except FileExistsError:
        with open(path) as fh:
            return fh.read().strip()
    except OSError:
        print("[SECURITY] ไม่มี SECRET_KEY และเขียนไฟล์คีย์ไม่ได้ ใช้คีย์ชั่วคราว (ทุกคนต้อง login ใหม่เมื่อรีสตาร์ต)")
        return secrets.token_hex(32)


app.config["SECRET_KEY"] = _load_secret_key()
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = os.environ.get("COOKIE_SECURE", "1" if os.environ.get("RAILWAY_ENVIRONMENT") else "0") == "1"

database_url = os.environ.get("DATABASE_URL")
if database_url:
    if database_url.startswith("postgres://"):
        database_url = database_url.replace("postgres://", "postgresql://", 1)
    app.config["SQLALCHEMY_DATABASE_URI"] = database_url
else:
    app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DB_PATH}"

app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

# --- CSRF: กันเว็บอื่นแอบสั่งลบ/แก้คะแนนผ่านเบราว์เซอร์ของเจ้าหน้าที่ที่ login ค้างไว้
from flask_wtf.csrf import CSRFProtect, CSRFError, generate_csrf
# หน้า Scorecard เปิดค้างได้ทั้งวัน จึงผูก token กับ session แทนการหมดอายุทุก 1 ชั่วโมง
app.config["WTF_CSRF_TIME_LIMIT"] = None
app.config["WTF_CSRF_SSL_STRICT"] = False  # Railway อยู่หลัง proxy; token + SameSite cookie ยังป้องกันอยู่
csrf = CSRFProtect(app)


@app.context_processor
def inject_csrf_token():
    return {"csrf_token": generate_csrf}


# หน้า Overview/Bracket ดึงข้อมูลซ้ำตลอดเวลา ถ้าข้อมูลไม่เปลี่ยนให้ตอบ 304 (ไม่มีเนื้อหา)
# ลดข้อมูลที่ส่งจากราว 56 KB ต่อครั้งเหลือไม่กี่ร้อยไบต์ และเบราว์เซอร์ไม่ต้องวาดตารางใหม่
POLLING_ENDPOINTS = {"overview_data", "overview_stats", "bracket_data"}


# --- แคชร่วมระยะสั้นสำหรับหน้าที่คนดูพร้อมกันเยอะ ---
# ผู้ชม 100 คนดูอีเวนต์เดียวกัน เดิมเซิร์ฟเวอร์คำนวณตารางใหม่ 100 รอบ ตอนนี้คำนวณครั้งเดียวแล้วส่งผลเดียวกันให้ทุกคน
# ข้อมูลช้าได้ไม่เกิน SHARED_POLL_CACHE_SECONDS และจะล้างทันทีเมื่อมีการบันทึกคะแนนของอีเวนต์นั้น (ใน process เดียวกัน)
import threading
SHARED_POLL_CACHE_SECONDS = float(os.environ.get("SHARED_POLL_CACHE_SECONDS", "1.0"))
_poll_cache: dict = {}
_poll_cache_locks: dict = {}
_poll_cache_guard = threading.Lock()


def invalidate_poll_cache(event_id: int) -> None:
    with _poll_cache_guard:
        for key in [k for k in _poll_cache if k[1] == event_id]:
            _poll_cache.pop(key, None)


def shared_poll_cache(view):
    @wraps(view)
    def wrapper(event_id: int, *args, **kwargs):
        if SHARED_POLL_CACHE_SECONDS <= 0:
            return view(event_id, *args, **kwargs)
        # บัญชีสนามเห็นเฉพาะสนามตัวเอง จึงแยกแคชตามสนาม ส่วนคนอื่นเห็นข้อมูลชุดเดียวกัน
        scope = (current_user.court_event_id, current_user.court_no) if is_court_user() else None
        key = (view.__name__, event_id, request.query_string.decode(), scope)
        now = time.monotonic()
        hit = _poll_cache.get(key)
        if hit and now - hit[0] < SHARED_POLL_CACHE_SECONDS:
            return app.response_class(hit[1], status=200, mimetype="application/json")
        with _poll_cache_guard:
            lock = _poll_cache_locks.setdefault(key, threading.Lock())
        with lock:  # คำขอที่มาพร้อมกันรอผลจากการคำนวณครั้งเดียว ไม่แย่งกันคำนวณ
            hit = _poll_cache.get(key)
            if hit and time.monotonic() - hit[0] < SHARED_POLL_CACHE_SECONDS:
                return app.response_class(hit[1], status=200, mimetype="application/json")
            response = app.make_response(view(event_id, *args, **kwargs))
            if response.status_code == 200 and response.mimetype == "application/json":
                _poll_cache[key] = (time.monotonic(), response.get_data())
            return response
    return wrapper


@app.after_request
def conditional_polling_response(response):
    if request.method not in ("GET", "HEAD", "OPTIONS") and response.status_code < 400 and _poll_cache:
        # มีการบันทึก/แก้ไขใด ๆ: ล้างแคชทันที ให้คำขอถัดไปได้ข้อมูลล่าสุด
        with _poll_cache_guard:
            _poll_cache.clear()
    if (request.endpoint in POLLING_ENDPOINTS and request.method == "GET"
            and response.status_code == 200 and not response.direct_passthrough):
        response.headers["Cache-Control"] = "private, no-cache"
        response.add_etag()
        response.make_conditional(request)
        # บีบอัด JSON (56 KB → ราว 5 KB) ลดเน็ตของผู้ชมและค่า egress ของเซิร์ฟเวอร์
        if (response.status_code == 200 and "gzip" in request.headers.get("Accept-Encoding", "")
                and not response.headers.get("Content-Encoding")):
            raw = response.get_data()
            if len(raw) > 1024:
                import gzip
                response.set_data(gzip.compress(raw, compresslevel=5))
                response.headers["Content-Encoding"] = "gzip"
                response.headers["Vary"] = "Accept-Encoding"
    return response


@app.errorhandler(CSRFError)
def handle_csrf_error(err):
    message = "หน้าเว็บหมดอายุหรือเปิดค้างจากการ login ครั้งก่อน กรุณารีเฟรชหน้าแล้วลองใหม่"
    if request.is_json or request.path.startswith("/api/"):
        return jsonify({"ok": False, "message": message}), 400
    flash(message, "warning")
    return redirect(request.referrer or url_for("index"))

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = "login"

LANG_LABELS = {
    "th": {
        "name": "ไทย",
        "shooting_title": "ประเภทสุดยอดความแม่นยำ (SHOOTING)",
    },
    "en": {
        "name": "English",
        "shooting_title": "Precision Shooting",
    },
    "fr": {
        "name": "Français",
        "shooting_title": "Tir de précision",
    },
    "zh": {
        "name": "中文",
        "shooting_title": "精准射击",
    },
}
SUPPORTED_LANGS = tuple(LANG_LABELS.keys())


def current_language() -> str:
    lang = session.get("lang", "th")
    return lang if lang in SUPPORTED_LANGS else "th"


@app.context_processor
def inject_language_options():
    lang = current_language()
    return {
        "current_lang": lang,
        "lang_labels": LANG_LABELS,
        "html_lang": "zh-CN" if lang == "zh" else lang,
    }

ROUND_LABELS = {
    1: "รอบที่ 1",
    2: "รอบที่ 2",
}

SCORECARD_ROUND_LABELS_8 = [
    "รอบที่ 1",
    "รอบที่ 2",
    "รอบ 8 คน",
    "รอบรองชนะเลิศ",
    "รอบชิงชนะเลิศ",
]

SCORECARD_ROUND_LABELS_16 = [
    "รอบที่ 1",
    "รอบที่ 2",
    "รอบ 16 คน",
    "รอบ 8 คน",
    "รอบรองชนะเลิศ",
    "รอบชิงชนะเลิศ",
]

ALL_SCORECARD_ROUNDS = [1, 2, 3, 4, 5, 6]

STATIONS = [1, 2, 3, 4, 5]
DISTANCES = [6, 7, 8, 9]
MAX_RED_CARDS = 2

# คะแนนที่ถูกต้องตามกติกายิงเปตอง
# สถานี 1-4: Carreau 5 / Réussi 3 / Touché 1 / Manqué 0
# สถานี 5 (But): Carreau 5 / Touché 3 / Manqué 0
ALLOWED_SCORES_BY_STATION = {1: {0, 1, 3, 5}, 2: {0, 1, 3, 5}, 3: {0, 1, 3, 5}, 4: {0, 1, 3, 5}, 5: {0, 3, 5}}


class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False, default="user")
    # role=court: ผูกบัญชีกับหมายเลขสนาม เช่น court03 -> สนาม 3
    court_no = db.Column(db.Integer, nullable=True)
    # Court account is scoped to one event. Internal username stays globally unique.
    court_event_id = db.Column(db.Integer, db.ForeignKey("event.id"), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def set_password(self, password: str) -> None:
        self.password_hash = generate_password_hash(password)

    def has_usable_password(self) -> bool:
        return bool(self.password_hash and self.password_hash != "!UNSET!")

    def check_password(self, password: str) -> bool:
        if not self.has_usable_password():
            return False
        return check_password_hash(self.password_hash, password)


class Event(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(255), nullable=False)
    event_group = db.Column(db.String(50), nullable=False)
    category = db.Column(db.String(50), nullable=False)
    competition_date = db.Column(db.Date, nullable=False)
    location = db.Column(db.String(255), nullable=False)
    lane_count = db.Column(db.Integer, nullable=False, default=1)
    direct_qualifiers = db.Column(db.Integer, nullable=False, default=0)
    has_round_two = db.Column(db.Boolean, default=False)
    round_two_cutoff_rank = db.Column(db.Integer, nullable=True)
    # วิธีคัดเข้ารอบ 2: "cutoff" = ถึง Class ที่ N · "next" = ต่อจากผู้ผ่านตรงอีก N ลำดับ
    round_two_mode = db.Column(db.String(20), nullable=False, default="cutoff")
    next_round_label = db.Column(db.String(50), nullable=False, default="รอบ 8 คน")
    round_two_advancers = db.Column(db.Integer, nullable=False, default=4)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    created_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)

    athletes = db.relationship("Athlete", backref="event", cascade="all, delete-orphan", lazy=True)


class Athlete(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    event_id = db.Column(db.Integer, db.ForeignKey("event.id"), nullable=False)
    bib_no = db.Column(db.String(20), nullable=False)
    name = db.Column(db.String(255), nullable=False)
    affiliation = db.Column(db.String(255), nullable=False)
    start_order = db.Column(db.Integer, nullable=False)
    lane_no = db.Column(db.Integer, nullable=False)
    lane_order = db.Column(db.Integer, nullable=False)
    status = db.Column(db.String(20), nullable=False, default="waiting")
    red_card_count = db.Column(db.Integer, nullable=False, default=0)
    # ผู้ดูแลสามารถปิดสิทธิ์รอบ 2 รายคนได้ โดยไม่ลบคะแนน/อันดับรอบ 1
    round_two_disabled = db.Column(db.Boolean, nullable=False, default=False)
    round_two_disabled_at = db.Column(db.DateTime, nullable=True)
    round_two_disabled_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    entries = db.relationship("ScoreEntry", backref="athlete", cascade="all, delete-orphan", lazy=True)
    signatures = db.relationship("ScoreSignature", backref="athlete", cascade="all, delete-orphan", lazy=True)
    tiebreaks = db.relationship("TieBreakEntry", backref="athlete", cascade="all, delete-orphan", lazy=True)


class ScoreEntry(db.Model):
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    athlete_id = db.Column(db.Integer, db.ForeignKey("athlete.id"), nullable=False)
    round_no = db.Column(db.Integer, nullable=False, default=1)
    station_no = db.Column(db.Integer, nullable=False)
    distance_m = db.Column(db.Integer, nullable=False)
    score = db.Column(db.Integer, nullable=False, default=0)
    is_red_card = db.Column(db.Boolean, nullable=False, default=False)
    # แยกสถานะ “ตีช่องนี้แล้ว” ออกจากคะแนน เพื่อแยก 0 คะแนนจริงกับช่องที่ยังไม่ได้ตี
    is_scored = db.Column(db.Boolean, nullable=False, default=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


class ScoreSignature(db.Model):
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    athlete_id = db.Column(db.Integer, db.ForeignKey("athlete.id"), nullable=False)
    round_no = db.Column(db.Integer, nullable=False)
    recorder_name = db.Column(db.String(255), nullable=True)
    referee_name = db.Column(db.String(255), nullable=True)
    athlete_name = db.Column(db.String(255), nullable=True)
    recorder_signature = db.Column(db.Text, nullable=True)
    referee_signature = db.Column(db.Text, nullable=True)
    athlete_signature = db.Column(db.Text, nullable=True)
    bypass_signed = db.Column(db.Boolean, default=False)
    started_at = db.Column(db.DateTime, nullable=True)
    finished_at = db.Column(db.DateTime, nullable=True)
    # รอบถูกยุติอัตโนมัติเมื่อได้รับใบแดงครบ 2 ครั้งตามกติกา Precision Shooting
    stopped_by_red = db.Column(db.Boolean, nullable=False, default=False)


class ScorecardLock(db.Model):
    """กันเจ้าหน้าที่สองเครื่องคีย์ Scorecard ใบเดียวกันพร้อมกันโดยไม่รู้ตัว
    เจ้าของล็อกคือหน้าเว็บหนึ่งหน้า (page_token) ต่ออายุด้วย heartbeat ทุก 30 วินาที
    ถ้าหน้าเดิมปิดไปหรือไม่ส่ง heartbeat เกิน SCORECARD_LOCK_TTL ล็อกหมดอายุเอง
    คนอื่นกด "รับช่วงคีย์ต่อ" ได้เสมอ (ระบบบันทึกชื่อไว้) จึงไม่มีทางติดล็อกค้าง"""
    __tablename__ = "scorecard_lock"
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    athlete_id = db.Column(db.Integer, db.ForeignKey("athlete.id"), nullable=False)
    round_no = db.Column(db.Integer, nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    username = db.Column(db.String(80), nullable=False, default="")
    page_token = db.Column(db.String(64), nullable=False, default="")
    heartbeat_at = db.Column(db.DateTime, nullable=False, default=datetime.utcnow)
    __table_args__ = (db.UniqueConstraint("athlete_id", "round_no", name="uq_scorecard_lock_athlete_round"),)


SCORECARD_LOCK_TTL_SECONDS = 90


def acquire_scorecard_lock(athlete_id: int, round_no: int, page_token: str, take_over: bool = False):
    """คืนค่า (ได้ล็อกหรือไม่, ชื่อผู้ถือล็อกปัจจุบัน)"""
    from datetime import timedelta
    from sqlalchemy.exc import IntegrityError
    page_token = (page_token or "")[:64]
    now = datetime.utcnow()
    who = court_public_id(current_user) if is_court_user() else current_user.username
    lock = ScorecardLock.query.filter_by(athlete_id=athlete_id, round_no=round_no).first()
    if lock:
        fresh = lock.heartbeat_at and (now - lock.heartbeat_at) < timedelta(seconds=SCORECARD_LOCK_TTL_SECONDS)
        mine = bool(page_token) and lock.page_token == page_token
        if fresh and not mine and not take_over:
            return False, lock.username
        lock.user_id = current_user.id
        lock.username = who
        lock.page_token = page_token
        lock.heartbeat_at = now
        db.session.commit()
        return True, who
    try:
        db.session.add(ScorecardLock(athlete_id=athlete_id, round_no=round_no, user_id=current_user.id,
                                     username=who, page_token=page_token, heartbeat_at=now))
        db.session.commit()
        return True, who
    except IntegrityError:
        # อีกเครื่องสร้างล็อกพร้อมกันพอดี: อ่านใหม่แล้วตัดสินตามปกติ
        db.session.rollback()
        return acquire_scorecard_lock(athlete_id, round_no, page_token, take_over)


def release_scorecard_lock(athlete_id: int, round_no: int, page_token: str) -> None:
    lock = ScorecardLock.query.filter_by(athlete_id=athlete_id, round_no=round_no).first()
    if lock and page_token and lock.page_token == page_token:
        db.session.delete(lock)
        db.session.commit()


class ScoreEditLog(db.Model):
    """ประวัติการแก้คะแนนหลังจากผลถูกลงลายเซ็นยืนยันแล้ว"""
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    athlete_id = db.Column(db.Integer, db.ForeignKey("athlete.id"), nullable=False)
    round_no = db.Column(db.Integer, nullable=False)
    station_no = db.Column(db.Integer, nullable=False)
    distance_m = db.Column(db.Integer, nullable=False)
    old_score = db.Column(db.Integer, nullable=False, default=0)
    new_score = db.Column(db.Integer, nullable=False, default=0)
    old_red = db.Column(db.Boolean, nullable=False, default=False)
    new_red = db.Column(db.Boolean, nullable=False, default=False)
    old_played = db.Column(db.Boolean, nullable=False, default=False)
    new_played = db.Column(db.Boolean, nullable=False, default=False)
    edited_by = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=True)
    editor_username = db.Column(db.String(80), nullable=True)
    editor_court_no = db.Column(db.Integer, nullable=True)
    editor_signature = db.Column(db.Text, nullable=True)
    reason = db.Column(db.String(500), nullable=True)
    edited_at = db.Column(db.DateTime, default=datetime.utcnow, nullable=False)


class TieBreakEntry(db.Model):
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    athlete_id = db.Column(db.Integer, db.ForeignKey("athlete.id"), nullable=False)
    round_no = db.Column(db.Integer, nullable=False)
    station_no = db.Column(db.Integer, nullable=False)
    score = db.Column(db.Integer, nullable=False, default=0)


class BracketMatch(db.Model):
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    event_id = db.Column(db.Integer, db.ForeignKey("event.id"), nullable=False)
    round_name = db.Column(db.String(20), nullable=False)
    match_no = db.Column(db.Integer, nullable=False)
    athlete_a_id = db.Column(db.Integer, nullable=True)
    athlete_b_id = db.Column(db.Integer, nullable=True)
    winner_id = db.Column(db.Integer, nullable=True)


class ResultsApprovedSetting(db.Model):
    id = db.Column(db.Integer, primary_key=True, autoincrement=True)
    event_id = db.Column(db.Integer, db.ForeignKey("event.id"), unique=True, nullable=False)
    competition_title = db.Column(db.String(255), nullable=True)
    host_line = db.Column(db.String(255), nullable=True)
    date_line = db.Column(db.String(255), nullable=True)
    location_line = db.Column(db.String(255), nullable=True)
    country_label = db.Column(db.String(80), nullable=False, default="COUNTRY")
    president_title = db.Column(db.String(255), nullable=True)
    president_name = db.Column(db.String(255), nullable=True)
    technical_title = db.Column(db.String(255), nullable=True)
    technical_name = db.Column(db.String(255), nullable=True)
    umpires_text = db.Column(db.Text, nullable=True)
    approved_text = db.Column(db.String(255), nullable=False, default="……………………………APPROVED")
    show_official_pages = db.Column(db.Boolean, nullable=False, default=True)
    # single = คอลัมน์ NAME เดียว (ค่าเริ่มต้น), first = คำแรกเป็นนามสกุล, last = คำสุดท้ายเป็นนามสกุล
    name_format = db.Column(db.String(10), nullable=False, default="single")
    cover_main_logo_path = db.Column(db.String(255), nullable=True)
    cover_bottom_logo_1_path = db.Column(db.String(255), nullable=True)
    cover_bottom_logo_2_path = db.Column(db.String(255), nullable=True)
    cover_bottom_logo_3_path = db.Column(db.String(255), nullable=True)
    header_logo_1_path = db.Column(db.String(255), nullable=True)
    header_logo_2_path = db.Column(db.String(255), nullable=True)
    header_logo_3_path = db.Column(db.String(255), nullable=True)
    header_logo_4_path = db.Column(db.String(255), nullable=True)
    side_logo_path = db.Column(db.String(255), nullable=True)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


class SiteTheme(db.Model):
    """ธีมหน้าตาของทั้งระบบ (ใช้งานได้ครั้งละ 1 ธีม)"""
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(120), nullable=False)
    short_name = db.Column(db.String(60), nullable=False, default="")
    eyebrow = db.Column(db.String(160), nullable=False, default="")
    title = db.Column(db.String(160), nullable=False, default="Petanque Shooting")
    subtitle = db.Column(db.String(255), nullable=False, default="")
    export_kicker = db.Column(db.String(160), nullable=False, default="Official Shooting Results")
    export_title = db.Column(db.String(255), nullable=False, default="")
    export_subtitle = db.Column(db.String(255), nullable=False, default="")
    footer_tagline = db.Column(db.String(255), nullable=False, default="")
    footer_event = db.Column(db.String(255), nullable=False, default="")
    show_hero = db.Column(db.Boolean, nullable=False, default=True)
    color_primary = db.Column(db.String(7), nullable=False, default="#ef4b12")
    color_primary_dark = db.Column(db.String(7), nullable=False, default="#c9360b")
    color_soft = db.Column(db.String(7), nullable=False, default="#fff1e8")
    color_cream = db.Column(db.String(7), nullable=False, default="#fffaf4")
    color_accent = db.Column(db.String(7), nullable=False, default="#37b8b1")
    color_ink = db.Column(db.String(7), nullable=False, default="#172033")
    color_line = db.Column(db.String(7), nullable=False, default="#f3d2bf")
    # ภาพตั้งต้นจากโฟลเดอร์ static (ธีมที่มากับระบบ) ถ้าอัปโหลดภาพใหม่จะใช้ภาพในฐานข้อมูลแทน
    poster_static = db.Column(db.String(255), nullable=True)
    logo_static = db.Column(db.String(255), nullable=True)
    partners_static = db.Column(db.String(255), nullable=True)
    is_active = db.Column(db.Boolean, nullable=False, default=False)
    is_builtin = db.Column(db.Boolean, nullable=False, default=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)

    assets = db.relationship("SiteThemeAsset", backref="theme", cascade="all, delete-orphan", lazy=True)


class SiteThemeAsset(db.Model):
    """ภาพของธีมเก็บในฐานข้อมูล เพื่อไม่หายเมื่อ Railway deploy ใหม่"""
    id = db.Column(db.Integer, primary_key=True)
    theme_id = db.Column(db.Integer, db.ForeignKey("site_theme.id"), nullable=False)
    kind = db.Column(db.String(20), nullable=False)  # poster / logo / partners
    mimetype = db.Column(db.String(40), nullable=False)
    data = db.deferred(db.Column(db.LargeBinary, nullable=False))  # โหลดเฉพาะตอนส่งภาพ
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    __table_args__ = (db.UniqueConstraint("theme_id", "kind", name="uq_theme_asset_kind"),)


@login_manager.user_loader
def load_user(user_id: str):
    return User.query.get(int(user_id))


def role_required(*roles: str):
    def decorator(func_):
        @wraps(func_)
        def wrapper(*args, **kwargs):
            if not current_user.is_authenticated:
                return redirect(url_for("login"))
            if current_user.role == "superadmin" or current_user.role in roles:
                return func_(*args, **kwargs)
            flash("คุณไม่มีสิทธิ์ทำรายการนี้", "warning")
            return redirect(url_for("index"))
        return wrapper
    return decorator


def next_manual_id(model):
    """Fallback สำหรับ PostgreSQL table เก่าที่ id ไม่มี sequence/default"""
    try:
        max_id = db.session.query(db.func.max(model.id)).scalar() or 0
        return int(max_id) + 1
    except Exception:
        db.session.rollback()
        return None

#---------------คะแนนเรียลไทม์------------
def event_has_round_of_16(event) -> bool:
    if event.next_round_label == "รอบ 16 คน":
        return True
    return BracketMatch.query.filter_by(event_id=event.id, round_name="R16").first() is not None


def scorecard_round_labels(event) -> list[str]:
    return SCORECARD_ROUND_LABELS_16 if event_has_round_of_16(event) else SCORECARD_ROUND_LABELS_8


def scorecard_round_numbers(event) -> list[int]:
    return list(range(1, len(scorecard_round_labels(event)) + 1))


def bracket_round_to_scorecard_round(round_name: str, event=None) -> int:
    if round_name == "R16":
        return 3
    if event is not None and event_has_round_of_16(event):
        return {"QF": 4, "SF": 5, "F": 6}[round_name]
    return {"QF": 3, "SF": 4, "F": 5}[round_name]

def build_bracket_row_data(event, athlete, round_name, seed_map=None):
    seed_map = seed_map or {}
    if not athlete:
        return {
            "athlete_id": None,
            "seed": "",
            "team": "",
            "name": "",
            "r1": "",
            "r1r2": "",
            "show_ref_score": False,
            "is_direct_qualifier": False,
            "stations": [0, 0, 0, 0, 0],
            "total": 0,
            "status": "waiting",
            "status_label": "รอคิว",
        }

    round_no = bracket_round_to_scorecard_round(round_name, event)

    # ใช้ group ปัจจุบันของระบบ: direct = เข้ารอบตรงจากรอบ 1
    progression_groups = get_progression_groups(event)
    is_direct_qualifier = athlete.id in progression_groups.get("direct", set())

    r1_summary = summarize_round(athlete.id, 1)
    r2_summary = summarize_round(athlete.id, 2) if event.has_round_two else {"total": 0}

    # S1-S5 และ SCORE ของ bracket ต้องดึงจาก scorecard รอบ bracket ปัจจุบันแบบ realtime
    current_summary = summarize_round(athlete.id, round_no) if round_no else {"by_station": {}, "total": 0}

    r1_total = r1_summary.get("total", 0)
    r2_total = r2_summary.get("total", 0) or 0

    # REF = รอบ 1 + รอบ 2
    # ซ่อน REF เฉพาะผู้เข้ารอบตรงจากรอบ 1 เท่านั้น
    ref_total = r1_total + r2_total
    show_ref_score = not is_direct_qualifier

    stations = []
    by_station = current_summary.get("by_station", {})
    for station in [1, 2, 3, 4, 5]:
        station_data = by_station.get(station, {"total": 0})
        stations.append(station_data.get("total", 0))

    round_status = athlete_round_status(athlete, round_no)

    return {
        "athlete_id": athlete.id,
        "seed": seed_map.get(athlete.id, ""),
        "team": athlete.affiliation or "",
        "name": athlete.name or "",
        "r1": r1_total,
        "r1r2": ref_total if show_ref_score else "",
        "show_ref_score": show_ref_score,
        "is_direct_qualifier": is_direct_qualifier,
        "stations": stations,
        "total": current_summary.get("total", 0),
        "status": round_status,
        "status_label": {"waiting": "รอคิว", "active": "กำลังตี", "finished": "ตีเสร็จแล้ว"}.get(round_status, "รอคิว"),
    }

def ensure_schema() -> None:
    os.makedirs(os.path.join(BASE_DIR, "instance"), exist_ok=True)

    # PostgreSQL บน Railway: ตารางที่เคยย้ายจาก SQLite บางตัวมี id NOT NULL
    # แต่ไม่มี DEFAULT nextval(sequence) ทำให้ INSERT แล้ว id = null
    # แก้แบบถาวรด้วยการสร้าง sequence และผูก default ให้ทุกตารางหลัก
    if db.engine.dialect.name == "postgresql":
        tables = [
            "user",
            "event",
            "athlete",
            "score_entry",
            "score_signature",
            "score_edit_log",
            "tie_break_entry",
            "bracket_match",
            "results_approved_setting",
        ]
        with db.engine.begin() as conn:
            for table in tables:
                seq = f"{table}_id_seq"
                conn.exec_driver_sql(f'CREATE SEQUENCE IF NOT EXISTS "{seq}"')
                conn.exec_driver_sql(
                    f"""SELECT setval(
                        '"{seq}"',
                        COALESCE((SELECT MAX(id) FROM "{table}"), 0) + 1,
                        false
                    )"""
                )
                conn.exec_driver_sql(
                    f'ALTER TABLE "{table}" ALTER COLUMN id SET DEFAULT nextval(\'"{seq}"\')'
                )
                conn.exec_driver_sql(
                    f'ALTER SEQUENCE "{seq}" OWNED BY "{table}".id'
                )
            # เพิ่มคอลัมน์สำหรับอัปโหลดโลโก้ Results Approved ในฐานข้อมูลเดิม
            ra_logo_columns = {
                "cover_main_logo_path": "VARCHAR(255)",
                "cover_bottom_logo_1_path": "VARCHAR(255)",
                "cover_bottom_logo_2_path": "VARCHAR(255)",
                "cover_bottom_logo_3_path": "VARCHAR(255)",
                "header_logo_1_path": "VARCHAR(255)",
                "header_logo_2_path": "VARCHAR(255)",
                "header_logo_3_path": "VARCHAR(255)",
                "header_logo_4_path": "VARCHAR(255)",
                "side_logo_path": "VARCHAR(255)",
                "name_format": "VARCHAR(10)",
            }
            for col, col_type in ra_logo_columns.items():
                conn.exec_driver_sql(f'ALTER TABLE "results_approved_setting" ADD COLUMN IF NOT EXISTS {col} {col_type}')
            conn.exec_driver_sql("UPDATE \"results_approved_setting\" SET name_format = 'single' WHERE name_format IS NULL")

            # ScoreEntry: เพิ่ม checkbox ตีแล้วสำหรับทุกช่องคะแนน
            conn.exec_driver_sql('ALTER TABLE "score_entry" ADD COLUMN IF NOT EXISTS is_scored BOOLEAN DEFAULT false')
            conn.exec_driver_sql('UPDATE "score_entry" SET is_scored = true WHERE COALESCE(score, 0) <> 0 OR COALESCE(is_red_card, false) = true')
            conn.exec_driver_sql('ALTER TABLE "score_signature" ADD COLUMN IF NOT EXISTS stopped_by_red BOOLEAN DEFAULT false')
            conn.exec_driver_sql('ALTER TABLE "user" ADD COLUMN IF NOT EXISTS court_no INTEGER')
            conn.exec_driver_sql('ALTER TABLE "user" ADD COLUMN IF NOT EXISTS court_event_id INTEGER')
            conn.exec_driver_sql("ALTER TABLE \"event\" ADD COLUMN IF NOT EXISTS round_two_mode VARCHAR(20) DEFAULT 'cutoff'")
            conn.exec_driver_sql("UPDATE \"event\" SET round_two_mode = 'cutoff' WHERE round_two_mode IS NULL")
            conn.exec_driver_sql('ALTER TABLE "athlete" ADD COLUMN IF NOT EXISTS round_two_disabled BOOLEAN DEFAULT false')
            conn.exec_driver_sql('ALTER TABLE "athlete" ADD COLUMN IF NOT EXISTS round_two_disabled_at TIMESTAMP')
            conn.exec_driver_sql('ALTER TABLE "athlete" ADD COLUMN IF NOT EXISTS round_two_disabled_by INTEGER')
        return

    # SQLite migration เดิม
    if db.engine.dialect.name != "sqlite":
        return

    with db.engine.begin() as conn:
        columns = {row[1] for row in conn.exec_driver_sql("PRAGMA table_info(score_signature)").fetchall()}
        if "recorder_signature" not in columns:
            conn.exec_driver_sql("ALTER TABLE score_signature ADD COLUMN recorder_signature TEXT")
        if "referee_signature" not in columns:
            conn.exec_driver_sql("ALTER TABLE score_signature ADD COLUMN referee_signature TEXT")
        if "athlete_signature" not in columns:
            conn.exec_driver_sql("ALTER TABLE score_signature ADD COLUMN athlete_signature TEXT")
        if "bypass_signed" not in columns:
            conn.exec_driver_sql("ALTER TABLE score_signature ADD COLUMN bypass_signed BOOLEAN DEFAULT 0")
        if "started_at" not in columns:
            conn.exec_driver_sql("ALTER TABLE score_signature ADD COLUMN started_at DATETIME")
        if "finished_at" not in columns:
            conn.exec_driver_sql("ALTER TABLE score_signature ADD COLUMN finished_at DATETIME")
        if "stopped_by_red" not in columns:
            conn.exec_driver_sql("ALTER TABLE score_signature ADD COLUMN stopped_by_red BOOLEAN DEFAULT 0")

        score_entry_columns = {row[1] for row in conn.exec_driver_sql("PRAGMA table_info(score_entry)").fetchall()}
        if "is_scored" not in score_entry_columns:
            conn.exec_driver_sql("ALTER TABLE score_entry ADD COLUMN is_scored BOOLEAN DEFAULT 0")
            conn.exec_driver_sql("UPDATE score_entry SET is_scored = 1 WHERE COALESCE(score, 0) <> 0 OR COALESCE(is_red_card, 0) = 1")

        event_columns = {row[1] for row in conn.exec_driver_sql("PRAGMA table_info(event)").fetchall()}
        if "has_round_two" not in event_columns:
            conn.exec_driver_sql("ALTER TABLE event ADD COLUMN has_round_two BOOLEAN DEFAULT 1")
        if "bracket_qualifiers" not in event_columns:
            conn.exec_driver_sql("ALTER TABLE event ADD COLUMN bracket_qualifiers INTEGER DEFAULT 8")
        if "direct_qualifiers" not in event_columns:
            conn.exec_driver_sql("ALTER TABLE event ADD COLUMN direct_qualifiers INTEGER DEFAULT 4")
        if "category" not in event_columns:
            conn.exec_driver_sql("ALTER TABLE event ADD COLUMN category VARCHAR(20) DEFAULT 'men'")
        if "next_round_label" not in event_columns:
            conn.exec_driver_sql("ALTER TABLE event ADD COLUMN next_round_label VARCHAR(50)")
        if "round_two_mode" not in event_columns:
            conn.exec_driver_sql("ALTER TABLE event ADD COLUMN round_two_mode VARCHAR(20) DEFAULT 'cutoff'")
            conn.exec_driver_sql("UPDATE event SET round_two_mode = 'cutoff' WHERE round_two_mode IS NULL")

        user_columns = {row[1] for row in conn.exec_driver_sql("PRAGMA table_info(user)").fetchall()}
        if "court_no" not in user_columns:
            conn.exec_driver_sql("ALTER TABLE user ADD COLUMN court_no INTEGER")
        if "court_event_id" not in user_columns:
            conn.exec_driver_sql("ALTER TABLE user ADD COLUMN court_event_id INTEGER")

        athlete_columns = {row[1] for row in conn.exec_driver_sql("PRAGMA table_info(athlete)").fetchall()}
        if "lane_no" not in athlete_columns:
            conn.exec_driver_sql("ALTER TABLE athlete ADD COLUMN lane_no INTEGER")
        if "lane_order" not in athlete_columns:
            conn.exec_driver_sql("ALTER TABLE athlete ADD COLUMN lane_order INTEGER")
        if "start_order" not in athlete_columns:
            conn.exec_driver_sql("ALTER TABLE athlete ADD COLUMN start_order INTEGER")
        if "round_two_disabled" not in athlete_columns:
            conn.exec_driver_sql("ALTER TABLE athlete ADD COLUMN round_two_disabled BOOLEAN DEFAULT 0")
        if "round_two_disabled_at" not in athlete_columns:
            conn.exec_driver_sql("ALTER TABLE athlete ADD COLUMN round_two_disabled_at DATETIME")
        if "round_two_disabled_by" not in athlete_columns:
            conn.exec_driver_sql("ALTER TABLE athlete ADD COLUMN round_two_disabled_by INTEGER")

        # Results Approved: คอลัมน์เก็บ path โลโก้ที่อัปโหลดเอง
        ra_columns = {row[1] for row in conn.exec_driver_sql("PRAGMA table_info(results_approved_setting)").fetchall()}
        ra_logo_columns = {
            "cover_main_logo_path": "VARCHAR(255)",
            "cover_bottom_logo_1_path": "VARCHAR(255)",
            "cover_bottom_logo_2_path": "VARCHAR(255)",
            "cover_bottom_logo_3_path": "VARCHAR(255)",
            "header_logo_1_path": "VARCHAR(255)",
            "header_logo_2_path": "VARCHAR(255)",
            "header_logo_3_path": "VARCHAR(255)",
            "header_logo_4_path": "VARCHAR(255)",
            "side_logo_path": "VARCHAR(255)",
            "name_format": "VARCHAR(10) DEFAULT 'single'",
        }
        for col, col_type in ra_logo_columns.items():
            if col not in ra_columns:
                conn.exec_driver_sql(f"ALTER TABLE results_approved_setting ADD COLUMN {col} {col_type}")
        conn.exec_driver_sql("UPDATE results_approved_setting SET name_format = 'single' WHERE name_format IS NULL")


def event_theme(category: str | None) -> str:
    """คืนค่า class สีของ event ตามประเภทการแข่งขัน"""
    category = (category or "men").strip().lower()
    if category in {"women", "female", "หญิง", "lady", "ladies"}:
        return "theme-women"
    if category in {"mixed", "mix", "ผสม", "คู่ผสม"}:
        return "theme-mixed"
    if category in {"youth", "junior", "เยาวชน"}:
        return "theme-youth"
    return "theme-men"


# รหัสผ่านที่เคยฝังในโค้ดเวอร์ชันก่อน ใครอ่านโค้ดก็รู้ จึงห้ามใช้เข้าระบบอีก
KNOWN_DEFAULT_PASSWORDS = {"yagami1225", "admin1234", "viewer1234"}
MIN_PASSWORD_LENGTH = 8


def password_problem(password: str) -> str | None:
    if len(password or "") < MIN_PASSWORD_LENGTH:
        return f"รหัสผ่านต้องยาวอย่างน้อย {MIN_PASSWORD_LENGTH} ตัวอักษร"
    if password in KNOWN_DEFAULT_PASSWORDS:
        return "ห้ามใช้รหัสผ่านตั้งต้นของระบบเดิม"
    return None


def seed_defaults() -> None:
    """สร้าง superadmin เฉพาะตอนฐานข้อมูลยังไม่มี superadmin เลย
    รหัสมาจาก SUPERADMIN_PASSWORD ถ้าไม่ตั้งจะสุ่มและพิมพ์ลง log ครั้งเดียว
    ไม่สร้างบัญชี admin/viewer รหัสตายตัวอีกแล้ว"""
    import secrets
    reset_pw = os.environ.get("RESET_SUPERADMIN_PASSWORD", "").strip()
    if not User.query.filter_by(role="superadmin").first():
        password = os.environ.get("SUPERADMIN_PASSWORD", "").strip() or secrets.token_urlsafe(12)
        u = User.query.filter_by(username="superadmin").first() or User(username="superadmin", role="superadmin")
        u.role = "superadmin"
        u.set_password(password)
        db.session.add(u)
        if not os.environ.get("SUPERADMIN_PASSWORD"):
            print(f"[SECURITY] สร้างบัญชี superadmin ใหม่ รหัสผ่านชั่วคราว: {password}  (เข้าระบบแล้วเปลี่ยนที่หน้าผู้ใช้)")
    elif reset_pw:
        # ทางออกฉุกเฉินเมื่อลืมรหัส superadmin: ตั้ง env นี้ รีสตาร์ต 1 ครั้ง แล้วลบ env ทิ้ง
        problem = password_problem(reset_pw)
        u = User.query.filter_by(username="superadmin", role="superadmin").first()
        if u and not problem:
            u.set_password(reset_pw)
            print("[SECURITY] ตั้งรหัส superadmin ใหม่จาก RESET_SUPERADMIN_PASSWORD แล้ว ให้ลบ env นี้ออกทันที")
        elif problem:
            print(f"[SECURITY] ไม่ได้ตั้งรหัส superadmin: {problem}")
    ensure_default_themes()
    db.session.commit()



def _request_cache() -> dict:
    """Request-local cache for read-heavy overview and bracket pages.

    The live polling cadence stays unchanged. This only prevents the same SQL
    queries from being repeated many times while building one response.
    """
    if not has_request_context():
        return {}
    cache = getattr(g, "shooting_request_cache", None)
    if cache is None:
        cache = {}
        g.shooting_request_cache = cache
    return cache


def clear_request_cache() -> None:
    if has_request_context():
        g.shooting_request_cache = {}


def preload_event_score_data(event: Event) -> None:
    """Load entries, signatures and tie-break scores for an event in 3 queries."""
    if not has_request_context():
        return
    cache = _request_cache()
    if ("preloaded_event", event.id) in cache:
        return
    athlete_ids = [a.id for a in event.athletes]
    cache[("preloaded_event", event.id)] = True
    if not athlete_ids:
        return

    entries_by_key = cache.setdefault("entries_by_athlete_round", {})
    for entry in ScoreEntry.query.filter(ScoreEntry.athlete_id.in_(athlete_ids)).all():
        entries_by_key.setdefault((entry.athlete_id, entry.round_no), []).append(entry)

    signatures_by_key = cache.setdefault("signature_by_athlete_round", {})
    for signature in ScoreSignature.query.filter(ScoreSignature.athlete_id.in_(athlete_ids)).all():
        signatures_by_key[(signature.athlete_id, signature.round_no)] = signature

    tiebreak_by_key = cache.setdefault("tiebreak_by_athlete_round", {})
    for entry in TieBreakEntry.query.filter(TieBreakEntry.athlete_id.in_(athlete_ids)).all():
        tiebreak_by_key.setdefault((entry.athlete_id, entry.round_no), []).append(entry)

def get_round_score_map(athlete_id: int, round_no: int) -> Dict[Tuple[int, int], ScoreEntry]:
    entries = ScoreEntry.query.filter_by(athlete_id=athlete_id, round_no=round_no).all()
    return {(e.station_no, e.distance_m): e for e in entries}


def get_round_signature(athlete_id: int, round_no: int) -> ScoreSignature | None:
    cache = _request_cache()
    signatures = cache.get("signature_by_athlete_round") if has_request_context() else None
    key = (athlete_id, round_no)
    if signatures is not None and key in signatures:
        return signatures[key]
    signature = ScoreSignature.query.filter_by(athlete_id=athlete_id, round_no=round_no).first()
    if has_request_context():
        cache.setdefault("signature_by_athlete_round", {})[key] = signature
    return signature


def ensure_round_entries(athlete_id: int, round_no: int) -> None:
    existing = {(e.station_no, e.distance_m) for e in ScoreEntry.query.filter_by(athlete_id=athlete_id, round_no=round_no).all()}
    changed = False
    for station_no in STATIONS:
        for distance_m in DISTANCES:
            key = (station_no, distance_m)
            if key not in existing:
                db.session.add(ScoreEntry(
                    athlete_id=athlete_id,
                    round_no=round_no,
                    station_no=station_no,
                    distance_m=distance_m,
                    score=0,
                    is_red_card=False,
                    is_scored=False,
                ))
                changed = True
    if changed:
        db.session.commit()
        clear_request_cache()


def ensure_signature(athlete_id: int, round_no: int) -> ScoreSignature:
    signature = get_round_signature(athlete_id, round_no)
    if signature:
        return signature
    signature = ScoreSignature(athlete_id=athlete_id, round_no=round_no)
    db.session.add(signature)
    db.session.commit()
    clear_request_cache()
    return signature


def summarize_round(athlete_id: int, round_no: int) -> dict:
    cache = _request_cache()
    summary_cache = cache.setdefault("round_summary", {}) if has_request_context() else {}
    key = (athlete_id, round_no)
    if key in summary_cache:
        return summary_cache[key]

    entries_map = cache.get("entries_by_athlete_round") if has_request_context() else None
    if entries_map is not None:
        entries = entries_map.get(key, [])
    else:
        entries = ScoreEntry.query.filter_by(athlete_id=athlete_id, round_no=round_no).all()

    total = sum(e.score for e in entries)
    count_5 = sum(1 for e in entries if e.score == 5)
    count_3 = sum(1 for e in entries if e.score == 3)
    red_cards = sum(1 for e in entries if e.is_red_card)
    by_station = {}
    for station_no in STATIONS:
        station_entries = [e for e in entries if e.station_no == station_no]
        played_distances = {e.distance_m: bool(getattr(e, "is_scored", False)) for e in station_entries}
        by_station[station_no] = {
            "distances": {e.distance_m: e.score for e in station_entries},
            "played_distances": played_distances,
            "played_count": sum(1 for v in played_distances.values() if v),
            "is_complete": all(played_distances.get(distance_m, False) for distance_m in DISTANCES),
            "total": sum(e.score for e in station_entries),
        }

    tiebreak_map = cache.get("tiebreak_by_athlete_round") if has_request_context() else None
    if tiebreak_map is not None:
        tiebreak_entries = tiebreak_map.get(key, [])
    else:
        tiebreak_entries = TieBreakEntry.query.filter_by(athlete_id=athlete_id, round_no=round_no).all()
    result = {
        "total": total,
        "count_5": count_5,
        "count_3": count_3,
        "red_cards": red_cards,
        "by_station": by_station,
        "tiebreak_total": sum(e.score for e in tiebreak_entries),
        "tiebreak_count": len(tiebreak_entries),
    }
    summary_cache[key] = result
    return result


def build_scorecard_template_data(athlete_id: int) -> dict:
    entries = ScoreEntry.query.filter_by(athlete_id=athlete_id).all()

    score_map: dict[tuple[int, int, int], dict] = {}
    station_totals: dict[tuple[int, int], int] = {}
    station_reds: dict[tuple[int, int], int] = {}
    round_totals: dict[int, int] = {}

    for round_no in ALL_SCORECARD_ROUNDS:
        round_totals[round_no] = 0
        for station_no in STATIONS:
            station_entries = [
                e for e in entries
                if e.round_no == round_no and e.station_no == station_no
            ]
            station_total = sum(e.score for e in station_entries)
            station_red = sum(1 for e in station_entries if e.is_red_card)

            station_totals[(round_no, station_no)] = station_total
            station_reds[(round_no, station_no)] = station_red
            round_totals[round_no] += station_total

            for distance_m in DISTANCES:
                entry = next(
                    (e for e in station_entries if e.distance_m == distance_m),
                    None
                )
                score_map[(round_no, station_no, distance_m)] = {
                    "score": entry.score if entry else "",
                    "is_red_card": entry.is_red_card if entry else False,
                    "is_scored": bool(getattr(entry, "is_scored", False)) if entry else False,
                }

    round_ranks = {1: "", 2: "", 3: "", 4: "", 5: ""}
    return {
        "score_map": score_map,
        "station_totals": station_totals,
        "station_reds": station_reds,
        "round_totals": round_totals,
        "round_ranks": round_ranks,
    }


def athlete_round_status(athlete: Athlete, round_no: int) -> str:
    signature = get_round_signature(athlete.id, round_no)
    if signature and signature.finished_at:
        return "finished"
    if signature and signature.started_at:
        return "active"
    return "waiting"


def athlete_round_is_approved(athlete: Athlete, round_no: int) -> bool:
    """Formal approval is separate from the live shooting/cursor state."""
    signature = get_round_signature(athlete.id, round_no)
    if not signature:
        return False
    if bool(getattr(signature, "bypass_signed", False)):
        return True
    return all([
        bool(getattr(signature, "recorder_signature", None)),
        bool(getattr(signature, "referee_signature", None)),
        bool(getattr(signature, "athlete_signature", None)),
    ])


def overview_score_complete(summary: dict) -> bool:
    """All 5 stations have recorded all 4 distances."""
    by_station = summary.get("by_station") or {}
    return bool(by_station) and all(
        bool((by_station.get(station) or {}).get("is_complete", False))
        for station in STATIONS
    )


def bracket_match_status(event: Event, match: BracketMatch) -> dict:
    """สถานะของคู่ใน bracket จากลายเซ็น/เวลาเริ่มของนักกีฬาทั้ง 2 ฝั่ง"""
    round_no = bracket_round_to_scorecard_round(match.round_name, event)
    athlete_ids = [aid for aid in [match.athlete_a_id, match.athlete_b_id] if aid]
    if not athlete_ids or len(athlete_ids) < 2:
        return {"key": "waiting", "label": "รอคู่แข่งขัน"}
    if match.winner_id:
        return {"key": "finished", "label": "แข่งเสร็จ"}
    statuses = []
    for aid in athlete_ids:
        sig = get_round_signature(aid, round_no)
        if sig and sig.finished_at:
            statuses.append("finished")
        elif sig and sig.started_at:
            statuses.append("active")
        else:
            statuses.append("waiting")
    if any(status == "active" for status in statuses):
        return {"key": "active", "label": "กำลังแข่งขัน"}
    if all(status == "finished" for status in statuses):
        return {"key": "ready_result", "label": "รอบันทึกผู้ชนะ"}
    if any(status == "finished" for status in statuses):
        return {"key": "active", "label": "กำลังแข่งขัน"}
    return {"key": "waiting", "label": "รอแข่งขัน"}


def ranking_key(item: dict):
    """
    สูตรจัดอันดับ Shooting ที่ใช้ในระบบ
    1) คะแนนรวม / TOTAL มากกว่า
    2) จำนวนคะแนน 5 มากกว่า
    3) จำนวนคะแนน 3 มากกว่า
    4) คะแนน Shoot-off 7 เมตรทุกสถานี (เมื่อมีการบันทึก)

    หมายเหตุ: ตัดคะแนนสถานี 5 และสถานี 4 ออกจากสูตรแล้ว
    เพื่อไม่ให้สถานีใดมีน้ำหนักมากกว่าสถานีอื่น
    """
    return (
        item["total"],
        item["count_5"],
        item["count_3"],
        item["tiebreak_total"],
    )


def apply_round1_display_ranking(rows: list[dict], direct_limit: int) -> None:
    """จัด Class รอบ 1 ตามกติกาที่ใช้ในงานนี้

    - ช่วงอันดับเข้ารอบตรง 1..direct_limit ใช้ TOTAL -> 5 -> 3 -> Shoot-off
      เพื่อให้ seed ที่ผ่านตรงแยกกันชัดเจน
    - หลังพ้นโควตาเข้ารอบตรงแล้ว ใช้ TOTAL เป็นตัวกำหนด Class เท่านั้น
      ถ้า TOTAL เท่ากัน ให้ Class เท่ากันแบบ competition ranking เช่น 5,5,5,8
    - ordinal_rank/view_order ยังเก็บตำแหน่งจริงของแถวไว้เพื่อใช้กับเส้นตัดและการเรียง
    """
    if not rows:
        return

    direct_limit = max(int(direct_limit or 0), 0)
    previous_direct_key = None
    previous_direct_rank = None
    previous_total = None
    previous_total_rank = None

    for idx, row in enumerate(rows, start=1):
        row["ordinal_rank"] = idx
        row["view_order"] = idx

        if idx <= direct_limit:
            # ใน 1-4 (หรือจำนวน direct ที่ตั้งไว้) ต้องแยกด้วย 5/3/Shoot-off
            direct_key = ranking_key(row)
            if previous_direct_key is not None and direct_key == previous_direct_key:
                rank = previous_direct_rank
            else:
                rank = idx
            previous_direct_key = direct_key
            previous_direct_rank = rank
        else:
            # หลังโควตา direct: คะแนนรวมเท่ากัน = Class เท่ากันทันที
            total = row.get("total", 0)
            if previous_total is not None and total == previous_total:
                rank = previous_total_rank
            else:
                rank = idx
            previous_total = total
            previous_total_rank = rank

        row["rank"] = rank
        row["display_rank"] = rank


def apply_round2_display_ranking(rows: list[dict], start_rank: int, advancer_limit: int) -> None:
    """จัด Class ของผู้เล่นรอบ 2

    ผู้ที่จะผ่านจากรอบ 2 จำนวน advancer_limit คน ต้องแยกด้วย
    SUM -> จำนวน 5 รวม -> จำนวน 3 รวม -> Shoot-off
    หลังพ้นโควตานี้แล้ว ถ้า SUM เท่ากันให้ Class เท่ากันแบบ competition ranking
    """
    if not rows:
        return

    start_rank = int(start_rank or 1)
    advancer_limit = max(int(advancer_limit or 0), 0)
    previous_seed_key = None
    previous_seed_rank = None
    previous_sum = None
    previous_sum_rank = None

    for offset, row in enumerate(rows):
        position_rank = start_rank + offset
        row["ordinal_rank"] = position_rank
        row["view_order"] = position_rank

        if offset < advancer_limit:
            seed_key = (
                row.get("combined_total", row.get("total", 0)),
                row.get("count_5", 0),
                row.get("count_3", 0),
                row.get("tiebreak_total", 0),
            )
            if previous_seed_key is not None and seed_key == previous_seed_key:
                rank = previous_seed_rank
            else:
                rank = position_rank
            previous_seed_key = seed_key
            previous_seed_rank = rank
        else:
            combined_total = row.get("combined_total", row.get("total", 0))
            if previous_sum is not None and combined_total == previous_sum:
                rank = previous_sum_rank
            else:
                rank = position_rank
            previous_sum = combined_total
            previous_sum_rank = rank

        row["rank"] = rank
        row["display_rank"] = rank




def bracket_size(event: Event) -> int:
    """จำนวนคนรอบ Knockout จาก label เช่น รอบ 16/8/4"""
    label = (event.next_round_label or "")
    if "16" in label:
        return 16
    if "4" in label or "รอง" in label:
        return 4
    return 8


def direct_quota(event: Event) -> int:
    """จำนวนคนที่เข้ารอบตรงจากรอบแรก

    ต้องยึดค่าที่ผู้ใช้กรอกในอีเวนต์ก่อน เช่น
    - รอบ 16 กรอกผ่านรอบแรก 8 = ต้องเป็น 8 จริง
    - รอบ 8 กรอกผ่านรอบแรก 4 = ต้องเป็น 4 จริง

    ถ้าอีเวนต์เก่าไม่ได้กรอก direct_qualifiers จึงค่อย fallback เป็นครึ่งหนึ่งของ bracket
    """
    size = bracket_size(event)
    configured = int(event.direct_qualifiers or 0)
    if configured > 0:
        return min(configured, size)
    if event.has_round_two:
        return max(size // 2, 0)
    return size


def round2_advancer_quota(event: Event) -> int:
    """จำนวนคนที่ผ่านจากรอบ 2 เข้า Knockout

    ต้องยึดค่าที่ผู้ใช้กรอกจริง เช่น รอบ 16 ถ้ากรอกให้ผ่านจากรอบ 2 = 8
    ก็ต้องใช้ 8 ไม่ใช่คำนวณเหลือเองจนเพี้ยน
    """
    if not event.has_round_two:
        return 0
    size = bracket_size(event)
    configured = int(event.round_two_advancers or 0)
    remaining = max(size - direct_quota(event), 0)
    if configured > 0:
        return min(configured, size)
    return remaining


def round2_cutoff_rank(event: Event) -> int:
    """อันดับสูงสุด (Class) ที่มีสิทธิ์ตีรอบ 2 จากรอบ 1.

    ตัวอย่าง: ตั้งค่า 16 = ทุกคนที่ Class <= 16 และไม่ได้ผ่านตรง
    มีสิทธิ์ตีรอบ 2 ทั้งหมด หากมีหลายคน Class 16 ก็ได้สิทธิ์ทุกคน.

    มี 2 แบบ (event.round_two_mode):
    - "cutoff": ถึง Class ที่ N เช่น N=16 → Class 5–16 เมื่อผ่านตรง 4 คน
    - "next":   ต่อจากผู้ผ่านตรงอีก N ลำดับ เช่น ผ่านตรง 4, N=16 → Class 5–20
    ทั้งสองแบบ ถ้ามีหลายคน Class เท่ากันตรงเส้นตัด ได้สิทธิ์ทุกคน
    """
    if not event.has_round_two:
        return 0
    n = max(int(event.round_two_cutoff_rank or 0), 0)
    if n and round_two_mode(event) == "next":
        return direct_quota(event) + n
    return n


ROUND_TWO_MODES = {"cutoff": "ถึงลำดับ (Class) ที่ N", "next": "ต่อจากผู้ผ่านตรงอีก N ลำดับ"}


def round_two_mode(event) -> str:
    mode = getattr(event, "round_two_mode", None) or "cutoff"
    return mode if mode in ROUND_TWO_MODES else "cutoff"


def parse_round_two_mode(form) -> str:
    mode = (form.get("round_two_mode") or "cutoff").strip()
    return mode if mode in ROUND_TWO_MODES else "cutoff"


def round1_round2_candidate_ids(
    rows: list[dict],
    direct_ids: set[int],
    direct_pending_ids: set[int],
    cutoff_rank: int,
) -> set[int]:
    """เลือกผู้มีสิทธิ์รอบ 2 โดยยึด Class ถึงอันดับที่กำหนด.

    กติกา:
    - ผู้ผ่านตรงไม่ต้องตีรอบ 2
    - กลุ่มที่ยังรอ Shoot-off เพื่อแยกผู้ผ่านตรงยังไม่ถูกฟันธง
    - ผู้เล่นที่เหลือซึ่งมี Class <= cutoff_rank ได้ตีรอบ 2 ทุกคน
    - ถ้ามีหลายคน Class เท่ากับ cutoff_rank (เช่น Class 16 หลายคน)
      ได้สิทธิ์ทั้งหมด ไม่มีการจำกัดจำนวนคน
    """
    cutoff_rank = max(int(cutoff_rank or 0), 0)
    if cutoff_rank <= 0:
        return set()

    return {
        row["athlete"].id
        for row in rows
        if row["athlete"].id not in direct_ids
        and row["athlete"].id not in direct_pending_ids
        and int(row.get("display_rank", row.get("rank", 999999)) or 999999) <= cutoff_rank
    }


def apply_sequential_rank(rows: list[dict], start: int = 1) -> None:
    """ใช้กับ Overview รอบ 2/Seed: ต้องไม่มีลำดับซ้ำ"""
    for idx, row in enumerate(rows, start=start):
        row["rank"] = idx
        row["ordinal_rank"] = idx
        row["display_rank"] = idx
        row["view_order"] = idx

def build_round_ranking(event: Event, round_no: int) -> List[dict]:
    cache = _request_cache()
    cache_key = ("round_ranking", event.id, round_no)
    if has_request_context() and cache_key in cache:
        return cache[cache_key]

    preload_event_score_data(event)
    athletes = sorted(event.athletes, key=lambda a: a.start_order)
    rows = []
    round2_display_map = {}
    if round_no == 2:
        round1_rows = build_round_ranking(event, 1)
        direct_ids_for_r2 = exact_cut_ids(round1_rows, direct_quota(event))
        direct_pending_ids_for_r2 = unresolved_tie_ids(round1_rows, direct_quota(event))
        round2_source_rows = [
            row for row in round1_rows
            if row["athlete"].id not in direct_ids_for_r2
            and row["athlete"].id not in direct_pending_ids_for_r2
            and is_round_two_candidate(event, row["athlete"])
        ]
        round2_source_rows.sort(key=lambda row: (
            row["total"], row["count_5"], row["count_3"], row["tiebreak_total"], row["athlete"].start_order,
        ))
        lane_count = max(event.lane_count or 1, 1)
        for idx, source_row in enumerate(round2_source_rows, start=1):
            round2_display_map[source_row["athlete"].id] = {
                "display_order": idx,
                "display_lane_no": ((idx - 1) % lane_count) + 1,
                "display_lane_order": ((idx - 1) // lane_count) + 1,
            }

    for athlete in athletes:
        if round_no == 2 and not get_round_signature(athlete.id, 2):
            continue
        summary = summarize_round(athlete.id, round_no)
        logical_status = athlete_round_status(athlete, round_no)
        approved = athlete_round_is_approved(athlete, round_no)
        score_complete = overview_score_complete(summary) or logical_status == "finished"
        row = {
            "athlete": athlete,
            "round_no": round_no,
            "total": summary["total"],
            "count_5": summary["count_5"],
            "count_3": summary["count_3"],
            "tiebreak_total": summary["tiebreak_total"],
            "tiebreak_count": summary.get("tiebreak_count", 0),
            "status": logical_status,
            "score_complete": score_complete,
            "approved": approved,
            "by_station": summary["by_station"],
            "red_cards": summary["red_cards"],
            "display_order": athlete.start_order,
            "display_lane_no": athlete.lane_no,
            "display_lane_order": athlete.lane_order,
        }
        if round_no == 2 and athlete.id in round2_display_map:
            row.update(round2_display_map[athlete.id])
        rows.append(row)
    rows.sort(key=ranking_key, reverse=True)
    if round_no == 1:
        apply_round1_display_ranking(rows, direct_quota(event))
    else:
        # build_round_ranking(round 2) เป็นคะแนนรอบ 2 ดิบ ใช้ลำดับตำแหน่งภายในก่อน
        # หน้า Overview รอบ 2 จะจัด Class จาก SUM แยกต่างหาก
        apply_sequential_rank(rows)
    if has_request_context():
        cache[cache_key] = rows
    return rows


def round_two_candidate_ids(event: Event) -> set[int]:
    cache = _request_cache()
    cache_key = ("round_two_candidate_ids", event.id)
    if has_request_context() and cache_key in cache:
        return cache[cache_key]

    ids: set[int] = set()
    cutoff_rank = round2_cutoff_rank(event)
    # ห้ามฟันธงสิทธิ์รอบ 2 ระหว่างที่รอบ 1 ยังยิงไม่ครบทั้งอีเวนต์
    # Ranking ยังแสดงสดได้ แต่ candidate list ต้องรอผลรอบ 1 สุดท้ายก่อน
    if event.has_round_two and cutoff_rank > 0 and is_round_one_complete(event):
        round1_rows = build_round_ranking(event, 1)
        direct = direct_quota(event)

        # กติกา: ผู้ที่ไม่ผ่านตรงและมี Class ถึงอันดับ cutoff_rank
        # มีสิทธิ์ตีรอบ 2 ทุกคน รวมทุกคนที่ Class เท่ากับอันดับตัด
        direct_ids = exact_cut_ids(round1_rows, direct)
        direct_shoot_ids = unresolved_tie_ids(round1_rows, direct)
        ids = round1_round2_candidate_ids(
            round1_rows,
            direct_ids,
            direct_shoot_ids,
            cutoff_rank,
        )

        # Manual override: ผู้ดูแลสามารถปิดนักกีฬารายคนจากรอบ 2 ได้
        disabled_ids = {
            athlete.id for athlete in event.athletes
            if bool(getattr(athlete, "round_two_disabled", False))
        }
        ids -= disabled_ids

    if has_request_context():
        cache[cache_key] = ids
    return ids


def is_round_two_candidate(event: Event, athlete: Athlete) -> bool:
    return athlete.id in round_two_candidate_ids(event)


def sync_round_two_candidates(event: Event) -> None:
    if not event.has_round_two:
        return

    candidate_ids = round_two_candidate_ids(event)
    athlete_ids = [athlete.id for athlete in event.athletes]
    if not athlete_ids:
        return

    signatures = {
        sig.athlete_id: sig
        for sig in ScoreSignature.query.filter(
            ScoreSignature.athlete_id.in_(athlete_ids),
            ScoreSignature.round_no == 2,
        ).all()
    }
    existing_entry_keys = {
        (entry.athlete_id, entry.station_no, entry.distance_m)
        for entry in ScoreEntry.query.filter(
            ScoreEntry.athlete_id.in_(athlete_ids),
            ScoreEntry.round_no == 2,
        ).all()
    }

    changed = False
    for athlete in event.athletes:
        sig2 = signatures.get(athlete.id)
        if athlete.id in candidate_ids:
            for station_no in STATIONS:
                for distance_m in DISTANCES:
                    key = (athlete.id, station_no, distance_m)
                    if key not in existing_entry_keys:
                        db.session.add(ScoreEntry(
                            athlete_id=athlete.id,
                            round_no=2,
                            station_no=station_no,
                            distance_m=distance_m,
                            score=0,
                            is_red_card=False,
                        ))
                        changed = True
            if not sig2:
                sig2 = ScoreSignature(athlete_id=athlete.id, round_no=2)
                db.session.add(sig2)
                signatures[athlete.id] = sig2
                changed = True
            if not sig2.started_at and not sig2.finished_at and athlete.status != "waiting":
                athlete.status = "waiting"
                changed = True
        elif sig2 and not sig2.started_at and not sig2.finished_at:
            ScoreEntry.query.filter_by(athlete_id=athlete.id, round_no=2).delete(synchronize_session=False)
            TieBreakEntry.query.filter_by(athlete_id=athlete.id, round_no=2).delete(synchronize_session=False)
            ScoreSignature.query.filter_by(athlete_id=athlete.id, round_no=2).delete(synchronize_session=False)
            changed = True

    if changed:
        db.session.commit()
        clear_request_cache()


def build_round_two_start_list(event: Event) -> list[dict]:
    # ยังไม่สร้างคิวรอบ 2 จนกว่ารอบ 1 จะจบครบทั้งอีเวนต์
    if not is_round_one_complete(event):
        return []
    round1_rows = build_round_ranking(event, 1)

    direct_ids = exact_cut_ids(round1_rows, direct_quota(event))
    direct_pending_ids = unresolved_tie_ids(round1_rows, direct_quota(event))

    candidates = [
        r for r in round1_rows
        if is_round_two_candidate(event, r["athlete"])
        and r["athlete"].id not in direct_ids
        and r["athlete"].id not in direct_pending_ids
    ]

    # คะแนนรอบ 1 น้อยสุด ได้ตีก่อน
    candidates.sort(key=lambda r: (
        r["total"],
        r["count_5"],
        r["count_3"],
        r["tiebreak_total"],
        r["athlete"].start_order if r["athlete"].start_order is not None else 9999
    ))

    return candidates


def sort_round_two_rows_for_start(rows: list[dict]) -> list[dict]:
    """เรียงลำดับการตีรอบ 2 ตาม display_order: คะแนนรอบ 1 น้อยสุดตีก่อน"""
    return sorted(
        rows,
        key=lambda row: (
            row.get("display_order") if row.get("display_order") is not None else 999999,
            row["athlete"].start_order if row["athlete"].start_order is not None else 999999,
            row["athlete"].id,
        )
    )

def build_combined_qualifiers(event: Event) -> List[dict]:
    cache = _request_cache()
    cache_key = ("combined_qualifiers", event.id)
    if has_request_context() and cache_key in cache:
        return cache[cache_key]
    round1_rows = build_round_ranking(event, 1)
    if not is_round_one_complete(event):
        if has_request_context():
            cache[cache_key] = []
        return []
    direct_ids = exact_cut_ids(round1_rows, direct_quota(event))
    direct_pending_ids = unresolved_tie_ids(round1_rows, direct_quota(event))
    direct_rows = []
    for row in round1_rows:
        if row["athlete"].id in direct_ids:
            direct_rows.append({
                **row,
                "round1_total": row["total"],
                "round2_total": None,
                "round1_by_station": row.get("by_station", {}),
                "round2_by_station": None,
                "combined_total": row["total"],
            })
    for idx, row in enumerate(direct_rows, start=1):
        row["seed"] = idx
    if not event.has_round_two:
        return direct_rows

    candidate_ids = round_two_candidate_ids(event) - direct_ids - direct_pending_ids
    finished_candidate_ids = {
        aid for aid in candidate_ids
        if athlete_round_status(Athlete.query.get(aid), 2) == "finished"
    }
    all_round2_candidates_finished = bool(candidate_ids) and candidate_ids == finished_candidate_ids

    # รวมผลรอบ 2 เพื่อใช้คัดเข้า Bracket เฉพาะคนที่ตีจบแล้วเท่านั้น
    # กันไม่ให้คนที่ระบบสร้าง Signature ไว้แต่ยังไม่ได้ตี ถูกจัดผ่าน/ตกรอบก่อนเวลา
    round2_rows = [
        row for row in build_round_ranking(event, 2)
        if row["athlete"].id in finished_candidate_ids
    ]
    combined_rows = []
    for row in round2_rows:
        round1_row = next((r for r in round1_rows if r["athlete"].id == row["athlete"].id), None)
        round1_total = round1_row["total"] if round1_row else 0
        combined_rows.append({
            **row,
            "total": round1_total + row["total"],
            "combined_total": round1_total + row["total"],
            "count_5": (round1_row["count_5"] if round1_row else 0) + row["count_5"],
            "count_3": (round1_row["count_3"] if round1_row else 0) + row["count_3"],
            # เมื่อคัดจากรอบ 2 ใช้ SUM(R1+R2) -> 5รวม -> 3รวม -> Shoot-off รอบ 2
            # ไม่บวก Shoot-off รอบ 1 เข้ามา เพราะรอบ 1 เป็นคนละเหตุผล/คนละเที่ยว
            "tiebreak_total": row["tiebreak_total"],
            "tiebreak_count": row.get("tiebreak_count", 0),
            "round1_tiebreak_total": (round1_row.get("tiebreak_total", 0) if round1_row else 0),
            "round1_tiebreak_count": (round1_row.get("tiebreak_count", 0) if round1_row else 0),
            "round2_tiebreak_total": row["tiebreak_total"],
            "round2_tiebreak_count": row.get("tiebreak_count", 0),
            "round1_total": round1_total,
            "round2_total": row["total"],
            "round1_by_station": (round1_row.get("by_station", {}) if round1_row else {}),
            "round2_by_station": row.get("by_station", {}),
        })
    combined_rows.sort(key=ranking_key, reverse=True)
    # Ranking ของผู้เล่นรอบ 2 เริ่มต่อจากผู้ผ่านตรง และ 4 ที่นั่งแรกต้องแยก seed ให้ชัด
    apply_round2_display_ranking(
        combined_rows,
        start_rank=len(direct_rows) + 1,
        advancer_limit=round2_advancer_quota(event),
    )

    # คัดผู้ผ่านจากรอบ 2 ต้องเอาตามจำนวนที่ตั้งไว้จริง ๆ
    # ถ้าคะแนนเท่ากันที่เส้นตัด ให้ตัดสินตาม TOTAL -> จำนวน 5 -> จำนวน 3 -> Shoot-off
    # ห้ามขยาย bracket เองเพราะจะทำให้ตั้ง 8 คน แต่หลุดเป็น 9/16 คน
    advancer_limit = max(round2_advancer_quota(event), 0)

    def base_tie_key(row):
        return (row["combined_total"], row["count_5"], row["count_3"])

    if all_round2_candidates_finished:
        shoot_off_ids = unresolved_tie_ids(combined_rows, advancer_limit)
        passed_ids = exact_cut_ids(combined_rows, advancer_limit)
    else:
        shoot_off_ids = set()
        passed_ids = set()

    seed_no = len(direct_rows) + 1
    for row in combined_rows:
        aid = row["athlete"].id
        row["shoot_off_required"] = aid in shoot_off_ids
        if row["shoot_off_required"]:
            row["seed"] = None
            row["passed_cut"] = False
        elif aid in passed_ids:
            row["seed"] = seed_no
            row["passed_cut"] = True
            seed_no += 1
        else:
            row["seed"] = None
            row["passed_cut"] = False
    result = direct_rows + combined_rows
    if has_request_context():
        cache[cache_key] = result
    return result




def base_shootoff_key(row: dict) -> tuple:
    """เกณฑ์ก่อนรอบพิเศษ: TOTAL -> จำนวน 5 -> จำนวน 3

    ถ้า 3 ตัวนี้ยังเท่ากัน แปลว่า "ยังจัดลำดับจริงไม่ได้"
    ต้องยิง Shoot-off ยกเว้นกรณีเส้นสุดท้ายของสิทธิ์ไปตีรอบ 2 ซึ่งกติกาให้ไปตีได้ทั้งหมด
    """
    return (
        row.get("combined_total", row.get("total", 0)),
        row.get("count_5", 0),
        row.get("count_3", 0),
    )


def tiebreak_done_key(row: dict) -> tuple:
    """ใช้ดูว่ารอบพิเศษตัดสินได้หรือยัง

    ต้องบันทึกจำนวนเที่ยวเท่ากันทุกคนก่อน แล้วจึงเอาคะแนนรอบพิเศษมาแยกอันดับ
    """
    return (row.get("tiebreak_count", 0), row.get("tiebreak_total", 0))


def cutoff_shootoff_ids(rows: list[dict], cutoff_count: int) -> set[int]:
    """คืน athlete_id ที่ต้อง Shoot-off เฉพาะเส้นตัดจริง

    เงื่อนไข:
    1) ตัดตามจำนวนที่ตั้งไว้จริง เช่น 8 คนคือ 8 คน
    2) ถ้าอันดับสุดท้ายกับคนถัดไปไม่เท่ากันตาม TOTAL -> 5 -> 3 = ไม่ต้อง Shoot-off
    3) ถ้าเท่ากัน ให้ทั้งกลุ่มที่มี key เดียวกันต้อง Shoot-off พร้อมกัน
    4) ถ้าบันทึก Shoot-off ครบทุกคนแล้วและคะแนนเที่ยวพิเศษต่างกัน = ตัดสินได้ ไม่ขึ้น Shoot-off
    5) ถ้าบันทึกยังไม่ครบ หรือบันทึกครบแล้วยังเท่ากัน = ยังขึ้น Shoot-off เพื่อให้ตีต่อ
    """
    if not cutoff_count or cutoff_count <= 0 or len(rows) <= cutoff_count:
        return set()

    cut_row = rows[cutoff_count - 1]
    next_row = rows[cutoff_count]
    base_key = base_shootoff_key(cut_row)
    if base_key != base_shootoff_key(next_row):
        return set()

    group = [row for row in rows if base_shootoff_key(row) == base_key]
    if len(group) < 2:
        return set()

    # กลุ่มต้องคร่อมเส้นตัดเท่านั้น ไม่ใช่เสมอกันภายในกลุ่มที่ผ่านหมด/ตกรอบหมด
    group_ids = {row["athlete"].id for row in group}
    above_ids = {row["athlete"].id for row in rows[:cutoff_count]}
    below_ids = {row["athlete"].id for row in rows[cutoff_count:]}
    if not (group_ids & above_ids and group_ids & below_ids):
        return set()

    counts = [row.get("tiebreak_count", 0) for row in group]
    totals = [row.get("tiebreak_total", 0) for row in group]

    # ต้องตีพร้อมกันทุกคนในกลุ่ม ถ้ามีคนใดยังไม่ได้บันทึก หรือจำนวนเที่ยวไม่เท่ากัน ให้ยังขึ้นทั้งกลุ่ม
    if min(counts) == 0 or len(set(counts)) > 1:
        return group_ids

    # บันทึกครบเท่ากันแล้ว ถ้าคะแนน Shoot-off ยังเท่ากัน ให้ตีต่อ
    if len(set(totals)) == 1:
        return group_ids

    # คะแนน Shoot-off ต่างกันแล้ว ตัดสินได้
    return set()


def unresolved_tie_ids(rows: list[dict], scope_count: int | None = None) -> set[int]:
    """หากเสมอหลัง TOTAL -> 5 -> 3 แล้วต้องเรียงลำดับจริง ให้ส่งไป Shoot-off

    scope_count ระบุจำนวนอันดับที่มีผล เช่น direct quota หรือจำนวนที่ผ่านจากรอบ 2
    ถ้ากลุ่มเสมอแตะอยู่ใน scope จะถือว่ายังตัดสินอันดับไม่ได้
    สิทธิ์ตีรอบ 2 รอบแรกใช้หลัก next N at least และขยายตาม TOTAL ที่เส้นตัดแยกต่างหาก
    """
    if not rows:
        return set()
    scoped_ids = None
    if scope_count is not None:
        if scope_count <= 0:
            return set()
        scoped_ids = {row["athlete"].id for row in rows[:min(scope_count, len(rows))]}

    groups: dict[tuple, list[dict]] = {}
    for row in rows:
        groups.setdefault(base_shootoff_key(row), []).append(row)

    result: set[int] = set()
    for group in groups.values():
        if len(group) < 2:
            continue
        group_ids = {row["athlete"].id for row in group}
        if scoped_ids is not None and not (group_ids & scoped_ids):
            continue

        counts = [row.get("tiebreak_count", 0) for row in group]
        totals = [row.get("tiebreak_total", 0) for row in group]

        # ยังไม่ได้ตีครบทุกคน หรือจำนวนเที่ยวไม่เท่ากัน = ต้อง Shoot-off / ตีต่อ
        if min(counts) == 0 or len(set(counts)) > 1:
            result.update(group_ids)
            continue

        # ตีครบเท่ากันแล้วแต่คะแนนพิเศษยังเท่ากัน = ต้องตีต่อ
        if len(set(totals)) == 1:
            result.update(group_ids)

    return result




def overview_unresolved_shootoff_ids(rows: list[dict], scope_count: int | None = None, round_no: int | None = None) -> set[int]:
    """หาแถวที่ต้องยิง Shoot-off ในหน้า Overview

    กติกาที่ใช้:
    - เรียงด้วย TOTAL -> จำนวน 5 -> จำนวน 3 ทันที
    - ถ้า 3 ตัวนี้ยังเท่ากัน แปลว่ายังจัดลำดับจริงไม่ได้
    - ถ้ากลุ่มนั้นอยู่ในช่วงที่ต้องใช้จัดอันดับ/เข้า seed ให้ขึ้นปุ่ม Shoot-off
    - ต้องกดจบการตีของทุกคนในกลุ่มก่อน จึงขึ้น Shoot-off
    - ถ้าบันทึก Shoot-off ไม่ครบทุกคน หรือจำนวนเที่ยวไม่เท่ากัน ยังถือว่ายังตัดสินไม่ได้
    """
    if not rows:
        return set()

    scoped_ids = None
    if scope_count is not None:
        if scope_count <= 0:
            return set()
        # กลุ่มที่ชนหรืออยู่ในช่วง scope ต้องถูกตรวจทั้งกลุ่ม ไม่ใช่แค่คนใน slice
        scoped_ids = {row["athlete"].id for row in rows[:min(scope_count, len(rows))]}

    groups: dict[tuple, list[dict]] = {}
    for row in rows:
        groups.setdefault(base_shootoff_key(row), []).append(row)

    result: set[int] = set()
    for group in groups.values():
        if len(group) < 2:
            continue
        group_ids = {row["athlete"].id for row in group}
        if scoped_ids is not None and not (group_ids & scoped_ids):
            continue
        if round_no is not None:
            # ขึ้นเมื่อคนในกลุ่มนี้ตีจบครบ ไม่ต้องรอทั้งตาราง
            if not all(athlete_round_status(row["athlete"], round_no) == "finished" for row in group):
                continue

        counts = [row.get("tiebreak_count", 0) for row in group]
        totals = [row.get("tiebreak_total", 0) for row in group]

        # ยังไม่มีผลรอบพิเศษครบทุกคน หรือจำนวนเที่ยวไม่เท่ากัน = ต้องยิง/ยิงต่อ
        if min(counts) == 0 or len(set(counts)) > 1:
            result.update(group_ids)
            continue

        # ยิงรอบพิเศษครบเท่ากันแล้ว แต่คะแนนพิเศษยังเท่ากัน = ต้องยิงต่อ
        if len(set(totals)) == 1:
            result.update(group_ids)

    return result

def exact_cut_ids(rows: list[dict], cutoff_count: int) -> set[int]:
    """เลือกผู้ผ่านแบบจำนวนตายตัว และไม่ปล่อยอันดับที่ยังเสมอเข้าไปก่อน

    ใช้กับเข้ารอบตรง/ผ่านจากรอบ 2/สร้าง bracket:
    ต้องเรียงลำดับจริงด้วย TOTAL -> 5 -> 3 -> Shoot-off
    ถ้ากลุ่มเสมอแตะตำแหน่งที่มีผล ให้รอ Shoot-off ก่อน

    สิทธิ์ตีรอบ 2 จากรอบแรกไม่ได้ใช้ฟังก์ชันนี้ แต่ยึด Class ถึงอันดับ
    ที่ตั้งไว้ (เช่น <= 16) และรับทุกคนที่มี Class เท่ากับอันดับตัด
    """
    if not cutoff_count or cutoff_count <= 0:
        return set()
    if not rows:
        return set()
    if len(rows) <= cutoff_count:
        pending = unresolved_tie_ids(rows, len(rows))
        return {row["athlete"].id for row in rows if row["athlete"].id not in pending}
    pending = unresolved_tie_ids(rows, cutoff_count)
    return {row["athlete"].id for row in rows[:cutoff_count] if row["athlete"].id not in pending}

def shootoff_group_ids(rows: list[dict], athlete_id: int, round_no: int | None = None) -> list[int]:
    target = next((row for row in rows if row["athlete"].id == athlete_id), None)
    if not target:
        return [athlete_id]
    key = base_shootoff_key(target)
    group = [row for row in rows if base_shootoff_key(row) == key]
    if round_no is not None:
        # ปุ่ม Shoot-off ต้องส่งเฉพาะคนในกลุ่มที่ตีจบรอบนั้นแล้ว
        # กันรอบ 2 ไปพ่วงคนรอคิว/คนเข้ารอบตรง ทำให้บันทึกแล้วดูเหมือนไม่เข้า
        group = [row for row in group if athlete_round_status(row["athlete"], round_no) == "finished"]
    ids = [row["athlete"].id for row in group]
    return ids or [athlete_id]


def _all_rows_finished_for_round(rows: list[dict], round_no: int) -> bool:
    """กันไม่ให้ขึ้น Shoot-off ก่อนกรรมการกดบันทึก/จบการตี"""
    if not rows:
        return False
    return all(athlete_round_status(row["athlete"], round_no) == "finished" for row in rows if row.get("athlete"))


def is_round_one_complete(event: Event) -> bool:
    """รอบ 1 จะถือว่าจบเมื่อผู้เล่นของอีเวนต์ทุกคนจบการยิงรอบ 1 แล้ว

    ระหว่างแข่ง Ranking ยังแสดงสดได้ตามคะแนนปัจจุบัน แต่สิทธิ์ผ่านตรง,
    รายชื่อรอบ 2, Shoot-off เพื่อคัด Top และ Cut line จะยังไม่ถูกฟันธง
    จนกว่ารอบ 1 จะจบครบทุกคน
    """
    rows = build_round_ranking(event, 1)
    return _all_rows_finished_for_round(rows, 1)


def round1_overview_unresolved_shootoff_ids(event: Event, rows: list[dict]) -> set[int]:
    """Shoot-off รอบ 1 เฉพาะกลุ่มที่เกี่ยวกับอันดับเข้ารอบตรง

    อันดับ 1..direct ต้องแยกด้วย TOTAL -> 5 -> 3 -> Shoot-off
    ตั้งแต่อันดับถัดไป ไม่ทำ Shoot-off เพื่อจัด Class; คะแนนรวมเท่ากันให้ Class เท่ากันได้
    สิทธิ์ไปตีรอบ 2 ใช้ Class ถึงอันดับที่กำหนด เช่น <= 16 และรับ Class 16 ทุกคน
    """
    if not rows:
        return set()
    direct = direct_quota(event)
    if direct <= 0:
        return set()
    return overview_unresolved_shootoff_ids(rows, direct, round_no=1)

def overview_shootoff_ids(event: Event, round_no: int) -> set[int]:
    """ตรวจ Shoot-off สำหรับ Overview

    รอบ 1:
    - อันดับเข้ารอบตรงใช้ TOTAL -> 5 -> 3 -> Shoot-off
    - หลังพ้นโควตาเข้ารอบตรง คะแนนรวมเท่ากันให้ Class เท่ากันได้
    - สิทธิ์ตีรอบ 2 ใช้ Class ถึงอันดับที่กำหนด เช่น <= 16 และรับ Class 16 ทุกคน

    รอบ 2:
    - ที่นั่งผ่านจากรอบ 2 ใช้ SUM(R1+R2) -> จำนวน 5 รวม -> จำนวน 3 รวม -> Shoot-off
    - หลังพ้นโควตาผ่าน SUM เท่ากันให้ Class เท่ากันได้
    """
    if round_no == 1:
        if not is_round_one_complete(event):
            return set()
        rows = build_round_ranking(event, 1)
        return round1_overview_unresolved_shootoff_ids(event, rows)

    if round_no == 2 and event.has_round_two:
        rows = [
            r for r in build_round_two_overview_rows(event)
            if not r.get("is_round2_direct_placeholder")
            and r.get("round2_has_played")
            and athlete_round_status(r["athlete"], 2) == "finished"
        ]
        rows.sort(key=lambda row: (
            -row.get("combined_total", row.get("total", 0)),
            -row.get("count_5", 0),
            -row.get("count_3", 0),
            -row.get("tiebreak_total", 0),
            row.get("display_order") if row.get("display_order") is not None else 999999,
            row["athlete"].id,
        ))
        # รอบ 2 ยิง Shoot-off เฉพาะกลุ่มที่เกี่ยวกับโควตาผ่านจากรอบ 2
        # หลังพ้นโควตาให้ SUM เท่ากันและ Class เท่ากันได้
        return overview_unresolved_shootoff_ids(
            rows,
            round2_advancer_quota(event),
            round_no=2,
        )

    return set()

def build_round_two_overview_rows(event: Event) -> List[dict]:
    """หน้า Overview รอบ 2 แบบ Official + realtime

    รายชื่อรอบ 2 จะเกิดขึ้นหลังรอบ 1 จบครบทั้งอีเวนต์แล้วเท่านั้น
    จากนั้นจึงแสดง realtime ของรอบ 2 โดยเก็บทั้งคะแนนรอบ 1/รอบ 2/SUM/RANKING
    """
    if not is_round_one_complete(event):
        return []
    round1_rows = build_round_ranking(event, 1)
    direct_limit = direct_quota(event)
    lane_count = max(event.lane_count or 1, 1)

    direct_ids = exact_cut_ids(round1_rows, direct_quota(event))
    direct_pending_ids = unresolved_tie_ids(round1_rows, direct_quota(event))
    direct_rows = []
    for row in round1_rows:
        if row["athlete"].id in direct_ids:
            direct_row = {**row}
            direct_row["round_no"] = 2
            direct_row["round1_total"] = row["total"]
            direct_row["round2_total"] = None
            direct_row["combined_total"] = row["total"]
            direct_row["round1_by_station"] = row.get("by_station", {})
            direct_row["round2_by_station"] = None
            direct_row["is_round2_direct_placeholder"] = True
            direct_row["display_order"] = "-"
            direct_row["display_lane_no"] = "-"
            direct_row["display_lane_order"] = "-"
            direct_row["status"] = "direct"
            direct_row["score_complete"] = False
            direct_row["approved"] = False
            direct_rows.append(direct_row)

    # ผู้ผ่านตรงจากรอบ 1 ใช้อันดับที่ตัดสินไว้แล้ว
    for idx, direct_row in enumerate(direct_rows, start=1):
        direct_row["rank"] = idx
        direct_row["ordinal_rank"] = idx
        direct_row["view_order"] = idx
        direct_row["display_rank"] = idx

    # รายชื่อรอบ 2 ต้องมาจากสิทธิ์หลังรอบ 1 ไม่ใช่จากลายเซ็น/คะแนนรอบ 2
    round2_source_rows = [
        row for row in round1_rows
        if row["athlete"].id not in direct_ids
        and row["athlete"].id not in direct_pending_ids
        and is_round_two_candidate(event, row["athlete"])
    ]
    # ลำดับตีรอบ 2: คะแนนรอบ 1 น้อยสุดก่อน
    round2_source_rows.sort(key=lambda row: (
        row["total"], row["count_5"], row["count_3"], row["tiebreak_total"],
        row["athlete"].start_order if row["athlete"].start_order is not None else 999999,
        row["athlete"].id,
    ))

    round2_rows = []
    for idx, source_row in enumerate(round2_source_rows, start=1):
        athlete = source_row["athlete"]
        r2 = summarize_round(athlete.id, 2)
        combined_total = source_row["total"] + r2["total"]
        logical_status = athlete_round_status(athlete, 2)
        approved = athlete_round_is_approved(athlete, 2)
        score_complete = overview_score_complete(r2) or logical_status == "finished"
        row = {
            "athlete": athlete,
            "round_no": 2,
            # หน้า Overview รอบ 2:
            # total = คะแนนรอบ 2 อย่างเดียว
            # combined_total = คะแนนรอบ 1 + รอบ 2
            "total": r2["total"],
            "combined_total": combined_total,
            "round1_total": source_row["total"],
            "round2_total": r2["total"],
            "count_5": source_row["count_5"] + r2["count_5"],
            "count_3": source_row["count_3"] + r2["count_3"],
            # รอบ 2 ต้องใช้ Shoot-off ของรอบ 2 เท่านั้นในการแยกอันดับ SUM
            # ห้ามเอา Shoot-off รอบ 1 มาบวก เพราะจะทำให้กดบันทึก Shoot-off รอบ 2 แล้วระบบยังมองว่าเที่ยวไม่เท่ากัน/ไม่ถูกบันทึก
            "tiebreak_total": r2["tiebreak_total"],
            "tiebreak_count": r2.get("tiebreak_count", 0),
            "round1_tiebreak_total": source_row.get("tiebreak_total", 0),
            "round1_tiebreak_count": source_row.get("tiebreak_count", 0),
            "round2_tiebreak_total": r2["tiebreak_total"],
            "round2_tiebreak_count": r2.get("tiebreak_count", 0),
            "status": logical_status,
            "score_complete": score_complete,
            "approved": approved,
            "by_station": r2["by_station"],
            "round1_by_station": source_row.get("by_station", {}),
            "round2_by_station": r2["by_station"],
            "red_cards": r2["red_cards"],
            "display_order": idx,
            "display_lane_no": ((idx - 1) % lane_count) + 1,
            "display_lane_order": ((idx - 1) // lane_count) + 1,
            "round1_rank": source_row["rank"],
            "is_round2_direct_placeholder": False,
        }
        round2_rows.append(row)

    # หน้า R2 ต้องเรียงคิวก่อน: คะแนนรอบ 1 ต่ำสุดในกลุ่มรอบ 2 ได้ตีก่อน
    # เมื่อมีคนเริ่มตี/มีคะแนนแล้ว จึงจัดอันดับเฉพาะคนที่ตีแล้วไว้ด้านบน
    played_rows = []
    waiting_rows = []
    for row in round2_rows:
        has_played = (
            row["status"] in {"active", "finished"}
            or row.get("round2_total", 0) > 0
            or row.get("tiebreak_count", 0) > 0
        )
        row["round2_has_played"] = has_played
        if has_played:
            played_rows.append(row)
        else:
            waiting_rows.append(row)

    played_rows.sort(key=lambda row: (
        -row["combined_total"],
        -row["count_5"],
        -row["count_3"],
        -row["tiebreak_total"],
        row["display_order"],
        row["athlete"].id,
    ))
    waiting_rows.sort(key=lambda row: (
        row["display_order"],
        row["athlete"].start_order if row["athlete"].start_order is not None else 999999,
        row["athlete"].id,
    ))

    base = len(direct_rows)
    # คนที่ตีรอบ 2 แล้ว: 4 ที่นั่ง (หรือค่าที่ตั้งไว้) ต้องแยก seed ด้วย SUM -> 5รวม -> 3รวม -> Shoot-off
    # หลังพ้นโควตานี้ SUM เท่ากันให้ Class เท่ากัน
    apply_round2_display_ranking(
        played_rows,
        start_rank=base + 1,
        advancer_limit=round2_advancer_quota(event),
    )

    # คนที่ยังไม่ตี แสดงตามคิวชั่วคราว ไม่ถือเป็น Ranking สุดท้าย
    waiting_base = base + len(played_rows)
    for idx, row in enumerate(waiting_rows, start=1):
        real_rank = waiting_base + idx
        row["rank"] = real_rank
        row["ordinal_rank"] = real_rank
        row["display_rank"] = real_rank
        row["view_order"] = real_rank

    return direct_rows + played_rows + waiting_rows

def scorecard_print_positions() -> dict:
    return {
        "header": {
            "bib_no": {"left": 165, "top": 18},
            "name": {"left": 380, "top": 18},
            "affiliation": {"left": 925, "top": 18},
        },
        "rows": {
            1: {"top": 362},  # รอบที่ 1
            2: {"top": 437},  # รอบที่ 2
            3: {"top": 512},  # รอบ 16 คน (เมื่อเปิดใช้สาย 16 คน) หรือรอบ 8 คนในอีเวนต์เดิม
            4: {"top": 587},  # รอบ 8 คน หรือรอบรองชนะเลิศในอีเวนต์เดิม
            5: {"top": 662},  # รอบรองชนะเลิศ หรือรอบชิงชนะเลิศในอีเวนต์เดิม
            6: {"top": 737},  # รอบชิงชนะเลิศ เมื่อเปิดใช้สาย 16 คน
        },
        "station_cols": {
            1: {"6": 170, "7": 211, "8": 252, "9": 293, "total": 334},
            2: {"6": 378, "7": 419, "8": 460, "9": 501, "total": 542},
            3: {"6": 586, "7": 627, "8": 668, "9": 709, "total": 750},
            4: {"6": 794, "7": 835, "8": 876, "9": 917, "total": 958},
            5: {"6": 1002, "7": 1043, "8": 1084, "9": 1125, "total": 1166},
        },
        "right_cols": {
            "grand_total": 1262,
            "rank": 1350,
            "athlete_signature": 1448,
        },
        "signature_rows": {
            1: {"judge": 776, "recorder": 776},
            2: {"judge": 814, "recorder": 814},
            3: {"judge": 852, "recorder": 852},
            4: {"judge": 890, "recorder": 890},
            5: {"judge": 928, "recorder": 928},
            6: {"judge": 966, "recorder": 966},
        },
    }


def scorecard_round_to_bracket_round(round_no: int, event: Event) -> str | None:
    """แปลงเลขรอบของ Scorecard เป็นชื่อรอบ bracket สำหรับหน้าพิมพ์รวม."""
    if round_no < 3:
        return None
    if event_has_round_of_16(event):
        return {3: "R16", 4: "QF", 5: "SF", 6: "F"}.get(round_no)
    return {3: "QF", 4: "SF", 5: "F"}.get(round_no)


def bracket_participant_order_map(event: Event, round_no: int) -> dict[int, dict]:
    """คืนข้อมูลลำดับ/สนามของนักกีฬาในรอบ bracket เพื่อใช้พิมพ์ Scorecard รวม."""
    round_name = scorecard_round_to_bracket_round(round_no, event)
    if not round_name:
        return {}

    matches = BracketMatch.query.filter_by(
        event_id=event.id,
        round_name=round_name,
    ).order_by(BracketMatch.match_no.asc()).all()

    result: dict[int, dict] = {}
    lane_count = max(event.lane_count or 1, 1)
    running_order = 1
    for match in matches:
        for side_no, athlete_id in enumerate([match.athlete_a_id, match.athlete_b_id], start=1):
            if not athlete_id or athlete_id in result:
                continue
            result[athlete_id] = {
                "display_order": running_order,
                "display_lane_no": ((match.match_no - 1) % lane_count) + 1,
                "display_lane_order": match.match_no,
                "match_no": match.match_no,
                "match_side": side_no,
            }
            running_order += 1
    return result


def athletes_for_scorecard_round(event: Event, round_no: int, selected_ids: list[int] | None = None) -> list[Athlete]:
    """เลือกรายชื่อนักกีฬาที่ควรพิมพ์ Scorecard รวมในแต่ละรอบ.

    - รอบ 1: นักกีฬาทั้งหมด
    - รอบ 2: เฉพาะผู้มีสิทธิ์ตีรอบ 2
    - รอบ bracket: เฉพาะคนที่อยู่ในคู่ของรอบนั้น
    """
    selected_set = {int(aid) for aid in selected_ids or []}

    def keep_selected(athletes: list[Athlete]) -> list[Athlete]:
        if not selected_set:
            return athletes
        return [athlete for athlete in athletes if athlete.id in selected_set]

    if round_no == 1:
        athletes = sorted(event.athletes, key=lambda athlete: (
            athlete.start_order if athlete.start_order is not None else 999999,
            athlete.id,
        ))
        return keep_selected(athletes)

    if round_no == 2:
        if not event.has_round_two:
            return []
        sync_round_two_candidates(event)
        candidate_ids = round_two_candidate_ids(event)
        start_rows = build_round_two_start_list(event)
        ordered_ids = [row["athlete"].id for row in start_rows if row["athlete"].id in candidate_ids]
        seen = set()
        athletes: list[Athlete] = []
        for athlete_id in ordered_ids:
            athlete = Athlete.query.get(athlete_id)
            if athlete and athlete.id not in seen:
                athletes.append(athlete)
                seen.add(athlete.id)

        # เผื่อข้อมูลเก่ามี signature รอบ 2 อยู่แล้ว แต่ไม่อยู่ใน start_rows
        for athlete in sorted(event.athletes, key=lambda a: (a.start_order, a.id)):
            if athlete.id in candidate_ids and athlete.id not in seen:
                athletes.append(athlete)
                seen.add(athlete.id)
        return keep_selected(athletes)

    round_name = scorecard_round_to_bracket_round(round_no, event)
    if not round_name:
        return []

    # สร้าง/เติม bracket ที่ยังขาดก่อน เพื่อให้ปุ่มพิมพ์รอบ knockout ใช้งานได้แม้ยังไม่เคยเปิดหน้าสาย
    ensure_bracket(event)
    order_map = bracket_participant_order_map(event, round_no)
    if not order_map:
        return []

    athletes_by_id = {athlete.id: athlete for athlete in event.athletes}
    ordered_ids = sorted(
        order_map,
        key=lambda aid: (
            order_map[aid].get("display_order", 999999),
            athletes_by_id.get(aid).start_order if athletes_by_id.get(aid) else 999999,
            aid,
        ),
    )
    athletes = [athletes_by_id[aid] for aid in ordered_ids if aid in athletes_by_id]
    return keep_selected(athletes)


def build_scorecard_print_context(athlete: Athlete, round_no: int) -> dict:
    """รวมข้อมูลให้ template พิมพ์ Scorecard แบบหลายคน โดยใช้โครงเดียวกับหน้าพิมพ์รายคน."""
    event = athlete.event
    template_data = build_scorecard_template_data(athlete.id)
    ranks = compute_round_ranks(event)

    round_ranks = {rn: "" for rn in ALL_SCORECARD_ROUNDS}
    for rn in [1, 2]:
        round_ranks[rn] = ranks.get(rn, {}).get(athlete.id, "")
    for rn in scorecard_round_numbers(event):
        if rn >= 3:
            row = next(
                (item for item in build_round_ranking(event, rn) if item["athlete"].id == athlete.id),
                None,
            )
            if row:
                round_ranks[rn] = row.get("rank", "")

    round_signatures = {
        rn: get_round_signature(athlete.id, rn)
        for rn in scorecard_round_numbers(event)
    }

    round_station_running_totals = {}
    for rn in scorecard_round_numbers(event):
        running = {}
        acc = 0
        for st in STATIONS:
            val = template_data["station_totals"].get((rn, st), 0)
            acc += val
            running[st] = acc
        round_station_running_totals[rn] = running

    display_order = athlete.start_order
    display_lane_no = athlete.lane_no
    display_lane_order = athlete.lane_order

    if round_no == 2 and event.has_round_two:
        for row in build_round_two_overview_rows(event):
            if row.get("athlete") and row["athlete"].id == athlete.id:
                display_order = row.get("display_order", display_order)
                display_lane_no = row.get("display_lane_no", display_lane_no)
                display_lane_order = row.get("display_lane_order", display_lane_order)
                break
    elif round_no >= 3:
        order_map = bracket_participant_order_map(event, round_no)
        if athlete.id in order_map:
            display_order = order_map[athlete.id].get("display_order", display_order)
            display_lane_no = order_map[athlete.id].get("display_lane_no", display_lane_no)
            display_lane_order = order_map[athlete.id].get("display_lane_order", display_lane_order)
    else:
        current_round_rows = build_round_ranking(event, round_no)
        current_row = next((row for row in current_round_rows if row["athlete"].id == athlete.id), None)
        if current_row:
            display_order = current_row.get("display_order", display_order)
            display_lane_no = current_row.get("display_lane_no", display_lane_no)
            display_lane_order = current_row.get("display_lane_order", display_lane_order)

    combined_rows = build_combined_qualifiers(event) if event.has_round_two else []
    current_combined = next(
        (row for row in combined_rows if row.get("athlete") and row["athlete"].id == athlete.id),
        None,
    )

    return {
        "athlete": athlete,
        "event": event,
        "score_map": template_data["score_map"],
        "station_totals": template_data["station_totals"],
        "station_reds": template_data["station_reds"],
        "round_totals": template_data["round_totals"],
        "round_ranks": round_ranks,
        "round_signatures": round_signatures,
        "round_station_running_totals": round_station_running_totals,
        "combined_rows": combined_rows,
        "current_combined": current_combined,
        "display_order": display_order,
        "display_lane_no": display_lane_no,
        "display_lane_order": display_lane_order,
    }



def get_progression_groups(event: Event) -> dict:
    cache = _request_cache()
    cache_key = ("progression_groups", event.id)
    if has_request_context() and cache_key in cache:
        return cache[cache_key]
    round1_rows = build_round_ranking(event, 1)
    round1_complete = is_round_one_complete(event)
    # ระหว่างรอบ 1 ยังไม่จบ: แสดง Ranking สดได้ แต่ยังไม่ประกาศสิทธิ์ผ่านตรง/R2
    direct_ids = exact_cut_ids(round1_rows, direct_quota(event)) if round1_complete else set()
    round2_candidate_ids = set()
    passed_round2_ids = set()
    eliminated_ids = set()
    if event.has_round_two and round1_complete:
        round2_candidate_ids = round_two_candidate_ids(event) - direct_ids
        combined = build_combined_qualifiers(event)
        passed_round2_ids = {
            r["athlete"].id for r in combined
            if r.get("round2_total") is not None and r.get("passed_cut")
        }
    for athlete in event.athletes:
        aid = athlete.id
        if aid in direct_ids:
            continue
        if aid in round2_candidate_ids and aid not in passed_round2_ids:
            eliminated_ids.add(aid)
    result = {"direct": direct_ids, "round2_candidates": round2_candidate_ids, "round2_passed": passed_round2_ids, "eliminated": eliminated_ids}
    if has_request_context():
        cache[cache_key] = result
    return result


def compute_round_ranks(event: Event) -> dict[int, dict[int,int]]:
    result = {}
    for rn in [1,2]:
        result[rn] = {}
        for row in build_round_ranking(event, rn):
            result[rn][row["athlete"].id] = row["rank"]
    return result


def configured_bracket_start_round(event: Event) -> str:
    return {
        "รอบ 16 คน": "R16",
        "รอบ 8 คน": "QF",
        "รอบ 4 คน": "SF",
        "รอบรองชนะเลิศ": "SF",
    }.get(event.next_round_label, "QF")


def ensure_bracket(event: Event) -> list[BracketMatch]:
    """สร้างตาราง bracket ล่วงหน้าทุกรอบถึงรอบชิง

    - หน้าประกบคู่จะเห็นช่องรอครบตั้งแต่ต้น เช่น R16 -> QF -> SF -> Final
    - ยังไม่ดึงชื่อขึ้นรอบถัดไปจนกว่าจะบันทึกผู้ชนะ แต่กล่องรอจะแสดงไว้แล้ว
    - ไม่ขยายจำนวนคนเองจากกรณีคะแนนเท่ากันที่เส้นตัด
    """
    desired_round = configured_bracket_start_round(event)
    qualifiers = build_combined_qualifiers(event)
    seeds = [row for row in qualifiers if row.get("seed")]
    seed_total = len(seeds)
    if seed_total > 4 and desired_round == "SF":
        desired_round = "QF"

    existing = BracketMatch.query.filter_by(event_id=event.id).all()
    existing_rounds = {m.round_name for m in existing}
    initial_round = "R16" if "R16" in existing_rounds else ("QF" if "QF" in existing_rounds else ("SF" if "SF" in existing_rounds else None))
    if existing and initial_round != desired_round and not any(m.winner_id for m in existing):
        BracketMatch.query.filter_by(event_id=event.id).delete(synchronize_session=False)
        db.session.commit()
        existing = []

    orders = {
        "R16": [(1,16),(8,9),(5,12),(4,13),(3,14),(6,11),(7,10),(2,15)],
        "QF": [(1,8),(4,5),(3,6),(2,7)],
        "SF": [(1,4),(2,3)],
        "F": [(1,2)],
    }
    rounds_after = {"R16": ["R16", "QF", "SF", "F"], "QF": ["QF", "SF", "F"], "SF": ["SF", "F"]}
    existing_keys = {(m.round_name, m.match_no): m for m in BracketMatch.query.filter_by(event_id=event.id).all()}
    changed = False
    for round_name in rounds_after.get(desired_round, [desired_round, "F"]):
        for idx, pair in enumerate(orders[round_name], start=1):
            if (round_name, idx) in existing_keys:
                continue
            a_id = b_id = None
            if round_name == desired_round:
                a, b = pair
                arow = seeds[a - 1] if len(seeds) >= a else None
                brow = seeds[b - 1] if len(seeds) >= b else None
                a_id = arow["athlete"].id if arow else None
                b_id = brow["athlete"].id if brow else None
            db.session.add(BracketMatch(event_id=event.id, round_name=round_name, match_no=idx, athlete_a_id=a_id, athlete_b_id=b_id))
            changed = True
    if changed:
        db.session.commit()
    maybe_advance_bracket(event)
    return BracketMatch.query.filter_by(event_id=event.id).order_by(BracketMatch.round_name, BracketMatch.match_no).all()


def maybe_advance_bracket(event: Event) -> None:
    matches = BracketMatch.query.filter_by(event_id=event.id).all()
    r16 = sorted([m for m in matches if m.round_name == "R16"], key=lambda m: m.match_no)
    qf = sorted([m for m in matches if m.round_name == "QF"], key=lambda m: m.match_no)
    sf = sorted([m for m in matches if m.round_name == "SF"], key=lambda m: m.match_no)
    fn = sorted([m for m in matches if m.round_name == "F"], key=lambda m: m.match_no)
    changed = False

    def fill_match(match, a_id, b_id):
        nonlocal changed
        if match and match.winner_id is None:
            if match.athlete_a_id != a_id or match.athlete_b_id != b_id:
                match.athlete_a_id = a_id
                match.athlete_b_id = b_id
                changed = True

    if r16 and len(r16) >= 8 and qf and all(m.winner_id for m in r16):
        pairings = [(r16[0].winner_id, r16[1].winner_id), (r16[2].winner_id, r16[3].winner_id), (r16[4].winner_id, r16[5].winner_id), (r16[6].winner_id, r16[7].winner_id)]
        for idx, (a, b) in enumerate(pairings):
            fill_match(qf[idx], a, b)
    if qf and len(qf) >= 4 and sf and all(m.winner_id for m in qf):
        pairings = [(qf[0].winner_id, qf[1].winner_id), (qf[2].winner_id, qf[3].winner_id)]
        for idx, (a, b) in enumerate(pairings):
            fill_match(sf[idx], a, b)
    if sf and len(sf) >= 2 and fn and all(m.winner_id for m in sf):
        fill_match(fn[0], sf[0].winner_id, sf[1].winner_id)
    if changed:
        db.session.commit()


def sync_match_winner_from_scores(match: BracketMatch) -> None:
    event = Event.query.get(match.event_id)
    round_no = bracket_round_to_scorecard_round(match.round_name, event)
    if not match.athlete_a_id or not match.athlete_b_id:
        return
    sig_a = get_round_signature(match.athlete_a_id, round_no)
    sig_b = get_round_signature(match.athlete_b_id, round_no)
    if not (sig_a and sig_a.finished_at and sig_b and sig_b.finished_at):
        return
    total_a = summarize_round(match.athlete_a_id, round_no)["total"]
    total_b = summarize_round(match.athlete_b_id, round_no)["total"]
    if total_a == total_b:
        return
    match.winner_id = match.athlete_a_id if total_a > total_b else match.athlete_b_id
    db.session.commit()


def build_bracket_match_row(event: Event, athlete: Athlete | None, round_name: str, seed_map: dict[int, int]) -> dict:
    if not athlete:
        return {
            "athlete": None,
            "athlete_id": None,
            "team": "-",
            "name": "-",
            "r1": "-",
            "r1r2": "",
            "show_ref_score": False,
            "is_direct_qualifier": False,
            "stations": ["-"] * 5,
            "total": "-",
            "seed": "",
            "status": "waiting",
            "status_label": "รอคิว",
        }

    progression_groups = get_progression_groups(event)
    is_direct_qualifier = athlete.id in progression_groups.get("direct", set())

    r1 = summarize_round(athlete.id, 1)["total"]
    r2 = summarize_round(athlete.id, 2)["total"] if event.has_round_two else 0
    ref_total = r1 + (r2 or 0)
    show_ref_score = not is_direct_qualifier

    round_no = bracket_round_to_scorecard_round(round_name, event)
    current = summarize_round(athlete.id, round_no)
    round_status = athlete_round_status(athlete, round_no)

    return {
        "athlete": athlete,
        "athlete_id": athlete.id,
        "team": athlete.affiliation,
        "name": athlete.name,
        "r1": r1,
        "r1r2": ref_total if show_ref_score else "",
        "show_ref_score": show_ref_score,
        "is_direct_qualifier": is_direct_qualifier,
        "stations": [current["by_station"][station]["total"] for station in STATIONS],
        "total": current["total"],
        "seed": seed_map.get(athlete.id, ""),
        "round_no": round_no,
        "status": round_status,
        "status_label": {"waiting": "รอคิว", "active": "กำลังตี", "finished": "ตีเสร็จแล้ว"}.get(round_status, "รอคิว"),
    }

def generate_next_bib_no(event_id: int) -> str:
    count = Athlete.query.filter_by(event_id=event_id).count()
    return str(count + 1)


def recalculate_event_orders(event: Event) -> None:
    athletes = Athlete.query.filter_by(event_id=event.id).order_by(Athlete.start_order, Athlete.id).all()
    for idx, athlete in enumerate(athletes, start=1):
        athlete.start_order = idx
        athlete.bib_no = str(idx)
        athlete.lane_no = ((idx - 1) % event.lane_count) + 1
        athlete.lane_order = ((idx - 1) // event.lane_count) + 1


def reset_event_bracket(event: Event) -> None:
    BracketMatch.query.filter_by(event_id=event.id).delete()
    db.session.commit()


def normalize_header(value: str) -> str:
    return (value or '').strip().lower()


def parse_athletes_excel(file_storage) -> list[tuple[str, str]]:
    workbook = load_workbook(file_storage, data_only=True)
    sheet = workbook.active
    headers = [normalize_header(cell.value if cell.value is not None else '') for cell in sheet[1]]
    try:
        name_idx = headers.index('ชื่อ')
        affiliation_idx = headers.index('สังกัด')
    except ValueError as exc:
        raise ValueError('ไฟล์ Excel ต้องมีหัวคอลัมน์ชื่อ และ สังกัด') from exc

    rows: list[tuple[str, str]] = []
    for row_no, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        name = str(row[name_idx]).strip() if len(row) > name_idx and row[name_idx] is not None else ''
        affiliation = str(row[affiliation_idx]).strip() if len(row) > affiliation_idx and row[affiliation_idx] is not None else ''
        if not name and not affiliation:
            continue
        if not name or not affiliation:
            raise ValueError(f'แถวที่ {row_no} ต้องมีทั้งชื่อและสังกัด')
        rows.append((name, affiliation))
    if not rows:
        raise ValueError('ไม่พบข้อมูลนักกีฬาในไฟล์ Excel')
    return rows

def dashboard_stats() -> dict:
    events_count = Event.query.count()
    athletes_count = Athlete.query.count()
    round1_rows = []
    for event in Event.query.all():
        round1_rows.extend(build_round_ranking(event, 1))
    top_score = max(round1_rows, key=lambda r: r["total"], default=None)
    affiliation_best = {}
    for row in round1_rows:
        key = row["athlete"].affiliation
        if key not in affiliation_best or row["total"] > affiliation_best[key]["total"]:
            affiliation_best[key] = row
    return {
        "events_count": events_count,
        "athletes_count": athletes_count,
        "top_score": top_score,
        "top_affiliations": sorted(affiliation_best.values(), key=lambda r: r["total"], reverse=True)[:8],
    }


@app.context_processor
def inject_globals():
    return {
        "now": datetime.now(),
        "MAX_RED_CARDS": MAX_RED_CARDS,
        "court_public_id": court_public_id,
    }


@app.route("/set-language/<lang>")
def set_language(lang):
    if lang in SUPPORTED_LANGS:
        session["lang"] = lang
    next_url = request.args.get("next") or request.referrer or url_for("index")
    return redirect(next_url)


def is_court_user() -> bool:
    return bool(
        current_user.is_authenticated
        and current_user.role == "court"
        and current_user.court_no
        and current_user.court_event_id
    )


def court_public_id(user: User | None = None) -> str:
    user = user or current_user
    if not user or getattr(user, "role", None) != "court" or not getattr(user, "court_no", None):
        return getattr(user, "username", "") if user else ""
    return f"court{int(user.court_no):02d}"


def court_can_access_event(event: Event) -> bool:
    if not is_court_user():
        return True
    return int(current_user.court_event_id) == int(event.id)


def effective_lane_for_athlete(event: Event, athlete: Athlete, round_no: int) -> int | None:
    """เลขสนามที่นักกีฬาถูกจัดให้ในรอบที่กำลังดู/คีย์"""
    if round_no == 1:
        return athlete.lane_no
    if round_no == 2 and event.has_round_two:
        row = next((r for r in build_round_two_overview_rows(event) if r["athlete"].id == athlete.id), None)
        lane = row.get("display_lane_no") if row else None
        return lane if isinstance(lane, int) else None
    # รอบ Knockout/รอบอื่น ใช้ข้อมูล display lane ถ้ามี
    try:
        row = next((r for r in build_round_ranking(event, round_no) if r["athlete"].id == athlete.id), None)
        lane = row.get("display_lane_no") if row else athlete.lane_no
        return lane if isinstance(lane, int) else athlete.lane_no
    except Exception:
        return athlete.lane_no


def court_can_access_athlete(athlete: Athlete, round_no: int) -> bool:
    if not is_court_user():
        return True
    if not court_can_access_event(athlete.event):
        return False
    return effective_lane_for_athlete(athlete.event, athlete, round_no) == current_user.court_no


def filter_rows_for_court(rows: list[dict]) -> list[dict]:
    if not is_court_user():
        return rows
    court_no = current_user.court_no
    return [r for r in rows if r.get("display_lane_no", r["athlete"].lane_no) == court_no]


def sync_event_court_users(event: Event) -> tuple[int, int]:
    """Keep event-scoped court IDs aligned with event.lane_count.

    Missing IDs are pre-created in a disabled state (!UNSET!) until superadmin
    sets a password. IDs above the current lane_count are removed.
    Returns (created_count, removed_count).
    """
    lane_count = max(1, int(event.lane_count or 1))
    existing = {
        int(u.court_no): u
        for u in User.query.filter_by(role="court", court_event_id=event.id).all()
        if u.court_no
    }
    created = 0
    removed = 0
    for court_no in range(1, lane_count + 1):
        internal_username = f"event{event.id}_court{court_no:02d}"
        user = existing.get(court_no)
        if user is None:
            user = User(
                username=internal_username,
                password_hash="!UNSET!",
                role="court",
                court_no=court_no,
                court_event_id=event.id,
            )
            db.session.add(user)
            created += 1
        else:
            user.username = internal_username
    for court_no, user in existing.items():
        if court_no > lane_count:
            db.session.delete(user)
            removed += 1
    return created, removed


@app.route("/")
def index():
    if is_court_user():
        return redirect(url_for("court_queue", event_id=current_user.court_event_id, round=1))
    stats = dashboard_stats()
    events = Event.query.order_by(Event.competition_date.desc(), Event.id.desc()).all()
    return render_template("index.html", events=events, stats=stats)


@app.route("/events/<int:event_id>/courts/create", methods=["POST"])
@login_required
@role_required("superadmin")
def create_event_court_user(event_id: int):
    event = Event.query.get_or_404(event_id)
    try:
        court_no = int(request.form.get("court_no", "0"))
    except ValueError:
        court_no = 0
    password = request.form.get("password", "").strip()
    if court_no < 1 or court_no > event.lane_count or not password:
        flash(f"กรุณาระบุเลขสนาม 1-{event.lane_count} และรหัสผ่าน", "warning")
        return redirect(url_for("manage_athletes", event_id=event.id))
    if password_problem(password):
        flash(password_problem(password), "warning")
        return redirect(url_for("manage_athletes", event_id=event.id))

    # Court IDs are normally pre-created from lane_count. Sync here as a safety net
    # for older events created before this feature existed.
    sync_event_court_users(event)
    user = User.query.filter_by(role="court", court_event_id=event.id, court_no=court_no).first()
    if not user:
        abort(500)
    user.set_password(password)
    db.session.commit()
    flash(f"ตั้งรหัสให้ court{court_no:02d} แล้ว", "success")
    return redirect(url_for("manage_athletes", event_id=event.id))


@app.route("/events/<int:event_id>/courts/<int:user_id>/delete", methods=["POST"])
@login_required
@role_required("superadmin")
def delete_event_court_user(event_id: int, user_id: int):
    event = Event.query.get_or_404(event_id)
    user = User.query.get_or_404(user_id)
    if user.role != "court" or user.court_event_id != event.id:
        abort(404)
    public_id = court_public_id(user)
    db.session.delete(user)
    db.session.commit()
    flash(f"ลบ {public_id} ของอีเวนต์นี้แล้ว", "info")
    return redirect(url_for("manage_athletes", event_id=event.id))


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = None
        # Court staff type only court01/court02/... . The DB username is event-scoped
        # (event12_court01) so different events never collide internally.
        if username.lower().startswith("court") and username[5:].isdigit():
            court_no = int(username[5:])
            matches = [
                u for u in User.query.filter_by(role="court", court_no=court_no).all()
                if u.court_event_id and u.check_password(password)
            ]
            if len(matches) == 1:
                user = matches[0]
            elif len(matches) > 1:
                flash("มี court ID นี้มากกว่า 1 อีเวนต์ที่ใช้รหัสเดียวกัน กรุณาเปลี่ยนรหัสของอีเวนต์ให้ไม่ซ้ำ", "danger")
                return render_template("login.html")
        else:
            user = User.query.filter_by(username=username).first()
        if user and user.check_password(password) and password in KNOWN_DEFAULT_PASSWORDS:
            flash("บัญชีนี้ยังใช้รหัสผ่านตั้งต้นที่เปิดเผยในโค้ด จึงถูกระงับ ให้ superadmin ตั้งรหัสใหม่ที่หน้าผู้ใช้", "danger")
            return render_template("login.html")
        if user and user.check_password(password):
            session.clear()
            login_user(user)
            flash("เข้าสู่ระบบสำเร็จ", "success")
            if user.role == "court" and user.court_event_id:
                return redirect(url_for("court_queue", event_id=user.court_event_id, round=1))
            return redirect(url_for("index"))
        flash("ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง", "danger")
    return render_template("login.html")


@app.route("/logout")
@login_required
def logout():
    logout_user()
    flash("ออกจากระบบแล้ว", "info")
    return redirect(url_for("index"))


@app.route("/admin/users", methods=["GET", "POST"])
@login_required
@role_required("superadmin")
def manage_users():
    if request.method == "POST" and request.form.get("action") == "reset_password":
        target = User.query.get(request.form.get("user_id", type=int) or 0)
        new_password = request.form.get("new_password", "")
        problem = password_problem(new_password)
        if not target or target.role == "court":
            flash("ไม่พบผู้ใช้ (บัญชีสนามให้เปลี่ยนรหัสจากหน้าอีเวนต์)", "danger")
        elif problem:
            flash(problem, "danger")
        else:
            target.set_password(new_password)
            db.session.commit()
            flash(f"ตั้งรหัสผ่านใหม่ให้ {target.username} แล้ว", "success")
        return redirect(url_for("manage_users"))
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        role = request.form.get("role", "user")
        if role not in {"user", "admin", "superadmin"}:
            flash("สิทธิ์ไม่ถูกต้อง", "danger")
            return redirect(url_for("manage_users"))
        if password_problem(password):
            flash(password_problem(password), "danger")
            return redirect(url_for("manage_users"))
        if role == "court":
            flash("บัญชีสนามให้สร้างจากหน้านักกีฬาของอีเวนต์", "warning")
            return redirect(url_for("manage_users"))
        if username and password and not User.query.filter_by(username=username).first():
            user = User(username=username, role=role)
            user.set_password(password)
            db.session.add(user)
            db.session.commit()
            flash("สร้างผู้ใช้สำเร็จ", "success")
        else:
            flash("สร้างผู้ใช้ไม่สำเร็จ กรุณาตรวจสอบข้อมูล", "danger")
    users = User.query.filter(User.role != "court").order_by(User.id).all()
    weak_users = {u.id for u in users if any(u.check_password(p) for p in KNOWN_DEFAULT_PASSWORDS)}
    return render_template("users.html", users=users, weak_users=weak_users)


@app.route("/events/new", methods=["GET", "POST"])
@login_required
@role_required("admin", "superadmin")
def create_event():
    if request.method == "POST":
        event = Event(
            name=request.form["name"].strip(),
            event_group=request.form["event_group"],
            category=request.form["category"],
            competition_date=date.fromisoformat(request.form["competition_date"]),
            location=request.form.get("location", "").strip(),
            lane_count=int(request.form["lane_count"]),
            direct_qualifiers=int(request.form["direct_qualifiers"]),
            has_round_two=request.form.get("has_round_two") == "yes",
            round_two_cutoff_rank=int(request.form["round_two_cutoff_rank"]) if request.form.get("round_two_cutoff_rank") else None,
            round_two_mode=parse_round_two_mode(request.form),
            next_round_label=request.form["next_round_label"],
            round_two_advancers=int(request.form.get("round_two_advancers") or 4),
            created_by=current_user.id,
        )
        db.session.add(event)
        db.session.flush()
        created_courts, _ = sync_event_court_users(event)
        db.session.commit()
        flash(f"สร้างอีเวนต์สำเร็จ · เตรียม Court ID {created_courts} สนามแล้ว", "success")
        return redirect(url_for("manage_athletes", event_id=event.id))
    return render_template("event_form.html")


@app.route("/events/<int:event_id>/edit", methods=["GET", "POST"])
@login_required
@role_required("admin", "superadmin")
def edit_event(event_id: int):
    event = Event.query.get_or_404(event_id)
    if request.method == "POST":
        event.name = request.form["name"].strip()
        event.event_group = request.form["event_group"]
        event.category = request.form["category"]
        event.competition_date = date.fromisoformat(request.form["competition_date"])
        event.location = request.form.get("location", "").strip()
        event.lane_count = int(request.form["lane_count"])
        event.direct_qualifiers = int(request.form["direct_qualifiers"])
        event.has_round_two = request.form.get("has_round_two") == "yes"
        event.round_two_cutoff_rank = int(request.form["round_two_cutoff_rank"]) if request.form.get("round_two_cutoff_rank") else None
        event.round_two_mode = parse_round_two_mode(request.form)
        event.next_round_label = request.form["next_round_label"]
        event.round_two_advancers = int(request.form.get("round_two_advancers") or 4)
        recalculate_event_orders(event)
        created_courts, removed_courts = sync_event_court_users(event)
        db.session.commit()
        reset_event_bracket(event)
        flash("แก้ไขอีเวนต์สำเร็จ", "success")
        return redirect(url_for("event_overview", event_id=event.id, round=1))
    return render_template("event_form.html", event=event, is_edit=True)


@app.route("/events/<int:event_id>/delete", methods=["POST"])
@login_required
@role_required("admin", "superadmin")
def delete_event(event_id: int):
    event = Event.query.get_or_404(event_id)
    BracketMatch.query.filter_by(event_id=event.id).delete()
    User.query.filter_by(role="court", court_event_id=event.id).delete(synchronize_session=False)
    db.session.delete(event)
    db.session.commit()
    flash("ลบอีเวนต์แล้ว", "info")
    return redirect(url_for("index"))


@app.route("/events/<int:event_id>/athletes", methods=["GET", "POST"])
@login_required
@role_required("admin", "superadmin")
def manage_athletes(event_id: int):
    event = Event.query.get_or_404(event_id)
    if request.method == "POST":
        name = request.form["name"].strip()
        affiliation = request.form["affiliation"].strip()
        next_order = Athlete.query.filter_by(event_id=event.id).count() + 1
        lane_no = ((next_order - 1) % event.lane_count) + 1
        lane_order = ((next_order - 1) // event.lane_count) + 1
        athlete = Athlete(
            event_id=event.id,
            bib_no=str(next_order),
            name=name,
            affiliation=affiliation,
            start_order=next_order,
            lane_no=lane_no,
            lane_order=lane_order,
            status="waiting",
        )
        db.session.add(athlete)
        db.session.commit()
        flash("เพิ่มนักกีฬาสำเร็จ", "success")
        return redirect(url_for("manage_athletes", event_id=event.id))
    athletes = Athlete.query.filter_by(event_id=event.id).order_by(Athlete.start_order).all()
    # Backfill court IDs automatically for older events too.
    if current_user.role == "superadmin":
        created_courts, removed_courts = sync_event_court_users(event)
        if created_courts or removed_courts:
            db.session.commit()
    court_users = (
        User.query.filter_by(role="court", court_event_id=event.id)
        .order_by(User.court_no, User.id)
        .all()
    )
    return render_template("athletes.html", event=event, athletes=athletes, court_users=court_users)


@app.route("/athletes/<int:athlete_id>/round2-toggle", methods=["POST"])
@login_required
@role_required("admin", "superadmin")
def toggle_athlete_round_two(athlete_id: int):
    athlete = Athlete.query.get_or_404(athlete_id)
    event = athlete.event
    if not event.has_round_two:
        flash("อีเวนต์นี้ไม่ได้เปิดใช้รอบ 2", "warning")
        return redirect(url_for("manage_athletes", event_id=event.id))

    athlete.round_two_disabled = not bool(getattr(athlete, "round_two_disabled", False))
    if athlete.round_two_disabled:
        athlete.round_two_disabled_at = datetime.utcnow()
        athlete.round_two_disabled_by = current_user.id
        message = f"ปิด {athlete.name} ไม่ให้ระบบส่งไปรอบ 2 แล้ว"
    else:
        athlete.round_two_disabled_at = None
        athlete.round_two_disabled_by = None
        message = f"เปิด {athlete.name} ให้ระบบพิจารณารอบ 2 ตามกติกาแล้ว"

    db.session.commit()
    clear_request_cache()
    sync_round_two_candidates(event)
    reset_event_bracket(event)
    db.session.commit()
    flash(message, "success")
    return redirect(url_for("manage_athletes", event_id=event.id))


@app.route("/athletes/<int:athlete_id>/delete", methods=["POST"])
@login_required
@role_required("admin", "superadmin")
def delete_athlete(athlete_id: int):
    athlete = Athlete.query.get_or_404(athlete_id)
    event = athlete.event
    athlete_name = athlete.name
    db.session.delete(athlete)
    db.session.commit()
    recalculate_event_orders(event)
    sync_round_two_candidates(event)
    reset_event_bracket(event)
    db.session.commit()
    flash(f"ลบรายการ {athlete_name} เรียบร้อย", "success")
    return redirect(url_for("manage_athletes", event_id=event.id))


@app.route("/events/<int:event_id>/athletes/import", methods=["POST"])
@login_required
@role_required("admin", "superadmin")
def import_athletes_excel(event_id: int):
    event = Event.query.get_or_404(event_id)
    file = request.files.get("excel_file")
    if not file or not file.filename:
        flash("กรุณาเลือกไฟล์ Excel", "danger")
        return redirect(url_for("manage_athletes", event_id=event.id))
    if not file.filename.lower().endswith(".xlsx"):
        flash("รองรับเฉพาะไฟล์ .xlsx", "danger")
        return redirect(url_for("manage_athletes", event_id=event.id))
    try:
        rows = parse_athletes_excel(file)
        next_order = Athlete.query.filter_by(event_id=event.id).count() + 1
        for name, affiliation in rows:
            athlete = Athlete(
                event_id=event.id,
                bib_no=str(next_order),
                name=name,
                affiliation=affiliation,
                start_order=next_order,
                lane_no=((next_order - 1) % event.lane_count) + 1,
                lane_order=((next_order - 1) // event.lane_count) + 1,
                status="waiting",
            )
            db.session.add(athlete)
            next_order += 1
        db.session.commit()
        flash(f"นำเข้านักกีฬาสำเร็จ {len(rows)} คน", "success")
    except Exception as exc:
        db.session.rollback()
        flash(str(exc), "danger")
    return redirect(url_for("manage_athletes", event_id=event.id))


@app.route("/athletes-import-template.xlsx")
def athletes_import_template():
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Athletes"
    sheet.append(["ชื่อ", "สังกัด"])
    sheet.append(["นายตัวอย่าง ใจดี", "ขอนแก่น"])
    sheet.append(["นางสาวตัวอย่าง แสนดี", "อุดรธานี"])
    stream = BytesIO()
    workbook.save(stream)
    stream.seek(0)
    return send_file(
        stream,
        as_attachment=True,
        download_name="athletes_import_template.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/events/<int:event_id>/athletes/randomize", methods=["POST"])
@login_required
@role_required("admin", "superadmin")
def randomize_athletes(event_id: int):
    event = Event.query.get_or_404(event_id)
    draw_event_lots(event)
    db.session.commit()
    invalidate_poll_cache(event.id)
    flash("สุ่มลำดับใหม่แล้ว", "success")
    return redirect(url_for("manage_athletes", event_id=event.id))


# ---------------------------------------------------------------------------
# จับสลากหลายอีเวนต์พร้อมกัน + สร้างหลายอีเวนต์ในครั้งเดียว
# ---------------------------------------------------------------------------
import random as _random

_DRAW_RNG = _random.SystemRandom()  # สุ่มจากระบบปฏิบัติการ เดาลำดับล่วงหน้าไม่ได้

EVENT_GROUP_PRESETS = ["ทั่วไป", "อาวุโส", "เยาวชน", "รุ่นอายุ 12 ปี", "รุ่นอายุ 14 ปี", "รุ่นอายุ 16 ปี", "รุ่นอายุ 18 ปี"]
EVENT_CATEGORIES = ["ชาย", "หญิง", "ผสม"]
NEXT_ROUND_LABELS = ["รอบ 16 คน", "รอบ 8 คน", "รอบ 4 คน", "รอบรองชนะเลิศ"]
MAX_BULK_EVENTS = 60


def draw_event_lots(event: Event, rng=None) -> int:
    """จับสลากลำดับยิงและสนามใหม่ทั้งอีเวนต์ คืนจำนวนนักกีฬาที่จับ"""
    rng = rng or _DRAW_RNG
    athletes = Athlete.query.filter_by(event_id=event.id).all()
    rng.shuffle(athletes)
    lanes = max(1, int(event.lane_count or 1))
    for idx, athlete in enumerate(athletes, start=1):
        athlete.start_order = idx
        athlete.lane_no = ((idx - 1) % lanes) + 1
        athlete.lane_order = ((idx - 1) // lanes) + 1
    return len(athletes)


def event_ids_with_scores(event_ids: list[int]) -> set[int]:
    """อีเวนต์ที่มีการกรอกคะแนนแล้วอย่างน้อย 1 ช่อง (ไม่ควรจับสลากซ้ำ)"""
    if not event_ids:
        return set()
    rows = (
        db.session.query(Athlete.event_id)
        .join(ScoreEntry, ScoreEntry.athlete_id == Athlete.id)
        .filter(Athlete.event_id.in_(event_ids), ScoreEntry.is_scored.is_(True))
        .distinct()
        .all()
    )
    return {int(r[0]) for r in rows}


def _form_int(value, default: int, minimum: int = 0) -> int:
    try:
        return max(minimum, int(str(value).strip()))
    except (TypeError, ValueError):
        return default


def quota_for_next_round(label: str) -> int:
    label = label or ""
    if "16" in label:
        return 8
    if "4" in label or "รอง" in label:
        return 2
    return 4


def new_event_from_settings(name: str, event_group: str, category: str, lane_count: int, shared: dict) -> Event:
    next_label = shared.get("next_round_label") or "รอบ 8 คน"
    return Event(
        name=name.strip()[:255],
        event_group=(event_group or "ทั่วไป").strip()[:50],
        category=category if category in EVENT_CATEGORIES else "ชาย",
        competition_date=shared["competition_date"],
        location=(shared.get("location") or "").strip()[:255],
        lane_count=max(1, int(lane_count or 1)),
        direct_qualifiers=shared.get("direct_qualifiers", quota_for_next_round(next_label)),
        has_round_two=bool(shared.get("has_round_two", True)),
        round_two_cutoff_rank=shared.get("round_two_cutoff_rank"),
        round_two_mode=shared.get("round_two_mode", "cutoff"),
        next_round_label=next_label,
        round_two_advancers=shared.get("round_two_advancers", quota_for_next_round(next_label)),
        created_by=current_user.id if current_user.is_authenticated else None,
    )


def parse_shared_event_settings(form) -> dict:
    """ค่าที่ใช้ร่วมกันทุกอีเวนต์ในการสร้างแบบหลายรายการ"""
    raw_date = (form.get("competition_date") or "").strip()
    try:
        competition_date = date.fromisoformat(raw_date)
    except ValueError as exc:
        raise ValueError("กรุณาระบุวันแข่งขัน") from exc
    next_label = form.get("next_round_label") or "รอบ 8 คน"
    if next_label not in NEXT_ROUND_LABELS:
        next_label = "รอบ 8 คน"
    quota = quota_for_next_round(next_label)
    has_round_two = form.get("has_round_two", "yes") == "yes"
    cutoff = form.get("round_two_cutoff_rank")
    return {
        "competition_date": competition_date,
        "location": form.get("location", ""),
        "next_round_label": next_label,
        "direct_qualifiers": _form_int(form.get("direct_qualifiers"), quota),
        "has_round_two": has_round_two,
        "round_two_cutoff_rank": _form_int(cutoff, 16, 1) if cutoff else None,
        "round_two_mode": parse_round_two_mode(form),
        "round_two_advancers": _form_int(form.get("round_two_advancers"), quota, 1),
    }


def lanes_for_entries(entry_count: int, per_lane: int, fallback: int) -> int:
    if entry_count <= 0 or per_lane <= 0:
        return max(1, fallback)
    return max(1, -(-entry_count // per_lane))  # ปัดขึ้น


def _cell_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if value != value:  # NaN
            return ""
        if value.is_integer():
            value = int(value)
    return str(value).strip()


def _read_workbook_rows(file_storage) -> list[list[list[str]]]:
    """อ่านทุกชีตของไฟล์ .xlsx/.xls เป็นตารางข้อความ"""
    filename = (file_storage.filename or "").lower()
    data = file_storage.read()
    sheets: list[list[list[str]]] = []
    if filename.endswith(".xls"):
        try:
            import xlrd  # type: ignore
        except ImportError as exc:
            raise ValueError("เซิร์ฟเวอร์ยังไม่ได้ติดตั้ง xlrd สำหรับอ่านไฟล์ .xls กรุณาบันทึกไฟล์เป็น .xlsx") from exc
        book = xlrd.open_workbook(file_contents=data)
        for sheet in book.sheets():
            sheets.append([[_cell_text(sheet.cell_value(r, c)) for c in range(sheet.ncols)] for r in range(sheet.nrows)])
    else:
        book = load_workbook(BytesIO(data), data_only=True, read_only=True)
        for sheet in book.worksheets:
            sheets.append([[_cell_text(v) for v in row] for row in sheet.iter_rows(values_only=True)])
    return sheets


def _guess_category(title: str) -> str:
    if "ผสม" in title:
        return "ผสม"
    if "หญิง" in title:
        return "หญิง"
    return "ชาย"


def _guess_event_group(title: str) -> str:
    import re
    match = re.search(r"รุ่น\s*อายุ\s*(\d+)\s*ปี", title)
    if match:
        return f"รุ่นอายุ {match.group(1)} ปี"
    match = re.search(r"รุ่น\s*([^\s]+)", title)
    if match:
        return f"รุ่น{match.group(1)}"[:50]
    return "ทั่วไป"


def parse_entry_list_workbook(file_storage, keyword: str = "") -> list[dict]:
    """อ่านไฟล์รายชื่อแบบหลายประเภทในไฟล์เดียว

    รองรับรูปแบบใบสมัครที่มีหัวข้อแต่ละประเภท เช่น "ประเภทชู้ตติ้งชาย รุ่นอายุ 12 ปี"
    ตามด้วยหัวตาราง (ที่ | ชื่อ... | อำเภอ | จังหวัด) และรายชื่อด้านล่าง
    """
    import re
    keyword = (keyword or "").strip()
    events: list[dict] = []
    seen_titles: set[str] = set()
    for rows in _read_workbook_rows(file_storage):
        current = None
        cols = None
        for row in rows:
            first = row[0] if row else ""
            filled = [c for c in row if c]
            # หัวข้อประเภท: มีคำว่า "ประเภท" อยู่ในช่องแรกและไม่มีข้อมูลช่องอื่นเป็นตาราง
            if first.startswith("ประเภท") and len(filled) <= 2 and not first.replace("ประเภท", "").strip() == "":
                title = " ".join(first.split())
                title = re.sub(r"\s*/.*$", "", title).strip()  # ตัด "/ 54 ทีม" ท้ายหัวข้อ
                current = None
                cols = None
                if keyword and keyword not in title:
                    continue
                if title in seen_titles:
                    continue
                seen_titles.add(title)
                current = {
                    "title": title,
                    "event_group": _guess_event_group(title),
                    "category": _guess_category(title),
                    "entries": [],
                }
                events.append(current)
                continue
            if current is None:
                continue
            if not filled:
                # แถวว่างหลังรายชื่อ = จบรายการของหัวข้อนี้
                if current["entries"]:
                    current = None
                continue
            if cols is None:
                lowered = [c.replace(" ", "") for c in row]
                name_idx = next((i for i, c in enumerate(lowered) if c.startswith("ชื่อ")), None)
                if name_idx is not None:
                    district_idx = next((i for i, c in enumerate(lowered) if c in {"อำเภอ", "เขต"}), None)
                    province_idx = next((i for i, c in enumerate(lowered) if c == "จังหวัด"), None)
                    aff_idx = next((i for i, c in enumerate(lowered) if c == "สังกัด"), None)
                    no_idx = next((i for i, c in enumerate(lowered) if c in {"ที่", "ลำดับ", "ลำดับที่"}), None)
                    cols = {"name": name_idx, "district": district_idx, "province": province_idx,
                            "affiliation": aff_idx, "no": no_idx}
                    last_no = 0
                continue

            def cell(key):
                idx = cols.get(key)
                return row[idx].strip() if idx is not None and idx < len(row) else ""

            name = cell("name")
            if not name:
                continue
            seq = cell("no")
            if seq.isdigit():
                # เลขลำดับเริ่มใหม่โดยไม่มีหัวข้อ = ข้อมูลค้างจากตารางอื่น หยุดอ่านหัวข้อนี้
                if int(seq) <= last_no:
                    current = None
                    continue
                last_no = int(seq)
            affiliation = cell("affiliation") or cell("province") or cell("district") or name
            current["entries"].append({"name": name[:255], "affiliation": affiliation[:255]})
    return [e for e in events if e["entries"]]


@app.route("/events/bulk-new", methods=["GET", "POST"])
@login_required
@role_required("admin", "superadmin")
def bulk_create_events():
    """สร้างหลายอีเวนต์พร้อมกันจากตารางที่กรอก"""
    if request.method == "POST":
        try:
            shared = parse_shared_event_settings(request.form)
        except ValueError as exc:
            flash(str(exc), "danger")
            return redirect(url_for("bulk_create_events"))
        names = request.form.getlist("row_name")
        groups = request.form.getlist("row_group")
        categories = request.form.getlist("row_category")
        lanes = request.form.getlist("row_lanes")
        default_lanes = _form_int(request.form.get("default_lane_count"), 4, 1)
        rows = []
        for idx, name in enumerate(names):
            name = (name or "").strip()
            if not name:
                continue
            rows.append((
                name,
                groups[idx] if idx < len(groups) else "ทั่วไป",
                categories[idx] if idx < len(categories) else "ชาย",
                _form_int(lanes[idx] if idx < len(lanes) else "", default_lanes, 1),
            ))
        if not rows:
            flash("กรุณาเพิ่มอย่างน้อย 1 อีเวนต์", "warning")
            return redirect(url_for("bulk_create_events"))
        if len(rows) > MAX_BULK_EVENTS:
            flash(f"สร้างได้ครั้งละไม่เกิน {MAX_BULK_EVENTS} อีเวนต์", "warning")
            return redirect(url_for("bulk_create_events"))
        created = []
        for name, group, category, lane_count in rows:
            event = new_event_from_settings(name, group, category, lane_count, shared)
            db.session.add(event)
            db.session.flush()
            sync_event_court_users(event)
            created.append(event)
        db.session.commit()
        flash(f"สร้างอีเวนต์สำเร็จ {len(created)} รายการ", "success")
        return redirect(url_for("draw_center", ids=",".join(str(e.id) for e in created)))
    return render_template(
        "events_bulk_form.html",
        group_presets=EVENT_GROUP_PRESETS,
        categories=EVENT_CATEGORIES,
        next_round_labels=NEXT_ROUND_LABELS,
        today=date.today().isoformat(),
    )


@app.route("/events/bulk-import", methods=["POST"])
@login_required
@role_required("admin", "superadmin")
def bulk_import_events():
    """อัปโหลดไฟล์รายชื่อหลายประเภท → แสดงตัวอย่างให้เลือกก่อนสร้าง"""
    file = request.files.get("entry_file")
    if not file or not file.filename:
        flash("กรุณาเลือกไฟล์ Excel", "danger")
        return redirect(url_for("bulk_create_events"))
    if not file.filename.lower().endswith((".xlsx", ".xls")):
        flash("รองรับไฟล์ .xlsx และ .xls", "danger")
        return redirect(url_for("bulk_create_events"))
    keyword = request.form.get("keyword", "").strip()
    try:
        events = parse_entry_list_workbook(file, keyword)
    except Exception as exc:  # ไฟล์เสีย/รูปแบบไม่รองรับ
        flash(f"อ่านไฟล์ไม่สำเร็จ: {exc}", "danger")
        return redirect(url_for("bulk_create_events"))
    if not events:
        flash("ไม่พบหัวข้อประเภทที่มีรายชื่อในไฟล์" + (f" (คำค้น: {keyword})" if keyword else ""), "warning")
        return redirect(url_for("bulk_create_events"))
    per_lane = _form_int(request.form.get("per_lane"), 6, 0)
    default_lanes = _form_int(request.form.get("default_lane_count"), 4, 1)
    prefix = request.form.get("name_prefix", "").strip()
    for item in events:
        item["lanes"] = lanes_for_entries(len(item["entries"]), per_lane, default_lanes)
        item["name"] = f"{prefix} {item['title']}".strip() if prefix else item["title"]
    return render_template(
        "events_bulk_import_preview.html",
        events=events,
        events_json=json.dumps(events, ensure_ascii=False),
        source_name=file.filename,
        total_entries=sum(len(e["entries"]) for e in events),
        categories=EVENT_CATEGORIES,
        next_round_labels=NEXT_ROUND_LABELS,
        today=date.today().isoformat(),
        per_lane=per_lane,
    )


@app.route("/events/bulk-import/confirm", methods=["POST"])
@login_required
@role_required("admin", "superadmin")
def bulk_import_events_confirm():
    try:
        shared = parse_shared_event_settings(request.form)
        payload = json.loads(request.form.get("events_json") or "[]")
    except (ValueError, json.JSONDecodeError) as exc:
        flash(str(exc) or "ข้อมูลไม่ถูกต้อง", "danger")
        return redirect(url_for("bulk_create_events"))
    selected = {int(i) for i in request.form.getlist("include") if str(i).isdigit()}
    chosen = [(i, item) for i, item in enumerate(payload) if i in selected]
    if not chosen:
        flash("ยังไม่ได้เลือกประเภทที่จะสร้าง", "warning")
        return redirect(url_for("bulk_create_events"))
    if len(chosen) > MAX_BULK_EVENTS:
        flash(f"สร้างได้ครั้งละไม่เกิน {MAX_BULK_EVENTS} อีเวนต์", "warning")
        return redirect(url_for("bulk_create_events"))
    draw_now = request.form.get("draw_now") == "yes"
    created = []
    total_athletes = 0
    for i, item in chosen:
        name = (request.form.get(f"name_{i}") or item.get("name") or item.get("title") or "").strip()
        if not name:
            continue
        group = request.form.get(f"group_{i}") or item.get("event_group") or "ทั่วไป"
        category = request.form.get(f"category_{i}") or item.get("category") or "ชาย"
        lanes = _form_int(request.form.get(f"lanes_{i}"), int(item.get("lanes") or 1), 1)
        event = new_event_from_settings(name, group, category, lanes, shared)
        db.session.add(event)
        db.session.flush()
        sync_event_court_users(event)
        for order, entry in enumerate(item.get("entries") or [], start=1):
            entry_name = str(entry.get("name") or "").strip()[:255]
            if not entry_name:
                continue
            db.session.add(Athlete(
                event_id=event.id,
                bib_no=str(order),
                name=entry_name,
                affiliation=(str(entry.get("affiliation") or "").strip() or entry_name)[:255],
                start_order=order,
                lane_no=((order - 1) % event.lane_count) + 1,
                lane_order=((order - 1) // event.lane_count) + 1,
                status="waiting",
            ))
            total_athletes += 1
        db.session.flush()
        if draw_now:
            draw_event_lots(event)
        created.append(event)
    db.session.commit()
    msg = f"สร้างอีเวนต์ {len(created)} รายการ · นำเข้ารายชื่อ {total_athletes} รายการ"
    if draw_now:
        msg += " · จับสลากเรียบร้อย"
    flash(msg, "success")
    return redirect(url_for("draw_center", ids=",".join(str(e.id) for e in created), drawn="1" if draw_now else None))


def _parse_id_list(raw: str | None) -> list[int]:
    return [int(x) for x in (raw or "").split(",") if x.strip().isdigit()]


def build_draw_lane_table(event: Event) -> list[dict]:
    athletes = Athlete.query.filter_by(event_id=event.id).order_by(Athlete.lane_no, Athlete.lane_order, Athlete.start_order).all()
    lanes: dict[int, list[Athlete]] = {}
    for athlete in athletes:
        lanes.setdefault(int(athlete.lane_no or 1), []).append(athlete)
    return [{"lane_no": lane_no, "athletes": lanes[lane_no]} for lane_no in sorted(lanes)]


@app.route("/events/draw", methods=["GET", "POST"])
@login_required
@role_required("admin", "superadmin")
def draw_center():
    """ศูนย์ควบคุมการจับสลาก: เลือกหลายอีเวนต์แล้วสั่งจับพร้อมกันในครั้งเดียว"""
    if request.method == "POST":
        ids = sorted({int(i) for i in request.form.getlist("event_ids") if str(i).isdigit()})
        if not ids:
            flash("กรุณาเลือกอีเวนต์ที่จะจับสลาก", "warning")
            return redirect(url_for("draw_center"))
        force = request.form.get("force") == "yes" and current_user.role == "superadmin"
        events = Event.query.filter(Event.id.in_(ids)).all()
        started = event_ids_with_scores([e.id for e in events])
        drawn, skipped = [], []
        for event in events:
            if event.id in started and not force:
                skipped.append(event)
                continue
            draw_event_lots(event)
            if event.id in started:
                reset_event_bracket_no_commit(event)
            drawn.append(event)
        db.session.commit()  # ทุกอีเวนต์ถูกจับและบันทึกพร้อมกันใน transaction เดียว
        for event in drawn:
            invalidate_poll_cache(event.id)
        drawn_at = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        if drawn:
            flash(f"จับสลากพร้อมกัน {len(drawn)} อีเวนต์ เวลา {drawn_at}", "success")
        if skipped:
            flash("ข้ามอีเวนต์ที่เริ่มกรอกคะแนนแล้ว: " + ", ".join(e.name for e in skipped), "warning")
        return redirect(url_for("draw_center", ids=",".join(str(e.id) for e in drawn) or None, drawn="1" if drawn else None))

    focus_ids = _parse_id_list(request.args.get("ids"))
    all_events = Event.query.order_by(Event.competition_date.desc(), Event.id.desc()).all()
    counts = dict(
        db.session.query(Athlete.event_id, func.count(Athlete.id)).group_by(Athlete.event_id).all()
    )
    started = event_ids_with_scores([e.id for e in all_events])
    dates = sorted({e.competition_date for e in all_events}, reverse=True)
    results = []
    if focus_ids:
        by_id = {e.id: e for e in all_events}
        results = [
            {"event": by_id[i], "lanes": build_draw_lane_table(by_id[i])}
            for i in focus_ids if i in by_id
        ]
    return render_template(
        "draw_center.html",
        events=all_events,
        counts=counts,
        started=started,
        dates=dates,
        focus_ids=set(focus_ids),
        results=results,
        just_drawn=request.args.get("drawn") == "1",
    )


def reset_event_bracket_no_commit(event: Event) -> None:
    BracketMatch.query.filter_by(event_id=event.id).delete()



def apply_overview_cut_lines(event: Event, rows: list[dict], round_no: int, round_complete: bool, groups: dict) -> None:
    """กำหนดเส้นแบ่งสิทธิ์บน Overview แบบสดตาม Class.

    ROUND 1 มี 2 เส้น:
    1) QUARTERFINALS = หลังโควตาผ่านตรงเข้าสู่ Knockout
    2) QUALIFIED FOR ROUND 2 = หลังคนสุดท้ายที่ Class <= cutoff rank
       เช่น cutoff 16: ถ้ามี Class 16 หลายคน เส้นอยู่หลัง Class 16 คนสุดท้าย

    ROUND 2 มีเส้น QUARTERFINALS หลังจำนวนผู้ผ่านจากรอบ 2 ตามที่ตั้งค่าไว้.
    """
    for row in rows:
        row["cut_line_after"] = False
        row["cut_line_label"] = ""
    if not rows:
        return

    def mark(idx: int | None, label: str) -> None:
        if idx is None:
            return
        if 0 <= idx < len(rows) - 1:
            rows[idx]["cut_line_after"] = True
            rows[idx]["cut_line_label"] = label

    if round_no == 1:
        direct = max(direct_quota(event), 0)
        cutoff_rank = max(round2_cutoff_rank(event), 0) if event.has_round_two else 0

        if round_complete:
            last_direct_idx = None
            last_round2_candidate_idx = None
            for idx, row in enumerate(rows):
                aid = row["athlete"].id
                if aid in groups.get("direct", set()):
                    last_direct_idx = idx
                if aid in groups.get("round2_candidates", set()):
                    last_round2_candidate_idx = idx
            mark(last_direct_idx, "QUARTERFINALS")
            mark(last_round2_candidate_idx, "QUALIFIED FOR ROUND 2")
            return

        # LIVE: เส้นผ่านตรงขึ้นเมื่อมีคนยิงจบอย่างน้อยตามจำนวน direct
        completed = [(idx, row) for idx, row in enumerate(rows) if row.get("score_complete")]
        if direct > 0 and len(completed) >= direct:
            mark(completed[direct - 1][0], "QUARTERFINALS")

        # LIVE: เส้นรอบ 2 ยึด Class ถึงอันดับที่กำหนด ไม่ได้นับจำนวนคน
        # จะแสดงเมื่อมีผลจบจนเกิด Class ถึง cutoff แล้ว และถ้า Class cutoff ซ้ำ
        # จะวางหลังคนสุดท้ายที่มี Class เท่ากับ cutoff
        if cutoff_rank > 0 and len(completed) >= cutoff_rank:
            eligible = []
            for idx, row in completed:
                rank = int(row.get("display_rank", row.get("rank", 999999)) or 999999)
                if rank <= cutoff_rank:
                    eligible.append((idx, row, rank))
            # เริ่มแสดงเมื่อมีผู้ตีจบอย่างน้อยถึงลำดับ cutoff แล้ว
            # ไม่บังคับว่าต้องมี Class เลข cutoff จริง เพราะ competition ranking
            # อาจกระโดด เช่น 15,15,17; ในกรณีนี้ทั้ง Class 15 ยังอยู่ในช่วง <=16
            if eligible:
                last_idx = max(idx for idx, _, _ in eligible)
                mark(last_idx, "QUALIFIED FOR ROUND 2")
        return

    if round_no == 2 and event.has_round_two:
        if round_complete:
            last_passed_idx = None
            for idx, row in enumerate(rows):
                if row["athlete"].id in groups.get("round2_passed", set()):
                    last_passed_idx = idx
            mark(last_passed_idx, "QUARTERFINALS")
            return

        quota = max(round2_advancer_quota(event), 0)
        completed_indices = [
            idx for idx, row in enumerate(rows)
            if not row.get("is_round2_direct_placeholder") and row.get("score_complete")
        ]
        if quota > 0 and len(completed_indices) >= quota:
            mark(completed_indices[quota - 1], "QUARTERFINALS")


@app.route("/events/<int:event_id>/overview")
def event_overview(event_id: int):
    event = Event.query.get_or_404(event_id)
    if is_court_user() and not court_can_access_event(event):
        abort(403)
    round_no = request.args.get("round", 1, type=int) or 1
    if round_no == 2 and event.has_round_two:
        sync_round_two_candidates(event)
    if round_no == 2 and event.has_round_two:
        all_rows = build_round_two_overview_rows(event)
        round2_competitors = [row for row in all_rows if not row.get("is_round2_direct_placeholder")]
        round_complete = bool(round2_competitors) and all(
            athlete_round_status(row["athlete"], 2) == "finished" for row in round2_competitors
        )
    else:
        all_rows = build_round_ranking(event, round_no)
        round_complete = _all_rows_finished_for_round(all_rows, round_no)
    rows = filter_rows_for_court(all_rows)
    groups = get_progression_groups(event)
    shoot_off_ids = overview_shootoff_ids(event, round_no)
    for row in rows:
        aid = row["athlete"].id
        row["shoot_off_required"] = aid in shoot_off_ids
        row["shoot_off_group_ids"] = shootoff_group_ids(rows, aid, round_no) if row["shoot_off_required"] else [aid]
        if row["shoot_off_required"]:
            row["progress_class"] = "shoot-off-required"
        elif round_no == 1:
            # หน้า Overview รอบ 1 แยกแค่เข้ารอบตรง/ได้สิทธิ์ตีรอบ 2 ก่อน ไม่ลงสี Knockout ที่นี่
            row["progress_class"] = "qualified-direct" if aid in groups["direct"] else ("round2-candidate" if aid in groups["round2_candidates"] else "")
        else:
            row["progress_class"] = ("qualified-round2" if aid in groups["round2_passed"] else ("eliminated" if aid in groups["eliminated"] else ""))
        row["cut_line_after"] = False

    apply_overview_cut_lines(event, rows, round_no, round_complete, groups)

    combined_rows = []
    return render_template(
        "overview.html",
        event=event,
        round_no=round_no,
        rows=rows,
        combined_rows=combined_rows,
        theme=event_theme(event.category),
        station_images=[f"station_{i}.png" for i in STATIONS],
        
    )


def court_queue_rows(event: Event, round_no: int, court_no: int) -> list[dict]:
    """นักกีฬาของสนามหนึ่งในรอบที่เลือก เรียงตามลำดับยิงในสนาม"""
    if round_no == 2 and event.has_round_two:
        sync_round_two_candidates(event)
        source = [r for r in build_round_two_overview_rows(event) if not r.get("is_round2_direct_placeholder")]
    else:
        source = build_round_ranking(event, 1)
        round_no = 1
    rows = []
    for row in source:
        lane = row.get("display_lane_no", row["athlete"].lane_no)
        if lane != court_no:
            continue
        athlete = row["athlete"]
        order = row.get("display_lane_order", athlete.lane_order)
        rows.append({
            "athlete": athlete,
            "order": order if isinstance(order, int) else 999,
            "status": athlete_round_status(athlete, round_no),
            "approved": athlete_round_is_approved(athlete, round_no),
            "total": row.get("total", 0),
        })
    rows.sort(key=lambda r: (r["order"], r["athlete"].start_order or 0, r["athlete"].id))
    next_row = next((r for r in rows if r["status"] == "active"), None) or next((r for r in rows if r["status"] == "waiting"), None)
    for r in rows:
        r["is_next"] = r is next_row
    return rows


@app.route("/events/<int:event_id>/court")
@login_required
@role_required("admin", "superadmin", "court")
def court_queue(event_id: int):
    """หน้าคิวของเจ้าหน้าที่สนาม: เห็นเฉพาะนักกีฬาในสนามตัวเอง ตามลำดับยิง พร้อมปุ่มเปิด Scorecard"""
    event = Event.query.get_or_404(event_id)
    if is_court_user() and not court_can_access_event(event):
        abort(403)
    round_no = request.args.get("round", 1, type=int) or 1
    if round_no not in (1, 2) or (round_no == 2 and not event.has_round_two):
        round_no = 1
    lane_count = max(event.lane_count or 1, 1)
    if is_court_user():
        court_no = current_user.court_no
    else:
        court_no = request.args.get("court", 1, type=int) or 1
        court_no = min(max(court_no, 1), lane_count)
    rows = court_queue_rows(event, round_no, court_no)
    done = sum(1 for r in rows if r["status"] == "finished")
    return render_template(
        "court_queue.html",
        event=event,
        round_no=round_no,
        court_no=court_no,
        lane_count=lane_count,
        rows=rows,
        done=done,
        theme=event_theme(event.category),
        work_page=True,
    )


@app.route("/events/<int:event_id>/overview-data")
@shared_poll_cache
def overview_data(event_id: int):
    event = Event.query.get_or_404(event_id)
    if is_court_user() and not court_can_access_event(event):
        return jsonify({"error": "forbidden"}), 403
    round_no = request.args.get("round", 1, type=int) or 1
    if round_no == 2 and event.has_round_two:
        sync_round_two_candidates(event)
    if round_no == 2 and event.has_round_two:
        all_rows = build_round_two_overview_rows(event)
        round2_competitors = [row for row in all_rows if not row.get("is_round2_direct_placeholder")]
        round_complete = bool(round2_competitors) and all(
            athlete_round_status(row["athlete"], 2) == "finished" for row in round2_competitors
        )
    else:
        all_rows = build_round_ranking(event, round_no)
        round_complete = _all_rows_finished_for_round(all_rows, round_no)
    rows = filter_rows_for_court(all_rows)
    groups = get_progression_groups(event)
    shoot_off_ids = overview_shootoff_ids(event, round_no)
    for row in rows:
        row["shoot_off_required"] = row["athlete"].id in shoot_off_ids
        row["shoot_off_group_ids"] = shootoff_group_ids(rows, row["athlete"].id, round_no) if row["shoot_off_required"] else [row["athlete"].id]
        row["cut_line_after"] = False
    apply_overview_cut_lines(event, rows, round_no, round_complete, groups)
    payload = []
    for row in rows:
        stations = {}
        for station in STATIONS:
            station_summary = row["by_station"][station]
            stations[str(station)] = {
                "6": station_summary["distances"].get(6, 0),
                "7": station_summary["distances"].get(7, 0),
                "8": station_summary["distances"].get(8, 0),
                "9": station_summary["distances"].get(9, 0),
                "total": station_summary["total"],
                "played": {
                    "6": bool(station_summary.get("played_distances", {}).get(6, False)),
                    "7": bool(station_summary.get("played_distances", {}).get(7, False)),
                    "8": bool(station_summary.get("played_distances", {}).get(8, False)),
                    "9": bool(station_summary.get("played_distances", {}).get(9, False)),
                },
                "played_count": station_summary.get("played_count", 0),
                "is_complete": station_summary.get("is_complete", False),
            }
        round1_stations = None
        round2_stations = None
        if round_no == 2 and event.has_round_two:
            round1_stations = {str(st): (row.get("round1_by_station") or {}).get(st, {}).get("total", 0) for st in STATIONS}
            if row.get("round2_by_station") is not None:
                round2_stations = {str(st): row["round2_by_station"].get(st, {}).get("total", 0) for st in STATIONS}
            else:
                round2_stations = {str(st): None for st in STATIONS}
        aid = row["athlete"].id
        payload.append({
            "rank": row["rank"],
            "display_rank": row.get("display_rank", row["rank"]),
            "name": row["athlete"].name,
            "affiliation": row["athlete"].affiliation,
            "start_order": row["athlete"].start_order,
            "lane_no": row["athlete"].lane_no,
            "lane_order": row["athlete"].lane_order,
            "display_order": row.get("display_order", row["athlete"].start_order),
            "display_lane_no": row.get("display_lane_no", row["athlete"].lane_no),
            "display_lane_order": row.get("display_lane_order", row["athlete"].lane_order),
            "status": row["status"],
            "score_complete": bool(row.get("score_complete", False)),
            "approved": bool(row.get("approved", False)),
            "total": row.get("round2_total", row["total"]) if round_no == 2 and event.has_round_two and not row.get("is_round2_direct_placeholder") else row["total"],
            "round1_total": row.get("round1_total"),
            "round2_total": row.get("round2_total"),
            "combined_total": row.get("combined_total", row["total"]),
            "round1_stations": round1_stations,
            "round2_stations": round2_stations,
            "athlete_id": row["athlete"].id,
            "progress_class": ("shoot-off-required" if row.get("shoot_off_required") else (("qualified-direct" if aid in groups["direct"] else ("round2-candidate" if aid in groups["round2_candidates"] else "")) if round_no == 1 else ("qualified-round2" if aid in groups["round2_passed"] else ("eliminated" if aid in groups["eliminated"] else "")))),
            "shoot_off_required": row.get("shoot_off_required", False),
            "shoot_off_group_ids": row.get("shoot_off_group_ids", [row["athlete"].id]),
            "is_round2_direct_placeholder": row.get("is_round2_direct_placeholder", False),
            "round2_has_played": row.get("round2_has_played", False),
            "cut_line_after": row.get("cut_line_after", False),
            "cut_line_label": row.get("cut_line_label", ""),
            # ใช้สำหรับเรียงแถว realtime: รอบ 1 ต้องเรียงตามคะแนน/Rank, รอบ 2 ใช้ view_order ที่ build_round_two_overview_rows กำหนด
            "view_order": row.get("view_order", row["rank"]),
            "stations": stations,
        })
    return jsonify(payload)


@app.route("/events/<int:event_id>/overview-stats")
@shared_poll_cache
def overview_stats(event_id: int):
    """สถิติ 5/3 สำหรับเปิดดูประกอบการจัดลำดับ โดยไม่ทำให้ตาราง Overview หลักรก"""
    event = Event.query.get_or_404(event_id)
    if is_court_user() and not court_can_access_event(event):
        return jsonify({"error": "forbidden"}), 403
    round_no = request.args.get("round", 1, type=int) or 1

    if round_no == 2 and event.has_round_two:
        rows = [r for r in build_round_two_overview_rows(event) if not r.get("is_round2_direct_placeholder")]
    else:
        rows = build_round_ranking(event, 1)
    rows = filter_rows_for_court(rows)

    # เรียงตามกติกาจริงที่ใช้ประกอบอันดับ: TOTAL -> 5 -> 3 -> Shoot-off
    rows = sorted(rows, key=lambda r: (
        -r.get("combined_total", r.get("total", 0)),
        -r.get("count_5", 0),
        -r.get("count_3", 0),
        -r.get("tiebreak_total", 0),
        r.get("display_order") if r.get("display_order") is not None else 999999,
        r["athlete"].id,
    ))

    data = []
    for idx, row in enumerate(rows, start=1):
        data.append({
            "rank": idx,
            "name": row["athlete"].name,
            "affiliation": row["athlete"].affiliation or "-",
            "total": row.get("combined_total", row.get("total", 0)),
            "count_5": row.get("count_5", 0),
            "count_3": row.get("count_3", 0),
            "status": ("ต้องตี Shoot-off" if row.get("shoot_off_required") else ""),
        })
    return jsonify(data)


@app.route("/api/scorecard/<int:athlete_id>/autosave", methods=["POST"])
@login_required
@role_required("admin", "superadmin", "court")
def autosave_scorecard(athlete_id: int):
    athlete = Athlete.query.get_or_404(athlete_id)
    payload = request.get_json(silent=True) or {}
    try:
        round_no = int(payload.get("round_no", 1))
        station_no = int(payload.get("station_no", 1))
        distance_m = int(payload.get("distance_m", 6))
    except (TypeError, ValueError):
        return jsonify({"ok": False, "message": "ข้อมูลรอบ/สถานี/ระยะไม่ถูกต้อง"}), 400
    if round_no not in ALL_SCORECARD_ROUNDS or station_no not in STATIONS or distance_m not in DISTANCES:
        return jsonify({"ok": False, "message": "ข้อมูลรอบ/สถานี/ระยะไม่ถูกต้อง"}), 400
    if not court_can_access_athlete(athlete, round_no):
        return jsonify({"ok": False, "message": f"บัญชีนี้คีย์ได้เฉพาะสนาม {current_user.court_no}"}), 403
    got_lock, lock_holder = acquire_scorecard_lock(athlete.id, round_no, str(payload.get("page_token", "")))
    if not got_lock:
        return jsonify({
            "ok": False,
            "locked_by": lock_holder,
            "message": f"ใบนี้กำลังถูกคีย์โดย {lock_holder} กด “รับช่วงคีย์ต่อ” ถ้าต้องการคีย์แทน",
        }), 423
    score_value = str(payload.get("score", "")).strip()
    red = bool(payload.get("red", False))
    played = bool(payload.get("played", False))
    if score_value != "":
        try:
            parsed_score = int(score_value)
        except ValueError:
            return jsonify({"ok": False, "message": "คะแนนต้องเป็นตัวเลข"}), 400
        allowed = ALLOWED_SCORES_BY_STATION[station_no]
        if parsed_score not in allowed:
            allowed_text = " / ".join(str(v) for v in sorted(allowed, reverse=True))
            return jsonify({"ok": False, "message": f"สถานี {station_no} ให้คะแนนได้เฉพาะ {allowed_text}"}), 400
    ensure_round_entries(athlete.id, round_no)
    signature = ensure_signature(athlete.id, round_no)
    if not signature.started_at:
        signature.started_at = datetime.utcnow()
    if not signature.finished_at:
        athlete.status = "active"
    entry = ScoreEntry.query.filter_by(athlete_id=athlete.id, round_no=round_no, station_no=station_no, distance_m=distance_m).first()
    old_state = (entry.score, bool(entry.is_red_card), bool(entry.is_scored)) if entry else (0, False, False)
    requested_value = 0 if score_value == "" else int(score_value)
    requested_state = (0 if red else requested_value, red, bool(played or red))
    is_change = old_state != requested_state
    already_signed = bool(signature.finished_at)
    edit_signature = str(payload.get("edit_signature", "")).strip()
    edit_reason = str(payload.get("edit_reason", "")).strip()
    if already_signed and is_change and not edit_signature:
        return jsonify({
            "ok": False,
            "requires_edit_signature": True,
            "message": "ผลนี้ลงลายเซ็นแล้ว กรุณาลงลายเซ็นผู้แก้ไขก่อนแก้คะแนน",
        }), 409

    round_entries = ScoreEntry.query.filter_by(
        athlete_id=athlete.id,
        round_no=round_no,
    ).all()
    existing_round_red = sum(1 for item in round_entries if item.is_red_card)

    # หลังใบแดงครั้งที่ 2 ห้ามแก้คะแนน/เพิ่มรายการอีก ยกเว้นการเอาใบแดงเดิมออกเพื่อแก้การกดผิด
    correcting_red = bool(entry and entry.is_red_card and not red)
    if existing_round_red >= MAX_RED_CARDS and not correcting_red:
        summary = summarize_round(athlete.id, round_no)
        return jsonify({
            "ok": False,
            "round_stopped": True,
            "round_red": existing_round_red,
            "station_total": summary["by_station"][station_no]["total"],
            "round_total": summary["total"],
            "message": "ใบแดงครบ 2 ครั้ง: ยุติการยิงรอบนี้และคงคะแนนที่ทำได้ไว้",
        }), 409

    if entry:
        value = requested_value
        entry.is_red_card = red
        entry.score = 0 if red else value
        entry.is_scored = bool(played or red)

    round_red = sum(1 for item in round_entries if item.is_red_card)
    round_stopped = round_red >= MAX_RED_CARDS
    if round_stopped:
        signature.stopped_by_red = True
        signature.finished_at = signature.finished_at or datetime.utcnow()
        athlete.status = "finished"
    elif signature.stopped_by_red:
        signature.stopped_by_red = False
        has_full_signoff = bool(signature.bypass_signed) or all([
            bool(signature.recorder_name or signature.recorder_signature),
            bool(signature.referee_name or signature.referee_signature),
            bool(signature.athlete_name or signature.athlete_signature),
        ])
        if not has_full_signoff:
            signature.finished_at = None
            athlete.status = "active"

    if already_signed and is_change:
        db.session.add(ScoreEditLog(
            athlete_id=athlete.id,
            round_no=round_no,
            station_no=station_no,
            distance_m=distance_m,
            old_score=old_state[0],
            new_score=requested_state[0],
            old_red=old_state[1],
            new_red=requested_state[1],
            old_played=old_state[2],
            new_played=requested_state[2],
            edited_by=current_user.id,
            editor_username=current_user.username,
            editor_court_no=current_user.court_no if current_user.role == "court" else None,
            editor_signature=edit_signature,
            reason=edit_reason or "แก้ไขคะแนนหลังลงลายเซ็นยืนยัน",
        ))

    db.session.commit()
    clear_request_cache()
    summary = summarize_round(athlete.id, round_no)
    station_entries = ScoreEntry.query.filter_by(
        athlete_id=athlete.id,
        round_no=round_no,
        station_no=station_no,
    ).all()
    station_red = sum(1 for e in station_entries if e.is_red_card)
    return jsonify({
        "ok": True,
        "station_total": summary["by_station"][station_no]["total"],
        "station_red": station_red,
        "round_red": round_red,
        "round_stopped": round_stopped,
        "round_total": summary["total"],
        "station_played_count": summary["by_station"][station_no].get("played_count", 0),
        "station_complete": summary["by_station"][station_no].get("is_complete", False),
    })


@app.route("/api/scorecard/<int:athlete_id>/lock", methods=["POST"])
@login_required
@role_required("admin", "superadmin", "court")
def scorecard_lock_api(athlete_id: int):
    athlete = Athlete.query.get_or_404(athlete_id)
    payload = request.get_json(silent=True) or request.form
    try:
        round_no = int(payload.get("round_no", 1))
    except (TypeError, ValueError):
        return jsonify({"ok": False, "message": "รอบไม่ถูกต้อง"}), 400
    if not court_can_access_athlete(athlete, round_no):
        return jsonify({"ok": False, "message": "ไม่มีสิทธิ์สนามนี้"}), 403
    page_token = str(payload.get("page_token", ""))
    if str(payload.get("action", "")) == "release":
        release_scorecard_lock(athlete.id, round_no, page_token)
        return jsonify({"ok": True, "released": True})
    if str(payload.get("action", "")) == "check":
        # เปิดหน้าดูเฉย ๆ ไม่ยึดล็อก: บอกแค่ว่ามีเครื่องอื่นกำลังคีย์อยู่หรือไม่
        from datetime import timedelta
        lock = ScorecardLock.query.filter_by(athlete_id=athlete.id, round_no=round_no).first()
        busy = bool(lock and lock.page_token != page_token and lock.heartbeat_at
                    and datetime.utcnow() - lock.heartbeat_at < timedelta(seconds=SCORECARD_LOCK_TTL_SECONDS))
        return jsonify({"ok": not busy, "locked_by": lock.username if busy else None})
    take_over = str(payload.get("take_over", "")).lower() in {"1", "true", "yes"}
    got_lock, holder = acquire_scorecard_lock(athlete.id, round_no, page_token, take_over=take_over)
    return jsonify({"ok": got_lock, "locked_by": holder})


@app.route("/athletes/<int:athlete_id>/approve-score", methods=["POST"])
@login_required
@role_required("superadmin")
def approve_score(athlete_id: int):
    """Superadmin รับรองผลที่ยิงครบแล้วแต่ไม่มีลายเซ็นครบ 3 ฝ่าย."""
    athlete = Athlete.query.get_or_404(athlete_id)
    event = athlete.event
    try:
        round_no = int(request.form.get("round", request.args.get("round", 1)))
    except (TypeError, ValueError):
        round_no = 1

    summary = summarize_round(athlete.id, round_no)
    if not overview_score_complete(summary):
        flash("ยังยิงไม่ครบทุกสถานี จึงยัง Approve ไม่ได้", "warning")
        if request.form.get("next") == "scorecard":
            return redirect(url_for("scorecard", athlete_id=athlete.id, round=round_no))
        return redirect(url_for("event_overview", event_id=event.id, round=round_no))

    signature = ensure_signature(athlete.id, round_no)
    # bypass_signed มีความหมายใหม่: รับรองด้วย Superadmin โดยตรง
    signature.bypass_signed = True
    signature.finished_at = signature.finished_at or datetime.utcnow()
    athlete.status = "finished"
    db.session.commit()
    clear_request_cache()

    if round_no == 1 and event.has_round_two:
        sync_round_two_candidates(event)
        reset_event_bracket(event)
    elif round_no == 2 and event.has_round_two:
        reset_event_bracket(event)

    flash(f"Approve ผลของ {athlete.name} แล้ว", "success")
    if request.form.get("next") == "scorecard":
        return redirect(url_for("scorecard", athlete_id=athlete.id, round=round_no))
    return redirect(url_for("event_overview", event_id=event.id, round=round_no))


@app.route("/athletes/<int:athlete_id>/scorecard", methods=["GET", "POST"])
@login_required
@role_required("admin", "superadmin", "court")
def scorecard(athlete_id: int):
    athlete = Athlete.query.get_or_404(athlete_id)
    event = athlete.event
    round_no = request.args.get("round", 1, type=int) or 1
    if not court_can_access_athlete(athlete, round_no):
        flash(f"บัญชี {court_public_id(current_user)} ใช้งานได้เฉพาะอีเวนต์ที่ได้รับมอบหมายและสนาม {current_user.court_no}", "warning")
        if is_court_user():
            return redirect(url_for("court_queue", event_id=current_user.court_event_id, round=1))
        return redirect(url_for("event_overview", event_id=event.id, round=round_no))
    if round_no == 2 and not is_round_two_candidate(event, athlete):
        flash("นักกีฬาคนนี้ไม่มีสิทธิ์ตีรอบ 2", "warning")
        return redirect(url_for("event_overview", event_id=event.id, round=1))

        # GET: เปิดหน้า scorecard อย่างเดียว
    # ไม่สร้าง score_entry / ไม่เปลี่ยนสถานะ / ไม่ commit
    # POST หรือ autosave ค่อยสร้างข้อมูลจริง
    if request.method == "POST":
        ensure_round_entries(athlete.id, round_no)
        signature = ensure_signature(athlete.id, round_no)
    else:
        signature = ScoreSignature.query.filter_by(
            athlete_id=athlete.id,
            round_no=round_no
        ).first()

        if not signature:
            signature = ScoreSignature(
                athlete_id=athlete.id,
                round_no=round_no
            )

    if request.method == "POST":
        referee_name_key = f"referee_name_{round_no}"
        recorder_name_key = f"recorder_name_{round_no}"
        athlete_name_key = f"athlete_name_{round_no}"
        referee_sig_key = f"ref_sig_{round_no}"
        recorder_sig_key = f"rec_sig_{round_no}"
        athlete_sig_key = f"ath_sig_{round_no}"

        if not signature.started_at:
            signature.started_at = datetime.utcnow()
            athlete.status = "active"

        signature.recorder_name = request.form.get(recorder_name_key, "").strip() or signature.recorder_name
        signature.referee_name = request.form.get(referee_name_key, "").strip() or signature.referee_name
        signature.athlete_name = request.form.get(athlete_name_key, "").strip() or signature.athlete_name

        signature.recorder_signature = request.form.get(recorder_sig_key, "").strip() or signature.recorder_signature
        signature.referee_signature = request.form.get(referee_sig_key, "").strip() or signature.referee_signature
        signature.athlete_signature = request.form.get(athlete_sig_key, "").strip() or signature.athlete_signature

        bypass_code = request.form.get("bypass_code", "").strip()
        # การข้ามลายเซ็นใช้เพื่อ “จบ/ส่งผล” เท่านั้น ไม่ถือว่า APPROVED
        # APPROVED อัตโนมัติเกิดเฉพาะเมื่อมีลายเซ็นจริงครบ 3 ฝ่าย
        # หรือ Superadmin กด Approve จากหน้า Overview ภายหลัง
        expected_bypass = os.environ.get("SIGNATURE_BYPASS_CODE", "7929")
        bypass_ok = current_user.role == "superadmin" or (
            current_user.role == "admin" and bool(expected_bypass) and bypass_code == expected_bypass
        )
        signed_ok = all([
            bool(signature.recorder_signature),
            bool(signature.referee_signature),
            bool(signature.athlete_signature),
        ])

        if bypass_ok or signed_ok:
            # ห้ามตั้ง bypass_signed จากการส่ง Scorecard เพราะจะทำให้แถวฟ้าทันที
            # ค่านี้สงวนไว้สำหรับการกด Approve โดย Superadmin เท่านั้น
            signature.finished_at = datetime.utcnow()
            athlete.status = "finished"
            db.session.commit()

            if round_no == 1 and event.has_round_two:
                sync_round_two_candidates(event)
                reset_event_bracket(event)
            elif round_no == 2 and event.has_round_two:
                reset_event_bracket(event)
            elif round_no >= 3:
                round_map = (
                    {3: "R16", 4: "QF", 5: "SF", 6: "F"}
                    if event_has_round_of_16(event)
                    else {3: "QF", 4: "SF", 5: "F"}
                )

                if round_no not in round_map:
                    flash("รอบแข่งขันไม่ถูกต้อง", "warning")
                    return redirect(url_for("bracket", event_id=event.id))

                match = BracketMatch.query.filter_by(
                    event_id=event.id,
                    round_name=round_map[round_no]
                ).filter(
                    (BracketMatch.athlete_a_id == athlete.id) |
                    (BracketMatch.athlete_b_id == athlete.id)
                ).first()

                if match:
                    sync_match_winner_from_scores(match)
                    maybe_advance_bracket(event)

                flash("จบการตีเรียบร้อย", "success")
                return redirect(url_for("bracket", event_id=event.id))

            flash("จบการตีเรียบร้อย", "success")
            if is_court_user() and round_no in (1, 2):
                return redirect(url_for("court_queue", event_id=event.id, round=round_no))
            return redirect(url_for("event_overview", event_id=event.id, round=round_no))

        flash("ต้องลงชื่ออย่างใดอย่างหนึ่ง (พิมพ์ชื่อหรือเขียน) ให้ครบทั้ง 3 ฝ่าย หรือใช้สิทธิ์ข้าม", "danger")
        return redirect(url_for("scorecard", athlete_id=athlete.id, round=round_no))

    template_data = build_scorecard_template_data(athlete.id)
    ranks = compute_round_ranks(event)
    template_data["round_ranks"] = {
        1: ranks.get(1, {}).get(athlete.id, ""),
        2: ranks.get(2, {}).get(athlete.id, ""),
        3: "",
        4: "",
        5: "",
        6: "",
    }

    round_station_running_totals = {}
    for rn in scorecard_round_numbers(event):
        running = {}
        acc = 0
        for st in [1, 2, 3, 4, 5]:
            val = template_data["station_totals"].get((rn, st), 0)
            acc += val
            running[st] = acc
        round_station_running_totals[rn] = running

    round_signatures = {}
    for rn in scorecard_round_numbers(event):
        round_signatures[rn] = get_round_signature(athlete.id, rn)

    combined_rows = build_combined_qualifiers(event) if event.has_round_two else []
    current_combined = next(
        (
            r for r in combined_rows
            if r.get("athlete") and r["athlete"].id == athlete.id
        ),
        None
    )

    display_order = athlete.start_order
    display_lane_no = athlete.lane_no
    display_lane_order = athlete.lane_order

    current_round_rows = build_round_ranking(event, round_no)
    current_row = next(
        (row for row in current_round_rows if row["athlete"].id == athlete.id),
        None
    )

    if current_row:
        display_order = current_row.get("display_order", display_order)
        display_lane_no = current_row.get("display_lane_no", display_lane_no)
        display_lane_order = current_row.get("display_lane_order", display_lane_order)

    return render_template(
        "scorecard.html",
        athlete=athlete,
        event=event,
        round_no=round_no,
        round_labels=scorecard_round_labels(event),
        score_map=template_data["score_map"],
        station_totals=template_data["station_totals"],
        station_reds=template_data["station_reds"],
        round_totals=template_data["round_totals"],
        round_ranks=template_data["round_ranks"],
        signature=signature,
        round_signatures=round_signatures,
        round_station_running_totals=round_station_running_totals,
        is_superadmin=(current_user.role == "superadmin"),
        theme=event_theme(event.category),
        combined_rows=combined_rows,
        current_combined=current_combined,
        station_images=[f"station_{i}.png" for i in STATIONS],
        display_order=display_order,
        display_lane_no=display_lane_no,
        display_lane_order=display_lane_order,
        score_edit_count=ScoreEditLog.query.filter_by(athlete_id=athlete.id, round_no=round_no).count(),
        current_round_score_complete=overview_score_complete(summarize_round(athlete.id, round_no)),
        current_round_approved=athlete_round_is_approved(athlete, round_no),
        is_finalized=bool(signature and signature.finished_at),
        current_round_red_cards=sum(
            template_data["station_reds"].get((round_no, station_no), 0)
            for station_no in STATIONS
        ),
        round_stopped_by_red=bool(
            getattr(signature, "stopped_by_red", False)
            or sum(
                template_data["station_reds"].get((round_no, station_no), 0)
                for station_no in STATIONS
            ) >= MAX_RED_CARDS
        ),
    )


@app.route("/athletes/<int:athlete_id>/score-history")
@login_required
@role_required("admin", "superadmin", "court")
def score_history(athlete_id: int):
    athlete = Athlete.query.get_or_404(athlete_id)
    round_no = request.args.get("round", 1, type=int) or 1
    if not court_can_access_athlete(athlete, round_no):
        flash("ไม่มีสิทธิ์ดูประวัติคะแนนของอีเวนต์หรือสนามอื่น", "warning")
        if is_court_user():
            return redirect(url_for("court_queue", event_id=current_user.court_event_id, round=1))
        return redirect(url_for("index"))
    logs = ScoreEditLog.query.filter_by(athlete_id=athlete.id, round_no=round_no).order_by(ScoreEditLog.edited_at.desc()).all()
    return render_template("score_history.html", athlete=athlete, event=athlete.event, round_no=round_no, logs=logs)


@app.route("/events/<int:event_id>/scorecards-print-select", methods=["GET", "POST"])
@login_required
@role_required("admin", "superadmin")
def scorecards_print_select(event_id: int):
    event = Event.query.get_or_404(event_id)
    round_no = int(request.values.get("round", 1))

    athletes = athletes_for_scorecard_round(event, round_no)

    if request.method == "POST":
        print_mode = request.form.get("print_mode", "selected")
        selected_ids = request.form.getlist("athlete_ids")

        if print_mode == "all":
            return redirect(url_for(
                "scorecards_print_bulk",
                event_id=event.id,
                round=round_no
            ))

        if not selected_ids:
            flash("กรุณาเลือกนักกีฬาอย่างน้อย 1 คน", "warning")
            return redirect(url_for(
                "scorecards_print_select",
                event_id=event.id,
                round=round_no
            ))

        ids_text = ",".join(selected_ids)
        return redirect(url_for(
            "scorecards_print_bulk",
            event_id=event.id,
            round=round_no,
            ids=ids_text
        ))

    return render_template(
        "scorecards_print_select.html",
        event=event,
        athletes=athletes,
        round_no=round_no,
        round_labels=scorecard_round_labels(event),
    )


@app.route("/events/<int:event_id>/scorecards-print-bulk")
@login_required
@role_required("admin", "superadmin")
def scorecards_print_bulk(event_id: int):
    event = Event.query.get_or_404(event_id)
    round_no = request.args.get("round", 1, type=int) or 1

    ids_text = request.args.get("ids", "").strip()
    selected_ids = []
    if ids_text:
        for raw_id in ids_text.split(","):
            raw_id = raw_id.strip()
            if raw_id.isdigit():
                selected_ids.append(int(raw_id))

    athletes = athletes_for_scorecard_round(event, round_no, selected_ids if selected_ids else None)

    print_items = [
        build_scorecard_print_context(athlete, round_no)
        for athlete in athletes
    ]

    return render_template(
        "scorecard_print.html",
        event=event,
        round_no=round_no,
        round_labels=scorecard_round_labels(event),
        print_items=print_items,
        is_bulk=True,
        station_images=[f"station_{i}.png" for i in STATIONS],
    )

@app.route("/athletes/<int:athlete_id>/scorecard-print")
@login_required
@role_required("admin", "superadmin")
def scorecard_print(athlete_id: int):
    athlete = Athlete.query.get_or_404(athlete_id)
    event = athlete.event
    round_no = request.args.get("round", 1, type=int) or 1

    if round_no == 2 and event.has_round_two and not is_round_two_candidate(event, athlete):
        flash("นักกีฬาคนนี้ไม่มีสิทธิ์ตีรอบ 2", "warning")
        return redirect(url_for("event_overview", event_id=event.id, round=1))

    print_item = build_scorecard_print_context(athlete, round_no)

    return render_template(
        "scorecard_print.html",
        event=event,
        round_no=round_no,
        round_labels=scorecard_round_labels(event),
        print_items=[print_item],
        is_bulk=False,
        station_images=[f"station_{i}.png" for i in STATIONS],
    )

@app.route("/athletes/<int:athlete_id>/activate", methods=["POST"])
@login_required
@role_required("admin", "superadmin")
def activate_scorecard(athlete_id: int):
    athlete = Athlete.query.get_or_404(athlete_id)
    round_no = request.args.get("round", 1, type=int) or 1
    signature = ensure_signature(athlete.id, round_no)
    if not signature.finished_at:
        if not signature.started_at:
            signature.started_at = datetime.utcnow()
        athlete.status = "active"
        db.session.commit()
    return jsonify({"ok": True, "status": athlete_round_status(athlete, round_no)})


@app.route("/athletes/<int:athlete_id>/tiebreak", methods=["GET", "POST"])
@login_required
@role_required("admin", "superadmin")
def tiebreak(athlete_id: int):
    athlete = Athlete.query.get_or_404(athlete_id)
    round_no = int(request.values.get("round", request.args.get("round", 1)) or 1)
    if request.method == "POST":
        # ไม่ลบของเดิม: หากยังเท่ากัน ให้กลับมาตีเที่ยวพิเศษเพิ่มได้เรื่อย ๆ
        for station_no in STATIONS:
            score = int(request.form.get(f"tb_{station_no}", 0) or 0)
            db.session.add(TieBreakEntry(athlete_id=athlete.id, round_no=round_no, station_no=station_no, score=score))
        db.session.commit()
        clear_request_cache()
        flash("บันทึกผลเที่ยวพิเศษแล้ว ถ้ายังเท่ากันให้บันทึกเที่ยวพิเศษเพิ่มอีกครั้ง", "success")
        return redirect(url_for("event_overview", event_id=athlete.event_id, round=round_no))
    entries = TieBreakEntry.query.filter_by(athlete_id=athlete.id, round_no=round_no).all()
    existing = {station: sum(e.score for e in entries if e.station_no == station) for station in STATIONS}
    return render_template("tiebreak.html", athlete=athlete, round_no=round_no, existing=existing, athletes=[athlete], event=athlete.event)


@app.route("/events/<int:event_id>/tiebreak", methods=["GET", "POST"])
@login_required
@role_required("admin", "superadmin")
def event_tiebreak(event_id: int):
    event = Event.query.get_or_404(event_id)
    round_no = int(request.values.get("round", request.args.get("round", 1)) or 1)
    raw_ids = request.values.get("ids", "")
    athlete_ids = []
    for part in raw_ids.replace(" ", "").split(','):
        if part.isdigit():
            athlete_ids.append(int(part))
    athletes = Athlete.query.filter(Athlete.event_id == event.id, Athlete.id.in_(athlete_ids)).order_by(Athlete.start_order, Athlete.id).all() if athlete_ids else []
    if not athletes:
        flash("กรุณาเลือกนักกีฬาที่ต้องตี Shoot-off อย่างน้อย 2 คน", "warning")
        return redirect(url_for("event_overview", event_id=event.id, round=round_no))
    if request.method == "POST":
        entries_data = []
        for athlete in athletes:
            for station_no in STATIONS:
                score = int(request.form.get(f"tb_{athlete.id}_{station_no}", 0) or 0)
                entries_data.append({
                    "athlete_id": athlete.id,
                    "round_no": round_no,
                    "station_no": station_no,
                    "score": score,
                })

        try:
            for item in entries_data:
                db.session.add(TieBreakEntry(**item))
            db.session.commit()
        except Exception as exc:
            # PostgreSQL บางฐานที่ย้ายมาจาก SQLite มี tie_break_entry.id เป็น NOT NULL
            # แต่ไม่มี default sequence ทำให้ INSERT แล้ว id เป็น null
            db.session.rollback()
            msg = str(exc).lower()
            if "tie_break_entry" in msg and ("null value in column" in msg or "not-null constraint" in msg):
                next_id = next_manual_id(TieBreakEntry)
                if next_id is None:
                    raise
                for item in entries_data:
                    item["id"] = next_id
                    next_id += 1
                    db.session.add(TieBreakEntry(**item))
                db.session.commit()
            else:
                raise

        clear_request_cache()
        flash("บันทึก Shoot-off พร้อมกันแล้ว ถ้ายังเท่ากันให้เลือกกลุ่มเดิมแล้วบันทึกเพิ่ม", "success")
        return redirect(url_for("event_overview", event_id=event.id, round=round_no))
    existing = {}
    for athlete in athletes:
        entries = TieBreakEntry.query.filter_by(athlete_id=athlete.id, round_no=round_no).all()
        existing[athlete.id] = {station: sum(e.score for e in entries if e.station_no == station) for station in STATIONS}
    return render_template("tiebreak.html", event=event, athletes=athletes, athlete=athletes[0], round_no=round_no, existing=existing, bulk_mode=True, ids=','.join(str(a.id) for a in athletes))


@app.route("/events/<int:event_id>/bracket")
def bracket(event_id: int):
    event = Event.query.get_or_404(event_id)
    qualifiers = build_combined_qualifiers(event)
    ensure_bracket(event)
    maybe_advance_bracket(event)
    matches = BracketMatch.query.filter_by(event_id=event.id).order_by(BracketMatch.round_name, BracketMatch.match_no).all()
    athlete_map = {a.id: a for a in event.athletes}
    seed_map = {row["athlete"].id: row.get("seed") for row in qualifiers if row.get("athlete") and row.get("seed")}
    grouped = {"R16": [], "QF": [], "SF": [], "F": []}
    for m in matches:
        a = athlete_map.get(m.athlete_a_id)
        b = athlete_map.get(m.athlete_b_id)
        grouped.setdefault(m.round_name, []).append({
            "match": m,
            "a": build_bracket_match_row(event, a, m.round_name, seed_map),
            "b": build_bracket_match_row(event, b, m.round_name, seed_map),
            "winner": athlete_map.get(m.winner_id),
            "round_no": bracket_round_to_scorecard_round(m.round_name, event),
            "status": bracket_match_status(event, m),
        })
    combined_rows = build_combined_qualifiers(event) if event.has_round_two else qualifiers
    start_round = configured_bracket_start_round(event)
    return render_template("bracket.html", event=event, grouped=grouped, combined_rows=combined_rows, start_round=start_round)


@app.route("/matches/<int:match_id>/winner", methods=["POST"])
@login_required
@role_required("admin", "superadmin")
def set_match_winner(match_id: int):
    match = BracketMatch.query.get_or_404(match_id)
    winner_id = int(request.form.get("winner_id"))
    if winner_id not in {match.athlete_a_id, match.athlete_b_id}:
        flash("ผู้ชนะไม่ถูกต้อง", "danger")
        return redirect(url_for("bracket", event_id=match.event_id))
    match.winner_id = winner_id
    db.session.commit()
    maybe_advance_bracket(Event.query.get(match.event_id))
    flash("บันทึกผู้ชนะแล้ว", "success")
    return redirect(url_for("bracket", event_id=match.event_id))


@app.route("/events/<int:event_id>/bracket.xlsx")
def bracket_excel(event_id: int):
    event = Event.query.get_or_404(event_id)
    qualifiers = build_combined_qualifiers(event)
    wb = Workbook()
    ws = wb.active
    ws.title = "Bracket"
    ws.append(["Seed","Name","Affiliation","Round1","Round2","Sum"])
    for idx, row in enumerate(qualifiers, start=1):
        ws.append([idx, row["athlete"].name, row["athlete"].affiliation, row.get("round1_total", row["total"]), row.get("round2_total", ""), row["total"]])
    stream = BytesIO(); wb.save(stream); stream.seek(0)
    return send_file(stream, as_attachment=True, download_name=f"event_{event.id}_bracket.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

@app.route("/events/<int:event_id>/bracket_data")
@shared_poll_cache
def bracket_data(event_id: int):
    event = Event.query.get_or_404(event_id)
    preload_event_score_data(event)

    maybe_advance_bracket(event)

    matches = BracketMatch.query.filter_by(event_id=event.id).order_by(
        BracketMatch.round_name,
        BracketMatch.match_no
    ).all()

    athlete_map = {a.id: a for a in event.athletes}
    qualifiers = build_combined_qualifiers(event)
    seed_map = {row["athlete"].id: row.get("seed", idx) for idx, row in enumerate(qualifiers, start=1) if row.get("athlete") and row.get("seed")}

    grouped = {"R16": [], "QF": [], "SF": [], "F": []}
    for m in matches:
        grouped.setdefault(m.round_name, []).append({
            "match_no": m.match_no,
            "round_no": bracket_round_to_scorecard_round(m.round_name, event),
            "winner_id": m.winner_id,
            "status": bracket_match_status(event, m),
            "a": build_bracket_row_data(event, athlete_map.get(m.athlete_a_id), m.round_name, seed_map),
            "b": build_bracket_row_data(event, athlete_map.get(m.athlete_b_id), m.round_name, seed_map),
        })

    return jsonify(grouped)


# =========================

# Results Approved report (SEA Games style)
# =========================
def _ra_text(value, default="") -> str:
    if value is None:
        return default
    return str(value)


def _ra_event_title(event: Event) -> str:
    raw = " ".join([event.event_group or "", event.category or "", event.name or ""]).strip()
    return (raw or event.name or "PÉTANQUE SHOOTING").upper()


def _ra_date_text(event: Event) -> str:
    try:
        return event.competition_date.strftime("%d %B %Y").upper()
    except Exception:
        return _ra_text(event.competition_date, "")


def _ra_default_setting(event: Event) -> SimpleNamespace:
    return SimpleNamespace(
        event_id=event.id,
        competition_title=(event.name or "33rd SEA GAMES THAILAND 2025").upper(),
        host_line=(event.location or "THAILAND : THE KINGDOM OF THAILAND").upper(),
        date_line=_ra_date_text(event),
        location_line=(event.location or "").upper(),
        country_label="COUNTRY",
        president_title="PRESIDENT OF WORLD PÉTANQUE AND BOULES FEDERATION",
        president_name="",
        technical_title="F.I.P.J.P VICE PRESIDENT OF\nPÉTANQUE TECHNICAL DELEGATE",
        technical_name="",
        umpires_text="",
        approved_text="……………………………APPROVED",
        show_official_pages=True,
        cover_main_logo_path=None,
        cover_bottom_logo_1_path=None,
        cover_bottom_logo_2_path=None,
        cover_bottom_logo_3_path=None,
        header_logo_1_path=None,
        header_logo_2_path=None,
        header_logo_3_path=None,
        header_logo_4_path=None,
        side_logo_path=None,
    )


def get_results_approved_setting(event: Event, create: bool = False):
    setting = ResultsApprovedSetting.query.filter_by(event_id=event.id).first()
    if setting:
        return setting
    default = _ra_default_setting(event)
    if not create:
        return default
    setting = ResultsApprovedSetting(
        event_id=event.id,
        competition_title=default.competition_title,
        host_line=default.host_line,
        date_line=default.date_line,
        location_line=default.location_line,
        country_label=default.country_label,
        president_title=default.president_title,
        president_name=default.president_name,
        technical_title=default.technical_title,
        technical_name=default.technical_name,
        umpires_text=default.umpires_text,
        approved_text=default.approved_text,
        show_official_pages=default.show_official_pages,
        cover_main_logo_path=default.cover_main_logo_path,
        cover_bottom_logo_1_path=default.cover_bottom_logo_1_path,
        cover_bottom_logo_2_path=default.cover_bottom_logo_2_path,
        cover_bottom_logo_3_path=default.cover_bottom_logo_3_path,
        header_logo_1_path=default.header_logo_1_path,
        header_logo_2_path=default.header_logo_2_path,
        header_logo_3_path=default.header_logo_3_path,
        header_logo_4_path=default.header_logo_4_path,
        side_logo_path=default.side_logo_path,
    )
    db.session.add(setting)
    db.session.commit()
    return setting


def _ra_setting_text(setting, attr: str, default: str = "") -> str:
    value = getattr(setting, attr, None)
    return _ra_text(value, default).strip() or default


RESULTS_APPROVED_LOGO_FIELDS = {
    "cover_main_logo": "cover_main_logo_path",
    "cover_bottom_logo_1": "cover_bottom_logo_1_path",
    "cover_bottom_logo_2": "cover_bottom_logo_2_path",
    "cover_bottom_logo_3": "cover_bottom_logo_3_path",
    "header_logo_1": "header_logo_1_path",
    "header_logo_2": "header_logo_2_path",
    "header_logo_3": "header_logo_3_path",
    "header_logo_4": "header_logo_4_path",
    "side_logo": "side_logo_path",
}

RESULTS_APPROVED_LOGO_DEFAULTS = {
    "cover_main": "results_approved_assets/thailand2025.png",
    "cover_bottom_1": "results_approved_assets/fipjp.png",
    "cover_bottom_2": "results_approved_assets/wpbf.png",
    "cover_bottom_3": "results_approved_assets/absc.png",
    "header_1": "results_approved_assets/absc.png",
    "header_2": "results_approved_assets/thailand2025.png",
    "header_3": "results_approved_assets/wpbf.png",
    "header_4": "results_approved_assets/fipjp.png",
    "side": "results_approved_assets/thailand2025.png",
}

ALLOWED_RESULTS_LOGO_EXTENSIONS = {"png", "jpg", "jpeg", "webp", "gif"}


def _ra_logo_value(setting, attr: str, default: str) -> str:
    value = getattr(setting, attr, None)
    value = _ra_text(value).strip()
    return value or default


def _ra_logo_map(setting) -> dict:
    return {
        "cover_main": _ra_logo_value(setting, "cover_main_logo_path", RESULTS_APPROVED_LOGO_DEFAULTS["cover_main"]),
        "cover_bottom_1": _ra_logo_value(setting, "cover_bottom_logo_1_path", RESULTS_APPROVED_LOGO_DEFAULTS["cover_bottom_1"]),
        "cover_bottom_2": _ra_logo_value(setting, "cover_bottom_logo_2_path", RESULTS_APPROVED_LOGO_DEFAULTS["cover_bottom_2"]),
        "cover_bottom_3": _ra_logo_value(setting, "cover_bottom_logo_3_path", RESULTS_APPROVED_LOGO_DEFAULTS["cover_bottom_3"]),
        "header_1": _ra_logo_value(setting, "header_logo_1_path", RESULTS_APPROVED_LOGO_DEFAULTS["header_1"]),
        "header_2": _ra_logo_value(setting, "header_logo_2_path", RESULTS_APPROVED_LOGO_DEFAULTS["header_2"]),
        "header_3": _ra_logo_value(setting, "header_logo_3_path", RESULTS_APPROVED_LOGO_DEFAULTS["header_3"]),
        "header_4": _ra_logo_value(setting, "header_logo_4_path", RESULTS_APPROVED_LOGO_DEFAULTS["header_4"]),
        "side": _ra_logo_value(setting, "side_logo_path", RESULTS_APPROVED_LOGO_DEFAULTS["side"]),
    }


def _ra_save_uploaded_logo(event_id: int, field_name: str) -> str | None:
    uploaded = request.files.get(field_name)
    if not uploaded or not uploaded.filename:
        return None
    original = secure_filename(uploaded.filename)
    ext = original.rsplit(".", 1)[-1].lower() if "." in original else ""
    if ext not in ALLOWED_RESULTS_LOGO_EXTENSIONS:
        flash(f"ไฟล์โลโก้ {field_name} ต้องเป็น png, jpg, jpeg, webp หรือ gif", "danger")
        return None
    upload_dir = os.path.join(BASE_DIR, "static", "uploads", "results_approved", f"event_{event_id}")
    os.makedirs(upload_dir, exist_ok=True)
    filename = f"{field_name}_{datetime.utcnow().strftime('%Y%m%d%H%M%S%f')}.{ext}"
    uploaded.save(os.path.join(upload_dir, filename))
    return f"uploads/results_approved/event_{event_id}/{filename}"


def _ra_static_abs_path(static_filename: str | None) -> str | None:
    static_filename = _ra_text(static_filename).strip()
    if not static_filename:
        return None
    path = os.path.join(BASE_DIR, "static", *static_filename.split("/"))
    return path if os.path.exists(path) else None


def _ra_docx_add_center_image(doc, static_filename: str | None, width_inches: float = 1.35):
    path = _ra_static_abs_path(static_filename)
    if not path:
        return None
    try:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Inches
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(path, width=Inches(width_inches))
        return p
    except Exception:
        return None


def _ra_docx_add_logo_row(doc, static_filenames: list[str], width_inches: float = 0.75):
    paths = [_ra_static_abs_path(x) for x in static_filenames if x]
    paths = [p for p in paths if p]
    if not paths:
        return None
    try:
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Inches
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for idx, path in enumerate(paths):
            run = p.add_run()
            run.add_picture(path, width=Inches(width_inches))
            if idx != len(paths) - 1:
                p.add_run("     ")
        return p
    except Exception:
        return None


def _ra_split_name(full_name: str, mode: str = "first") -> tuple[str, str]:
    """แยกชื่อ-นามสกุลตามรูปแบบที่ตั้งไว้ในหน้าตั้งค่า Results
    first = คำแรกเป็นนามสกุล (แบบสากลในเล่มผล เช่น NGUYEN | THI HIEN)
    last  = คำสุดท้ายเป็นนามสกุล (แบบไทย เช่น สมชาย | ใจดี -> ใจดี | สมชาย)
    อย่างอื่น = ไม่แยก"""
    parts = [p for p in _ra_text(full_name).strip().split() if p]
    if len(parts) < 2 or mode not in {"first", "last"}:
        return "", " ".join(parts).upper()
    if mode == "first":
        return parts[0].upper(), " ".join(parts[1:]).upper()
    return parts[-1].upper(), " ".join(parts[:-1]).upper()


def _ra_umpire_rows(setting) -> list[dict]:
    text = _ra_setting_text(setting, "umpires_text", "")
    rows = []
    for idx, line in enumerate([ln.strip() for ln in text.splitlines() if ln.strip()], start=1):
        parts = [p.strip() for p in line.split("|")]
        if len(parts) >= 2:
            name, federation = parts[0], parts[1]
        else:
            name, federation = line, ""
        rows.append({"no": idx, "name": name.upper(), "federation": federation.upper()})
    return rows


def _ra_station_cells(summary: dict) -> list[int]:
    cells = []
    by_station = summary.get("by_station", {}) if summary else {}
    for station_no in STATIONS:
        cells.append(int(by_station.get(station_no, {}).get("total", 0) or 0))
    return cells


def _ra_distance_cells(summary: dict) -> list[int]:
    cells = []
    by_station = summary.get("by_station", {}) if summary else {}
    for station_no in STATIONS:
        distances = by_station.get(station_no, {}).get("distances", {})
        for distance_m in DISTANCES:
            cells.append(int(distances.get(distance_m, 0) or 0))
        cells.append(int(by_station.get(station_no, {}).get("total", 0) or 0))
    return cells


def _ra_station_groups(summary: dict) -> list[dict]:
    groups = []
    by_station = summary.get("by_station", {}) if summary else {}
    for station_no in STATIONS:
        station = by_station.get(station_no, {})
        distances = station.get("distances", {}) if station else {}
        values = [int(distances.get(distance_m, 0) or 0) for distance_m in DISTANCES]
        groups.append({
            "station": station_no,
            "values": values,
            "total": int(station.get("total", 0) or 0),
        })
    return groups


def _ra_athlete_rank(row: dict | None) -> str:
    if not row:
        return ""
    return _ra_text(row.get("display_rank") or row.get("ordinal_rank") or row.get("rank") or "")


def _ra_loser(match: BracketMatch):
    if not match or not match.winner_id:
        return None
    if match.athlete_a_id == match.winner_id:
        return Athlete.query.get(match.athlete_b_id) if match.athlete_b_id else None
    if match.athlete_b_id == match.winner_id:
        return Athlete.query.get(match.athlete_a_id) if match.athlete_a_id else None
    return None


def _ra_match_score(event: Event, match: BracketMatch, athlete_id: int | None) -> int | str:
    if not athlete_id:
        return ""
    round_no = bracket_round_to_scorecard_round(match.round_name, event)
    return summarize_round(athlete_id, round_no).get("total", 0)


def _ra_build_medal_rows(event: Event, bracket_matches: list[BracketMatch], fallback_rows: list[dict]) -> list[dict]:
    final = next((m for m in bracket_matches if m.round_name == "F"), None)
    semis = [m for m in bracket_matches if m.round_name == "SF"]
    medals = []

    if final and final.winner_id:
        gold = Athlete.query.get(final.winner_id)
        silver = _ra_loser(final)
        if gold:
            medals.append({"medal": "GOLD", "athlete": gold})
        if silver:
            medals.append({"medal": "SILVER", "athlete": silver})
        for semi in sorted(semis, key=lambda m: m.match_no):
            bronze = _ra_loser(semi)
            if bronze:
                medals.append({"medal": "BRONZE", "athlete": bronze})

    if medals:
        return medals[:4]

    labels = ["GOLD", "SILVER", "BRONZE", "BRONZE"]
    fallback = []
    for idx, row in enumerate(fallback_rows[:4]):
        athlete = row.get("athlete") if isinstance(row, dict) else None
        if athlete:
            fallback.append({"medal": labels[idx] if idx < len(labels) else "", "athlete": athlete})
    return fallback


def build_results_approved_context(event: Event) -> dict:
    """เตรียมข้อมูลรายงาน Results Approved SEA Games style จากข้อมูลจริงในระบบ"""
    preload_event_score_data(event)
    setting = get_results_approved_setting(event)
    country_label = _ra_setting_text(setting, "country_label", "COUNTRY").upper()
    approved_text = _ra_setting_text(setting, "approved_text", "……………………………APPROVED")
    athletes = sorted(event.athletes, key=lambda a: (a.start_order or 999999, a.id))

    round1_ranking = build_round_ranking(event, 1)
    round1_by_id = {row["athlete"].id: row for row in round1_ranking}

    entry_countries = []
    seen_country = set()
    for athlete in athletes:
        country = (athlete.affiliation or "").upper()
        key = country.strip().lower()
        if key and key not in seen_country:
            seen_country.add(key)
            entry_countries.append({"no": len(entry_countries) + 1, "country": country})

    name_format = (getattr(setting, "name_format", None) or "single").strip().lower()
    name_rows = []
    qf1_rows = []
    qf1_detail_rows = []
    for athlete in athletes:
        family, given = _ra_split_name(athlete.name, name_format)
        name_rows.append({
            "no": athlete.start_order,
            "country": (athlete.affiliation or "").upper(),
            "family_name": family,
            "given_name": given,
            "name": (athlete.name or "").upper(),
        })
        row = round1_by_id.get(athlete.id)
        summary = summarize_round(athlete.id, 1)
        rank = _ra_athlete_rank(row)
        qf1_rows.append({
            "no": athlete.start_order,
            "country": (athlete.affiliation or "").upper(),
            "name": (athlete.name or "").upper(),
            "lane": athlete.lane_no,
            "points": summary.get("total", 0),
            "rank": rank,
        })
        qf1_detail_rows.append({
            "rank": rank,
            "country": (athlete.affiliation or "").upper(),
            "name": (athlete.name or "").upper(),
            "station_groups": _ra_station_groups(summary),
            "stations": _ra_station_cells(summary),
            "distance_cells": _ra_distance_cells(summary),
            "total": summary.get("total", 0),
        })

    qf2_rows = []
    qf2_detail_rows = []
    direct_rows = []
    if event.has_round_two:
        overview_r2 = build_round_two_overview_rows(event)
        direct_rows = [row for row in overview_r2 if row.get("is_round2_direct_placeholder")]
        round2_rows = [row for row in overview_r2 if not row.get("is_round2_direct_placeholder")]
        round2_rows = sorted(round2_rows, key=lambda row: (row.get("display_order") if isinstance(row.get("display_order"), int) else 999999, row["athlete"].id))
        for row in round2_rows:
            athlete = row["athlete"]
            r2_summary = summarize_round(athlete.id, 2)
            qf2_rows.append({
                "qf1_rank": row.get("round1_rank") or _ra_athlete_rank(round1_by_id.get(athlete.id)),
                "country": (athlete.affiliation or "").upper(),
                "name": (athlete.name or "").upper(),
                "lane": row.get("display_lane_no", ""),
                "r1": row.get("round1_total", 0),
                "r2": row.get("round2_total", 0),
                "total": row.get("combined_total", 0),
                "qf2_rank": _ra_athlete_rank(row),
            })
            qf2_detail_rows.append({
                "qf1_rank": row.get("round1_rank") or _ra_athlete_rank(round1_by_id.get(athlete.id)),
                "country": (athlete.affiliation or "").upper(),
                "name": (athlete.name or "").upper(),
                "station_groups": _ra_station_groups(r2_summary),
                "stations": _ra_station_cells(r2_summary),
                "distance_cells": _ra_distance_cells(r2_summary),
                "r1": row.get("round1_total", 0),
                "r2": row.get("round2_total", 0),
                "total": row.get("combined_total", 0),
                "qf2_rank": _ra_athlete_rank(row),
            })

    bracket_matches = BracketMatch.query.filter_by(event_id=event.id).order_by(BracketMatch.round_name, BracketMatch.match_no).all()
    bracket_round_order = {"R16": 1, "QF": 2, "SF": 3, "F": 4}
    bracket_matches = sorted(bracket_matches, key=lambda m: (bracket_round_order.get(m.round_name, 99), m.match_no))
    try:
        fallback_qualifiers = build_combined_qualifiers(event)
    except Exception:
        fallback_qualifiers = round1_ranking
    seed_map = {}
    for idx, row in enumerate(fallback_qualifiers or [], start=1):
        athlete = row.get("athlete") if isinstance(row, dict) else None
        if athlete:
            seed_map[athlete.id] = idx

    bracket_rows = []
    for match in bracket_matches:
        athlete_a = Athlete.query.get(match.athlete_a_id) if match.athlete_a_id else None
        athlete_b = Athlete.query.get(match.athlete_b_id) if match.athlete_b_id else None
        bracket_rows.append({
            "round_name": match.round_name,
            "round_label": {"R16": "ROUND OF 16", "QF": "QUARTERFINAL ROUND", "SF": "SEMIFINAL ROUND", "F": "FINAL ROUND"}.get(match.round_name, match.round_name),
            "match_no": match.match_no,
            "lane": match.match_no,
            "athlete_a": athlete_a,
            "athlete_b": athlete_b,
            "rank_a": seed_map.get(match.athlete_a_id, ""),
            "rank_b": seed_map.get(match.athlete_b_id, ""),
            "country_a": (athlete_a.affiliation if athlete_a else "").upper(),
            "country_b": (athlete_b.affiliation if athlete_b else "").upper(),
            "name_a": (athlete_a.name if athlete_a else "").upper(),
            "name_b": (athlete_b.name if athlete_b else "").upper(),
            "points_a": _ra_match_score(event, match, match.athlete_a_id),
            "points_b": _ra_match_score(event, match, match.athlete_b_id),
            "winner_id": match.winner_id,
        })

    medal_rows = _ra_build_medal_rows(event, bracket_matches, fallback_qualifiers or round1_ranking)
    medal_rows_out = []
    for r in medal_rows:
        athlete = r["athlete"]
        family, given = _ra_split_name(athlete.name, name_format)
        medal_rows_out.append({
            "medal": r["medal"],
            "athlete": athlete,
            "country": (athlete.affiliation or "").upper(),
            "family_name": family,
            "given_name": given,
        })

    # ชื่อนักกีฬาในระบบเก็บเป็นช่องเดียว จะแยก FAMILY / GIVEN ได้ก็ต่อเมื่อชื่อมีมากกว่า 1 คำ
    # ถ้าทั้งอีเวนต์ไม่มีชื่อที่แยกได้ (เช่น ลงทะเบียนเป็นชื่อประเทศ/ชื่อเดียว) ให้ใช้คอลัมน์ NAME เดียว
    # จะได้ไม่มีคอลัมน์ว่างเปล่าในเอกสารรับรองผล
    use_split_names = name_format in {"first", "last"} and any(r.get("family_name") for r in name_rows)

    return {
        "event": event,
        "setting": setting,
        "use_split_names": use_split_names,
        "competition_title": _ra_setting_text(setting, "competition_title", event.name).upper(),
        "host_line": _ra_setting_text(setting, "host_line", event.location).upper(),
        "date_line": _ra_setting_text(setting, "date_line", _ra_date_text(event)).upper(),
        "location_line": _ra_setting_text(setting, "location_line", event.location).upper(),
        "country_label": country_label,
        "approved_text": approved_text,
        "event_title": _ra_event_title(event),
        "event_date_text": _ra_setting_text(setting, "date_line", _ra_date_text(event)).upper(),
        "athletes": athletes,
        "entry_countries": entry_countries,
        "name_rows": name_rows,
        "umpire_rows": _ra_umpire_rows(setting),
        "qf1_rows": qf1_rows,
        "qf1_detail_rows": qf1_detail_rows,
        "direct_rows": direct_rows,
        "qf2_rows": qf2_rows,
        "qf2_detail_rows": qf2_detail_rows,
        "bracket_rows": bracket_rows,
        "semifinal_rows": [r for r in bracket_rows if r["round_name"] == "SF"],
        "final_rows": [r for r in bracket_rows if r["round_name"] == "F"],
        "other_ko_rows": [r for r in bracket_rows if r["round_name"] not in {"SF", "F"}],
        "medal_rows": medal_rows_out,
        "logos": _ra_logo_map(setting),
        "stations": STATIONS,
        "distances": DISTANCES,
    }


def _ra_docx_set_cell_text(cell, text, bold=False, align="center", size_pt=9):
    cell.text = ""
    p = cell.paragraphs[0]
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT if align == "left" else WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(_ra_text(text))
    run.bold = bold
    run.font.name = "Times New Roman"
    try:
        from docx.shared import Pt
        run.font.size = Pt(size_pt)
    except Exception:
        pass


def _ra_docx_shade(cell, fill="D9D9D9"):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def _ra_docx_add_title(doc, text, size=18, spacing_after=4):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(_ra_text(text))
    run.bold = True
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    p.paragraph_format.space_after = Pt(spacing_after)
    return p


def _ra_docx_add_approved(doc, text="……………………………APPROVED"):
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run(text)
    run.bold = True
    run.font.name = "Times New Roman"
    run.font.size = Pt(11)
    return p


def _ra_docx_add_table(doc, headers, rows, align_left_cols: set[int] | None = None, font_size=8):
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    for i, header in enumerate(headers):
        _ra_docx_set_cell_text(table.rows[0].cells[i], header, bold=True, size_pt=font_size)
        _ra_docx_shade(table.rows[0].cells[i])
    align_left_cols = align_left_cols or set()
    for row in rows:
        cells = table.add_row().cells
        for i, value in enumerate(row):
            _ra_docx_set_cell_text(cells[i], value, align="left" if i in align_left_cols else "center", size_pt=font_size)
            cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    return table


def _ra_has_any_score(rows: list[dict], key: str = "points") -> bool:
    return any(int(r.get(key) or 0) > 0 for r in rows)


def _ra_setup_section(section, landscape: bool = False):
    """A4 ตามตัวอย่างเล่มผล: แนวตั้งขอบ 2 ซม. / แนวนอนขอบแคบลงให้ตารางรายสถานีพอดีหน้า"""
    from docx.enum.section import WD_ORIENT
    from docx.shared import Cm
    section.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
    section.page_width, section.page_height = (Cm(29.7), Cm(21.0)) if landscape else (Cm(21.0), Cm(29.7))
    section.top_margin = Cm(1.3)
    section.bottom_margin = Cm(1.3)
    section.left_margin = Cm(1.2 if landscape else 2.0)
    section.right_margin = Cm(1.2 if landscape else 2.0)
    section.header_distance = Cm(0.6)
    section.footer_distance = Cm(0.6)


def _ra_add_page_number_field(paragraph):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    run = paragraph.add_run()
    for tag, attr in (("w:fldChar", {"w:fldCharType": "begin"}), ("w:instrText", None), ("w:fldChar", {"w:fldCharType": "end"})):
        el = OxmlElement(tag)
        if attr:
            for k, v in attr.items():
                el.set(qn(k), v)
        else:
            el.set(qn("xml:space"), "preserve")
            el.text = "PAGE"
        run._r.append(el)
    return run


def _ra_setup_header_footer(doc, ctx):
    """หัวกระดาษ (โลโก้ + ชื่องาน) และเลขหน้า "n | Page" ทุกหน้า ยกเว้นหน้าปก"""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Inches, Pt, RGBColor
    section = doc.sections[0]
    section.different_first_page_header_footer = True
    header = section.header
    logos = ctx.get("logos", {})
    paths = [_ra_static_abs_path(logos.get(k)) for k in ("header_1", "header_2", "header_3", "header_4")]
    paths = [p for p in paths if p]
    p = header.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for idx, path in enumerate(paths):
        try:
            p.add_run().add_picture(path, height=Inches(0.42))
            if idx != len(paths) - 1:
                p.add_run("    ")
        except Exception:
            pass
    title = header.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(ctx["competition_title"])
    run.bold = True
    run.font.size = Pt(10)
    footer_p = section.footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _ra_add_page_number_field(footer_p)
    tail = footer_p.add_run(" | Page")
    tail.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
    for r in footer_p.runs:
        r.font.size = Pt(8)


def _ra_docx_page_title(doc, lines: list[tuple[str, int]], page_break_before: bool = False):
    first = True
    for idx, (text, size) in enumerate(lines):
        if text:
            p = _ra_docx_add_title(doc, text, size, spacing_after=2 if idx < len(lines) - 1 else 10)
            if first and page_break_before:
                # ขึ้นหน้าใหม่ที่หัวข้อ แทนการแทรกย่อหน้าตัวแบ่งหน้า (กันหน้าว่างเมื่อหน้าก่อนเต็มพอดี)
                p.paragraph_format.page_break_before = True
            first = False


def _ra_docx_approved_block(doc, text: str):
    """บรรทัดลงนามรับรองท้ายหน้า (ให้กรรมการเซ็นจริงบนกระดาษ ไม่ใส่ภาพลายเซ็นอัตโนมัติ)"""
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(36)
    p.paragraph_format.keep_together = True
    run = p.add_run(text or "……………………………APPROVED")
    run.bold = True
    run.font.size = Pt(11)
    return p


def _ra_docx_fix_widths(table, widths_cm, cell_margin_cm: float | None = None):
    """กำหนดความกว้างคอลัมน์ให้ทั้ง Word และ LibreOffice (ต้องตั้งทั้ง tblGrid, tblLayout และทุกเซลล์)"""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:type"), "dxa")
    tbl_w.set(qn("w:w"), str(int(sum(widths_cm) * 567)))
    if cell_margin_cm is not None:
        mar = tbl_pr.find(qn("w:tblCellMar"))
        if mar is None:
            mar = OxmlElement("w:tblCellMar")
            tbl_pr.append(mar)
        for side in ("left", "right"):
            el = mar.find(qn(f"w:{side}"))
            if el is None:
                el = OxmlElement(f"w:{side}")
                mar.append(el)
            el.set(qn("w:w"), str(int(cell_margin_cm * 567)))
            el.set(qn("w:type"), "dxa")
    grid = tbl.tblGrid
    for i, gc in enumerate(grid.findall(qn("w:gridCol"))):
        if i < len(widths_cm):
            gc.set(qn("w:w"), str(int(widths_cm[i] * 567)))
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            if i < len(widths_cm):
                cell.width = Cm(widths_cm[i])


def _ra_docx_table(doc, headers, rows, widths_cm=None, left_cols=None, font_size=10, header_fill="D9D9D9", row_height_cm=0.62):
    """ตารางอ่านง่าย: ตัวอักษรไม่เล็กเกินไป หัวตารางพิมพ์ซ้ำเมื่อขึ้นหน้าใหม่ และแถวไม่ถูกตัดครึ่ง"""
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm
    left_cols = left_cols or set()
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    table.autofit = widths_cm is None

    def mark_row(row, is_header=False):
        tr_pr = row._tr.get_or_add_trPr()
        cant = OxmlElement("w:cantSplit")
        tr_pr.append(cant)
        if is_header:
            tbl_header = OxmlElement("w:tblHeader")
            tr_pr.append(tbl_header)
        row.height = Cm(row_height_cm)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST

    hdr = table.rows[0]
    mark_row(hdr, is_header=True)
    for i, h in enumerate(headers):
        _ra_docx_set_cell_text(hdr.cells[i], h, bold=True, size_pt=font_size)
        _ra_docx_shade(hdr.cells[i], header_fill)
        hdr.cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for values in rows:
        row = table.add_row()
        mark_row(row)
        for i, v in enumerate(values):
            _ra_docx_set_cell_text(row.cells[i], v, align="left" if i in left_cols else "center", size_pt=font_size)
            row.cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    if widths_cm:
        _ra_docx_fix_widths(table, widths_cm)
    return table


def _ra_docx_station_detail_table(doc, ctx, detail_rows, round_label_cols):
    """ตารางคะแนนละเอียดรายสถานี (แนวนอน): Atelier 1-5 x ระยะ 6/7/8/9 + Tot."""
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm
    lead = ["RANK", ctx["country_label"], "NAME"]
    tail = [label for label, _ in round_label_cols]
    ncols = len(lead) + len(STATIONS) * 5 + len(tail)
    table = doc.add_table(rows=2, cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = "Table Grid"
    table.autofit = False
    _widths = [1.0, 2.6, 3.6] + [0.66] * (len(STATIONS) * 5) + [1.1] * len(tail)
    _scale = min(1.0, 27.2 / sum(_widths))
    _ra_docx_fix_widths(table, [w * _scale for w in _widths], cell_margin_cm=0.04)
    top, sub = table.rows
    for row in (top, sub):
        tr_pr = row._tr.get_or_add_trPr()
        tr_pr.append(OxmlElement("w:cantSplit"))
        tr_pr.append(OxmlElement("w:tblHeader"))
    for i, h in enumerate(lead):
        cell = top.cells[i].merge(sub.cells[i])
        _ra_docx_set_cell_text(cell, h, bold=True, size_pt=7)
        _ra_docx_shade(cell)
    col = len(lead)
    for st in STATIONS:
        merged = top.cells[col].merge(top.cells[col + 4])
        _ra_docx_set_cell_text(merged, f"ATELIER {st}", bold=True, size_pt=7)
        _ra_docx_shade(merged)
        for j, d in enumerate([f"{m}m" for m in DISTANCES] + ["Tot."]):
            _ra_docx_set_cell_text(sub.cells[col + j], d, bold=True, size_pt=7)
            _ra_docx_shade(sub.cells[col + j], "EDEDED" if j < 4 else "D9D9D9")
        col += 5
    for i, h in enumerate(tail):
        cell = top.cells[col + i].merge(sub.cells[col + i])
        _ra_docx_set_cell_text(cell, h, bold=True, size_pt=7)
        _ra_docx_shade(cell)
    for r in detail_rows:
        row = table.add_row()
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        values = [r.get("rank_display", ""), r["country"], r["name"], *r["distance_cells"], *[r.get(key, "") for _, key in round_label_cols]]
        for i, v in enumerate(values):
            is_station_total = len(lead) <= i < len(lead) + len(STATIONS) * 5 and (i - len(lead)) % 5 == 4
            _ra_docx_set_cell_text(row.cells[i], v, bold=is_station_total or i >= ncols - len(tail), align="left" if i in (1, 2) else "center", size_pt=7)
            row.cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if is_station_total:
                _ra_docx_shade(row.cells[i], "F2F2F2")
    # A4 แนวนอน พื้นที่พิมพ์ ~27.3 ซม.
    _ra_docx_fix_widths(table, [w * _scale for w in _widths], cell_margin_cm=0.04)
    return table


def _ra_docx_match_table(doc, ctx, match_rows, with_rank=True):
    """หนึ่งตารางต่อหนึ่งแมตช์ แบบในเล่มผลตัวอย่าง: MATCH | LANE | RANK | COUNTRY | NAME | POINTS"""
    from docx.shared import Pt
    for r in match_rows:
        headers = ["MATCH", "LANE"] + (["RANK"] if with_rank else []) + [ctx["country_label"], "NAME", "POINTS"]
        rows = []
        for side in ("a", "b"):
            rows.append([r["match_no"], r["lane"]] + ([r[f"rank_{side}"]] if with_rank else []) + [r[f"country_{side}"], r[f"name_{side}"], r[f"points_{side}"]])
        widths = [2.0, 1.6] + ([1.6] if with_rank else []) + [3.6, 5.0, 1.9]
        table = _ra_docx_table(doc, headers, rows, widths_cm=widths, left_cols={len(headers) - 3, len(headers) - 2}, font_size=10)
        # ช่อง MATCH/LANE รวมสองแถวเป็นช่องเดียว และทำตัวหนาฝั่งผู้ชนะ
        for c in (0, 1):
            merged = table.rows[1].cells[c].merge(table.rows[2].cells[c])
            _ra_docx_set_cell_text(merged, rows[0][c], size_pt=10)
        winner_side = "a" if r.get("winner_id") and r["athlete_a"] and r["athlete_a"].id == r["winner_id"] else ("b" if r.get("winner_id") else None)
        if winner_side:
            win_row = table.rows[1 if winner_side == "a" else 2]
            for c in range(2, len(headers)):
                for run in win_row.cells[c].paragraphs[0].runs:
                    run.bold = True
        spacer = doc.add_paragraph()
        spacer.paragraph_format.space_after = Pt(6)


def make_results_approved_docx(event: Event) -> BytesIO:
    """เล่มผล Results Approved ของอีเวนต์ Shooting

    ใส่เฉพาะส่วนที่มีข้อมูลจริงในระบบ: ส่วนไหนยังไม่แข่ง/ยังไม่มีผล จะไม่ถูกพิมพ์เป็นหน้าว่าง
    """
    from docx import Document
    from docx.enum.text import WD_BREAK
    from docx.shared import Pt, Inches

    ctx = build_results_approved_context(event)
    setting = ctx["setting"]
    doc = Document()
    _ra_setup_section(doc.sections[0])
    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(10)
    style.paragraph_format.space_after = Pt(0)
    _ra_setup_header_footer(doc, ctx)
    logos = ctx.get("logos", {})
    event_title = ctx["event_title"]
    date_line = ctx["date_line"]
    approved = ctx["approved_text"]
    pending_break = [False]

    def new_page(landscape=False):
        """ขึ้นหน้าใหม่: ถ้าแนวกระดาษเปลี่ยนใช้ section break, ถ้าไม่เปลี่ยนให้หัวข้อถัดไปขึ้นหน้าใหม่เอง"""
        from docx.enum.section import WD_ORIENT, WD_SECTION
        current_landscape = doc.sections[-1].orientation == WD_ORIENT.LANDSCAPE
        if landscape != current_landscape:
            new_section = doc.add_section(WD_SECTION.NEW_PAGE)
            new_section.different_first_page_header_footer = False
            _ra_setup_section(new_section, landscape)
            pending_break[0] = False
        else:
            pending_break[0] = True

    def page_title(lines):
        _ra_docx_page_title(doc, lines, page_break_before=pending_break[0])
        pending_break[0] = False

    def _chunks(rows, per_page):
        """แบ่งแถวเป็นหน้า ๆ ให้จำนวนแถวใกล้เคียงกัน ไม่ให้เหลือเศษไม่กี่แถวไปขึ้นหน้าใหม่โดด ๆ"""
        if len(rows) <= per_page:
            return [rows]
        pages = -(-len(rows) // per_page)
        size = -(-len(rows) // pages)
        return [rows[i:i + size] for i in range(0, len(rows), size)]

    def paged_table(title_lines, headers, rows, per_page=34, landscape=False, detail_cols=None, **kw):
        """พิมพ์ตารางยาวแบบแบ่งหน้า: ทุกหน้ามีหัวข้อ หัวตาราง และบรรทัดรับรองครบ"""
        parts = _chunks(rows, per_page)
        for idx, part in enumerate(parts):
            if idx:
                new_page(landscape=landscape)
            lines = list(title_lines)
            if len(parts) > 1:
                lines.append((f"(PAGE {idx + 1} OF {len(parts)})", 9))
            page_title(lines)
            if detail_cols is not None:
                _ra_docx_station_detail_table(doc, ctx, part, detail_cols)
            else:
                _ra_docx_table(doc, headers, part, **kw)
            _ra_docx_approved_block(doc, approved)

    # ---------- ปก (ใช้ข้อความจากหน้าตั้งค่า ไม่ฝังคำว่า SEA GAMES ตายตัว) ----------
    doc.add_paragraph().paragraph_format.space_after = Pt(40)
    _ra_docx_add_center_image(doc, logos.get("cover_main"), width_inches=1.6)
    _ra_docx_page_title(doc, [(ctx["competition_title"], 20), (date_line, 12), (ctx["host_line"], 12)])
    doc.add_paragraph().paragraph_format.space_after = Pt(24)
    _ra_docx_add_title(doc, "PÉTANQUE", 26)
    _ra_docx_add_title(doc, event_title, 14)
    doc.add_paragraph().paragraph_format.space_after = Pt(12)
    _ra_docx_add_logo_row(doc, [logos.get("cover_bottom_1"), logos.get("cover_bottom_2"), logos.get("cover_bottom_3")], width_inches=0.8)
    doc.add_paragraph().paragraph_format.space_after = Pt(24)
    _ra_docx_add_title(doc, "RESULTS", 24)
    _ra_docx_add_title(doc, "APPROVED", 11)

    # ---------- เจ้าหน้าที่ (เฉพาะเมื่อกรอกข้อมูลไว้จริง) ----------
    president = _ra_setting_text(setting, "president_name", "")
    technical = _ra_setting_text(setting, "technical_name", "")
    if getattr(setting, "show_official_pages", True) and (president or technical or ctx["umpire_rows"]):
        new_page()
        if president:
            page_title([(_ra_setting_text(setting, "president_title", "").upper(), 12), (president.upper(), 12)])
        if technical:
            page_title([(_ra_setting_text(setting, "technical_title", "").upper(), 12), (technical.upper(), 12)])
        if ctx["umpire_rows"]:
            page_title([("UMPIRE", 16)])
            _ra_docx_table(doc, ["NO", "FAMILY NAME - GIVEN NAME", "FEDERATION"],
                           [[r["no"], r["name"], r["federation"]] for r in ctx["umpire_rows"]],
                           widths_cm=[1.4, 9.0, 4.4], left_cols={1}, font_size=10)
        _ra_docx_approved_block(doc, approved)

    if not ctx["athletes"]:
        out = BytesIO(); doc.save(out); out.seek(0)
        return out

    # ---------- ประเทศที่ส่งเข้าแข่ง ----------
    new_page()
    paged_table([(event_title, 16), (date_line, 11)], ["NO.", ctx["country_label"]],
                [[r["no"], r["country"]] for r in ctx["entry_countries"]],
                per_page=32, widths_cm=[2.0, 8.0], font_size=11)

    # ---------- รายชื่อ (ชื่อเต็มตามที่ลงทะเบียน ไม่ตัดแยกนามสกุลเอง เพราะชื่อไทย/ต่างชาติเรียงไม่เหมือนกัน) ----------
    new_page()
    if ctx["use_split_names"]:
        paged_table([("NAME LISTS", 16), (event_title, 11)], ["NO.", ctx["country_label"], "FAMILY NAME", "GIVEN NAME"],
                    [[r["no"], r["country"], r["family_name"], r["given_name"]] for r in ctx["name_rows"]],
                    per_page=34, widths_cm=[1.4, 4.2, 5.2, 5.2], left_cols={1, 2, 3}, font_size=10)
    else:
        paged_table([("NAME LISTS", 16), (event_title, 11)], ["NO.", ctx["country_label"], "NAME"],
                    [[r["no"], r["country"], r["name"]] for r in ctx["name_rows"]],
                    per_page=34, widths_cm=[1.6, 5.0, 9.0], left_cols={1, 2}, font_size=10)

    # ---------- รอบคัดเลือก 1 (เฉพาะเมื่อมีคะแนนแล้ว) ----------
    if _ra_has_any_score(ctx["qf1_rows"]):
        new_page()
        paged_table([(event_title, 13), ("QUALIFICATION ROUND 1", 13), (date_line, 10)],
                    ["NO.", ctx["country_label"], "NAME", "LANE", "POINTS", "RANK\n(QF1)"],
                    [[r["no"], r["country"], r["name"], r["lane"], r["points"], r["rank"]] for r in ctx["qf1_rows"]],
                    per_page=34, widths_cm=[1.3, 3.6, 6.0, 1.4, 1.8, 1.9], left_cols={1, 2}, font_size=10)

        detail = sorted(ctx["qf1_detail_rows"], key=lambda r: (int(r["rank"]) if str(r["rank"]).isdigit() else 9999))
        for r in detail:
            r["rank_display"] = r["rank"]
        new_page(landscape=True)
        paged_table([(f"{event_title} · QUALIFICATION SHOOTING ROUND 1", 12)], None, detail,
                    per_page=36, landscape=True, detail_cols=[("TOTAL", "total")])

    # ---------- รอบคัดเลือก 2 ----------
    if ctx["qf2_rows"] and _ra_has_any_score(ctx["qf2_rows"], "r2"):
        new_page()
        paged_table([(event_title, 13), ("QUALIFICATION ROUND 2", 13), (date_line, 10)],
                    ["RANK\n(QF1)", ctx["country_label"], "NAME", "LANE", "R1", "R2", "TOTAL", "RANK\n(QF2)"],
                    [[r["qf1_rank"], r["country"], r["name"], r["lane"], r["r1"], r["r2"], r["total"], r["qf2_rank"]] for r in ctx["qf2_rows"]],
                    per_page=34, widths_cm=[1.7, 3.0, 4.4, 1.4, 1.1, 1.1, 1.7, 1.7], left_cols={1, 2}, font_size=10)

        for r in ctx["qf2_detail_rows"]:
            r["rank_display"] = r["qf2_rank"]
        detail2 = sorted(ctx["qf2_detail_rows"], key=lambda r: (int(r["qf2_rank"]) if str(r["qf2_rank"]).isdigit() else 9999))
        new_page(landscape=True)
        paged_table([(f"{event_title} · QUALIFICATION SHOOTING ROUND 2", 12)], None, detail2,
                    per_page=36, landscape=True, detail_cols=[("R1", "r1"), ("R2", "r2"), ("TOTAL", "total")])

    # ---------- รอบน็อกเอาต์: เฉพาะแมตช์ที่มีผู้เล่นครบทั้งสองฝั่ง ----------
    # เฉพาะแมตช์ที่แข่งแล้วจริง: มีผู้เล่นครบสองฝั่ง และมีผู้ชนะหรือมีคะแนนแล้ว (ไม่พิมพ์คู่ที่ยังไม่ได้ยิง)
    def _played(r):
        pts = [p for p in (r["points_a"], r["points_b"]) if isinstance(p, int)]
        return bool(r.get("winner_id")) or any(p > 0 for p in pts)
    ko_real = [r for r in ctx["bracket_rows"] if r["athlete_a"] and r["athlete_b"] and _played(r)]
    round_titles = [("R16", "ROUND OF 16"), ("QF", "QUARTERFINAL ROUND"), ("SF", "SEMIFINAL ROUND"), ("F", "FINAL ROUND")]
    ko_by_round = [(title, [r for r in ko_real if r["round_name"] == key]) for key, title in round_titles]
    ko_by_round = [(t, rows) for t, rows in ko_by_round if rows]
    if ko_by_round:
        new_page()
        page_title([(event_title, 13), (date_line, 10)])
        for title, rows in ko_by_round:
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.keep_with_next = True
            run = p.add_run(title)
            run.bold = True
            run.font.size = Pt(12)
            _ra_docx_match_table(doc, ctx, rows, with_rank=title != "FINAL ROUND")
        _ra_docx_approved_block(doc, approved)

    # ---------- ผลอันดับ/เหรียญ: เฉพาะเมื่อรอบชิงมีผู้ชนะแล้วจริง (ไม่เดาจากอันดับรอบคัดเลือก) ----------
    final_done = any(r["round_name"] == "F" and r.get("winner_id") for r in ctx["bracket_rows"])
    if final_done and ctx["medal_rows"]:
        new_page()
        page_title([("RANKING RESULT", 16), (event_title, 13), (date_line, 10)])
        if ctx["use_split_names"]:
            _ra_docx_table(doc, ["MEDAL", ctx["country_label"], "FAMILY NAME", "GIVEN NAME"],
                           [[r["medal"], r["country"], r["family_name"], r["given_name"]] for r in ctx["medal_rows"]],
                           widths_cm=[2.8, 4.0, 4.4, 4.4], left_cols={1, 2, 3}, font_size=11, row_height_cm=1.1)
        else:
            _ra_docx_table(doc, ["MEDAL", ctx["country_label"], "NAME"],
                           [[r["medal"], r["country"], (r["athlete"].name or "").upper()] for r in ctx["medal_rows"]],
                           widths_cm=[3.0, 4.6, 8.0], left_cols={1, 2}, font_size=11, row_height_cm=1.1)
        _ra_docx_approved_block(doc, approved)

    out = BytesIO()
    doc.save(out)
    out.seek(0)
    return out


@app.route("/events/<int:event_id>/results-approved/settings", methods=["GET", "POST"])
@login_required
@role_required("admin", "superadmin")
def results_approved_settings(event_id: int):
    event = Event.query.get_or_404(event_id)
    setting = get_results_approved_setting(event, create=True)
    if request.method == "POST":
        name_format = request.form.get("name_format", "single").strip().lower()
        setting.name_format = name_format if name_format in {"single", "first", "last"} else "single"
        setting.competition_title = request.form.get("competition_title", "").strip() or event.name
        setting.host_line = request.form.get("host_line", "").strip()
        setting.date_line = request.form.get("date_line", "").strip() or _ra_date_text(event)
        setting.location_line = request.form.get("location_line", "").strip()
        setting.country_label = request.form.get("country_label", "COUNTRY").strip().upper() or "COUNTRY"
        setting.president_title = request.form.get("president_title", "").strip()
        setting.president_name = request.form.get("president_name", "").strip()
        setting.technical_title = request.form.get("technical_title", "").strip()
        setting.technical_name = request.form.get("technical_name", "").strip()
        setting.umpires_text = request.form.get("umpires_text", "").strip()
        setting.approved_text = request.form.get("approved_text", "").strip() or "……………………………APPROVED"
        setting.show_official_pages = request.form.get("show_official_pages") == "yes"
        for field_name, attr in RESULTS_APPROVED_LOGO_FIELDS.items():
            if request.form.get(f"clear_{field_name}") == "yes":
                setattr(setting, attr, None)
            saved_logo = _ra_save_uploaded_logo(event.id, field_name)
            if saved_logo:
                setattr(setting, attr, saved_logo)
        db.session.commit()
        flash("บันทึกตั้งค่า Results Approved แล้ว", "success")
        return redirect(url_for("results_approved", event_id=event.id))
    logo_field_labels = [
        ("cover_main_logo", "โลโก้หลักบนปก"),
        ("cover_bottom_logo_1", "โลโก้ล่างปก 1"),
        ("cover_bottom_logo_2", "โลโก้ล่างปก 2"),
        ("cover_bottom_logo_3", "โลโก้ล่างปก 3"),
        ("header_logo_1", "โลโก้หัวกระดาษ 1"),
        ("header_logo_2", "โลโก้หัวกระดาษ 2"),
        ("header_logo_3", "โลโก้หัวกระดาษ 3"),
        ("header_logo_4", "โลโก้หัวกระดาษ 4"),
        ("side_logo", "โลโก้มุมซ้ายในหน้าผล"),
    ]
    return render_template(
        "results_approved_settings.html",
        event=event,
        setting=setting,
        logos=_ra_logo_map(setting),
        logo_field_labels=logo_field_labels,
    )


@app.route("/events/<int:event_id>/results-approved")
@login_required
def results_approved(event_id: int):
    event = Event.query.get_or_404(event_id)
    ctx = build_results_approved_context(event)
    return render_template("results_approved.html", **ctx)


@app.route("/events/<int:event_id>/results-approved.docx")
@login_required
def results_approved_docx(event_id: int):
    event = Event.query.get_or_404(event_id)
    try:
        stream = make_results_approved_docx(event)
    except ModuleNotFoundError:
        flash("ยังไม่ได้ติดตั้ง python-docx ให้รัน: python -m pip install -r requirements.txt", "warning")
        return redirect(url_for("results_approved", event_id=event.id))
    filename = f"results_approved_event_{event.id}.docx"
    return send_file(
        stream,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

@app.route("/events/<int:event_id>/stats")
def event_stats(event_id: int):
    event = Event.query.get_or_404(event_id)
    round1_rows = build_round_ranking(event, 1)
    best = round1_rows[0] if round1_rows else None
    return render_template("stats.html", event=event, best=best, round1_rows=round1_rows)


# -----------------------------------------------------------------------------
# Public Live Report API / iframe support for Live Report Board
# -----------------------------------------------------------------------------
@app.after_request
def live_report_public_headers(response):
    # เปิด CORS เฉพาะ API สาธารณะ (อ่านอย่างเดียว) ไม่เปิดให้ทุกหน้า
    if request.path.startswith("/api/public/"):
        response.headers.setdefault("Access-Control-Allow-Origin", "*")
        response.headers.setdefault("Access-Control-Allow-Headers", "Content-Type")
        response.headers.setdefault("Access-Control-Allow-Methods", "GET, OPTIONS")
    # ให้ Report Board ฝังหน้า overview/bracket ผ่าน iframe ได้
    response.headers.pop("X-Frame-Options", None)
    response.headers.setdefault("Content-Security-Policy", "frame-ancestors *")
    return response


def _lr_shooting_event_payload(event):
    return {
        "id": event.id,
        "name": event.name,
        "event_group": event.event_group,
        "category": event.category,
        "competition_date": event.competition_date.isoformat() if event.competition_date else None,
        "location": event.location,
        "lane_count": event.lane_count,
        "athlete_count": len(event.athletes),
        "overview_url": url_for('event_overview', event_id=event.id, _external=True),
        "public_live_url": url_for('api_public_shooting_report', event_id=event.id, _external=True),
    }


def _lr_athlete_payload(event, athlete, rank=None, official_ranks=None):
    s1 = summarize_round(athlete.id, 1)
    s2 = summarize_round(athlete.id, 2)
    r1 = s1.get("total", 0)
    r2 = s2.get("total", 0)
    total = r1 + r2
    official_ranks = official_ranks or {}
    return {
        "rank": rank,
        "id": athlete.id,
        "bib_no": athlete.bib_no,
        "name": athlete.name,
        "affiliation": athlete.affiliation,
        "lane_no": athlete.lane_no,
        "lane_order": athlete.lane_order,
        "start_order": athlete.start_order,
        "status": athlete.status,
        "red_card_count": athlete.red_card_count,
        "round1_total": r1,
        "round2_total": r2,
        "total": total,
        # เกณฑ์ตัดสินเดียวกับหน้า Overview (ranking_key): คะแนนรวม > จำนวนลูก 5 > จำนวนลูก 3 > shoot-off
        "count_5": s1.get("count_5", 0) + s2.get("count_5", 0),
        "count_3": s1.get("count_3", 0) + s2.get("count_3", 0),
        "tiebreak_total": s1.get("tiebreak_total", 0) + s2.get("tiebreak_total", 0),
        "official_round1_rank": official_ranks.get(1, {}).get(athlete.id),
        "official_round2_rank": official_ranks.get(2, {}).get(athlete.id),
    }


@app.route('/api/public/shooting/events')
def api_public_shooting_events():
    events = Event.query.order_by(Event.created_at.desc(), Event.id.desc()).all()
    return jsonify({"ok": True, "source": "shooting", "events": [_lr_shooting_event_payload(e) for e in events]})


@app.route('/api/public/shooting/event/<int:event_id>/report')
def api_public_shooting_report(event_id: int):
    event = Event.query.get_or_404(event_id)
    preload_event_score_data(event)
    athletes = sorted(event.athletes, key=lambda a: (a.lane_no, a.lane_order, a.start_order))
    official_ranks = compute_round_ranks(event)
    rows = [_lr_athlete_payload(event, a, official_ranks=official_ranks) for a in athletes]
    if not event.has_round_two:
        # อีเวนต์รอบเดียว: ใช้อันดับทางการของรอบ 1 ตรงกับหน้า Overview ทุกประการ
        ranking = sorted(rows, key=lambda r: (r["official_round1_rank"] is None, r["official_round1_rank"] or 0, r["start_order"] or 0))
    else:
        ranking = sorted(rows, key=lambda r: (r["total"], r["count_5"], r["count_3"], r["tiebreak_total"]), reverse=True)
    for idx, row in enumerate(ranking, start=1):
        row["rank"] = idx
    return jsonify({
        "ok": True,
        "source": "shooting",
        "event": _lr_shooting_event_payload(event),
        "athletes": rows,
        "ranking": ranking,
        "live_url": url_for('event_overview', event_id=event.id, _external=True),
    })


@app.route('/public/shooting/<int:event_id>/live')
def public_shooting_live(event_id: int):
    # หน้า public สำหรับเอาไป iframe ใน Report Board โดยไม่ต้อง login
    event = Event.query.get_or_404(event_id)
    return redirect(url_for('event_overview', event_id=event.id, round=request.args.get('round', 1)))


# ---------------------------------------------------------------------------
# ธีมของทั้งระบบ
# ---------------------------------------------------------------------------
import re as _re

THEME_COLOR_FIELDS = ["color_primary", "color_primary_dark", "color_soft", "color_cream",
                      "color_accent", "color_ink", "color_line"]
THEME_TEXT_FIELDS = {
    "name": 120, "short_name": 60, "eyebrow": 160, "title": 160, "subtitle": 255,
    "export_kicker": 160, "export_title": 255, "export_subtitle": 255,
    "footer_tagline": 255, "footer_event": 255,
}
THEME_IMAGE_KINDS = {"poster": "ภาพพื้นหลังแบนเนอร์", "logo": "โลโก้งาน", "partners": "แถบโลโก้ผู้สนับสนุน"}
THEME_IMAGE_MAX_BYTES = 6 * 1024 * 1024
_HEX_RE = _re.compile(r"^#[0-9a-fA-F]{6}$")

KKU_THEME_DEFAULTS = dict(
    name="KKU 2026 World Championship", short_name="KKU 2026",
    eyebrow="Khon Kaen University · Thailand", title="Petanque Shooting",
    subtitle="52nd PÉTANQUE WORLD CHAMPIONSHIP 2026 · 48 Nations, One Spirit",
    export_kicker="Official Shooting Results", export_title="52nd PÉTANQUE WORLD CHAMPIONSHIP 2026",
    export_subtitle="Khon Kaen University, Thailand",
    footer_tagline="SPORT  |  CULTURE  |  FRIENDSHIP  |  A BRIGHTER TOMORROW",
    footer_event="KKU 2026 · Pétanque Unites the World",
    show_hero=True,
    color_primary="#ef4b12", color_primary_dark="#c9360b", color_soft="#fff1e8", color_cream="#fffaf4",
    color_accent="#37b8b1", color_ink="#172033", color_line="#f3d2bf",
    poster_static="kku2026_theme_poster.jpg", logo_static="kku2026_event_logo.jpg",
    partners_static="kku2026_partner_logos.jpg",
)
PLAIN_THEME_DEFAULTS = dict(
    name="มาตรฐาน (ทั่วไป)", short_name="Petanque",
    eyebrow="Petanque Shooting System", title="Petanque Shooting",
    subtitle="ระบบบันทึกคะแนนและจัดอันดับเปตอง Shooting",
    export_kicker="Official Shooting Results", export_title="PETANQUE SHOOTING", export_subtitle="",
    footer_tagline="", footer_event="Petanque Shooting", show_hero=True,
    color_primary="#2563eb", color_primary_dark="#1e40af", color_soft="#eff4ff", color_cream="#f8fafc",
    color_accent="#14b8a6", color_ink="#0f172a", color_line="#cbd5e1",
)

PHUPHAN_THEME_DEFAULTS = dict(
    name="ภูพานเกมส์ · สกลนคร", short_name="ภูพานเกมส์",
    eyebrow="เมืองสกลนคร · SMART & SPIRIT", title="ภูพานเกมส์",
    subtitle="ศรัทธา สานฝัน มุ่งมั่นชัยชนะ · การแข่งขันเปตอง Shooting",
    export_kicker="Official Shooting Results", export_title="ภูพานเกมส์ · การแข่งขันเปตอง Shooting",
    export_subtitle="จังหวัดสกลนคร", footer_tagline="ศรัทธา  |  สานฝัน  |  มุ่งมั่นชัยชนะ",
    footer_event="ภูพานเกมส์ · SMART & SPIRIT", show_hero=True,
    color_primary="#aa241a", color_primary_dark="#642d2b", color_soft="#f7ebea", color_cream="#fcf7f7",
    color_accent="#f2b400", color_ink="#211723", color_line="#e9c6c3",
    poster_static="themes/phuphan_poster.jpg", logo_static="themes/phuphan_logo.jpg",
)
BUILTIN_THEMES = [KKU_THEME_DEFAULTS, PLAIN_THEME_DEFAULTS, PHUPHAN_THEME_DEFAULTS]


def ensure_default_themes() -> None:
    """เพิ่มธีมที่มากับระบบที่ยังไม่มี (ตามชื่อ) · ธีมแรกที่สร้างในฐานข้อมูลว่างจะถูกเปิดใช้"""
    try:
        existing = {t.name for t in SiteTheme.query.filter_by(is_builtin=True).all()}
        empty = SiteTheme.query.count() == 0
    except Exception:
        db.session.rollback()
        return
    for idx, defaults in enumerate(BUILTIN_THEMES):
        if defaults["name"] not in existing:
            db.session.add(SiteTheme(is_active=empty and idx == 0, is_builtin=True, **defaults))
    db.session.flush()


def _hex_to_rgb(value: str) -> tuple[int, int, int]:
    value = value.lstrip("#")
    return int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16)


def _mix_hex(a: str, b: str, t: float) -> str:
    ra, rb = _hex_to_rgb(a), _hex_to_rgb(b)
    return "#" + "".join(f"{round(x + (y - x) * t):02x}" for x, y in zip(ra, rb))


def _theme_image_version(theme, kind: str) -> str | None:
    for asset in theme.assets:
        if asset.kind == kind:
            return str(int((asset.updated_at or datetime.utcnow()).timestamp()))
    return None


def theme_image_url(theme, kind: str) -> str | None:
    if theme is None:
        return None
    version = _theme_image_version(theme, kind) if getattr(theme, "id", None) else None
    if version:
        return url_for("theme_asset", theme_id=theme.id, kind=kind, v=version)
    static_name = getattr(theme, f"{kind}_static", None)
    if static_name and os.path.exists(os.path.join(BASE_DIR, "static", static_name)):
        return url_for("static", filename=static_name)
    return None


def build_theme_view(theme) -> SimpleNamespace:
    """ค่าที่ template ใช้ (สี, ภาพ, ข้อความ) จากธีมที่เปิดใช้งาน"""
    colors = {f: getattr(theme, f, None) or KKU_THEME_DEFAULTS[f] for f in THEME_COLOR_FIELDS}
    for f, v in colors.items():
        if not _HEX_RE.match(v):
            colors[f] = KKU_THEME_DEFAULTS[f]
    primary, dark = colors["color_primary"], colors["color_primary_dark"]
    css_vars = {
        "--th-primary": primary,
        "--th-primary-dark": dark,
        "--th-primary-light": _mix_hex(primary, "#ffffff", 0.18),
        "--th-soft": colors["color_soft"],
        "--th-tint": _mix_hex(primary, "#ffffff", 0.78),
        "--th-cream": colors["color_cream"],
        "--th-accent": colors["color_accent"],
        "--th-ink": colors["color_ink"],
        "--th-line": colors["color_line"],
        "--th-primary-rgb": ",".join(map(str, _hex_to_rgb(primary))),
        "--th-dark-rgb": ",".join(map(str, _hex_to_rgb(_mix_hex(dark, "#000000", 0.45)))),
        "--th-accent-rgb": ",".join(map(str, _hex_to_rgb(colors["color_accent"]))),
        "--th-cream-rgb": ",".join(map(str, _hex_to_rgb(colors["color_cream"]))),
        # ชื่อเดิมของธีม KKU ยังใช้ได้
        "--kku-orange": primary, "--kku-orange-dark": dark, "--kku-orange-soft": colors["color_soft"],
        "--kku-cream": colors["color_cream"], "--kku-teal": colors["color_accent"],
        "--kku-ink": colors["color_ink"], "--kku-line": colors["color_line"],
    }
    text = {f: (getattr(theme, f, None) or "") for f in THEME_TEXT_FIELDS}
    slug = _re.sub(r"[^A-Za-z0-9]+", "", text["short_name"] or "") or "Results"
    return SimpleNamespace(
        id=getattr(theme, "id", None),
        css_vars="".join(f"{k}:{v};" for k, v in css_vars.items()),
        poster_url=theme_image_url(theme, "poster"),
        logo_url=theme_image_url(theme, "logo"),
        partners_url=theme_image_url(theme, "partners"),
        show_hero=bool(getattr(theme, "show_hero", True)),
        download_suffix=slug,
        **text,
    )


def active_site_theme():
    cached = getattr(request, "_site_theme_view", None) if has_request_context() else None
    if cached is not None:
        return cached
    theme = None
    try:
        theme = SiteTheme.query.filter_by(is_active=True).order_by(SiteTheme.id).first()
        if theme is None and not SiteTheme.query.count():
            ensure_default_themes()
            db.session.commit()
            theme = SiteTheme.query.filter_by(is_active=True).first()
    except Exception:
        db.session.rollback()
        theme = None
    view = build_theme_view(theme or SimpleNamespace(id=None, assets=[], **KKU_THEME_DEFAULTS))
    if has_request_context():
        request._site_theme_view = view
    return view


@app.context_processor
def inject_site_theme():
    return {"site_theme": active_site_theme()}


def _sniff_image_mimetype(data: bytes) -> str | None:
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return "image/png"
    if data.startswith(b"\xff\xd8\xff"):
        return "image/jpeg"
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return "image/gif"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return "image/webp"
    return None


def _apply_theme_form(theme: SiteTheme, form, files) -> list[str]:
    """อัปเดตธีมจากฟอร์ม คืนรายการปัญหา (ถ้ามี) โดยไม่บันทึกส่วนที่ผิด"""
    problems = []
    for field, limit in THEME_TEXT_FIELDS.items():
        if field in form:
            setattr(theme, field, (form.get(field) or "").strip()[:limit])
    if not theme.name:
        theme.name = "ธีมใหม่"
    for field in THEME_COLOR_FIELDS:
        value = (form.get(field) or "").strip()
        if value:
            if _HEX_RE.match(value):
                setattr(theme, field, value.lower())
            else:
                problems.append(f"สี {field} ไม่ถูกต้อง")
    theme.show_hero = form.get("show_hero") == "yes"
    existing = {a.kind: a for a in theme.assets}
    for kind, label in THEME_IMAGE_KINDS.items():
        if form.get(f"remove_{kind}") == "yes":
            if kind in existing:
                db.session.delete(existing.pop(kind))
            setattr(theme, f"{kind}_static", None)
        upload = files.get(f"image_{kind}")
        if not upload or not upload.filename:
            continue
        data = upload.read(THEME_IMAGE_MAX_BYTES + 1)
        if len(data) > THEME_IMAGE_MAX_BYTES:
            problems.append(f"{label}: ไฟล์ใหญ่เกิน 6 MB")
            continue
        mimetype = _sniff_image_mimetype(data)
        if not mimetype:
            problems.append(f"{label}: ต้องเป็นไฟล์ภาพ PNG, JPG, WEBP หรือ GIF")
            continue
        asset = existing.get(kind)
        if asset is None:
            asset = SiteThemeAsset(theme=theme, kind=kind, mimetype=mimetype, data=data)
            db.session.add(asset)
        else:
            asset.mimetype, asset.data, asset.updated_at = mimetype, data, datetime.utcnow()
    theme.updated_at = datetime.utcnow()
    return problems


@app.route("/theme-asset/<int:theme_id>/<kind>")
def theme_asset(theme_id: int, kind: str):
    if kind not in THEME_IMAGE_KINDS:
        abort(404)
    asset = SiteThemeAsset.query.filter_by(theme_id=theme_id, kind=kind).first_or_404()
    response = app.response_class(asset.data, mimetype=asset.mimetype)
    response.headers["Cache-Control"] = "public, max-age=31536000, immutable"
    response.headers["X-Content-Type-Options"] = "nosniff"
    return response


@app.route("/admin/themes")
@login_required
@role_required("superadmin")
def manage_themes():
    themes = SiteTheme.query.order_by(SiteTheme.is_active.desc(), SiteTheme.id).all()
    return render_template("themes.html", themes=[(t, build_theme_view(t)) for t in themes])


@app.route("/admin/themes/new", methods=["GET", "POST"])
@app.route("/admin/themes/<int:theme_id>/edit", methods=["GET", "POST"])
@login_required
@role_required("superadmin")
def edit_theme(theme_id: int | None = None):
    theme = SiteTheme.query.get_or_404(theme_id) if theme_id else None
    if request.method == "POST":
        is_new = theme is None
        if is_new:
            theme = SiteTheme(**{k: v for k, v in PLAIN_THEME_DEFAULTS.items()})
            db.session.add(theme)
        problems = _apply_theme_form(theme, request.form, request.files)
        activate = request.form.get("activate") == "yes"
        if activate:
            SiteTheme.query.update({SiteTheme.is_active: False})
            theme.is_active = True
        db.session.commit()
        for p in problems:
            flash(p, "warning")
        flash(("สร้าง" if is_new else "บันทึก") + f"ธีม “{theme.name}” แล้ว" + (" · เปิดใช้ทั้งระบบ" if activate else ""), "success")
        return redirect(url_for("edit_theme", theme_id=theme.id))
    view = build_theme_view(theme or SimpleNamespace(id=None, assets=[], **PLAIN_THEME_DEFAULTS))
    source = theme or SimpleNamespace(**PLAIN_THEME_DEFAULTS)
    return render_template("theme_form.html", theme=theme, source=source, view=view,
                           image_kinds=THEME_IMAGE_KINDS, color_fields=THEME_COLOR_FIELDS)


@app.route("/admin/themes/<int:theme_id>/activate", methods=["POST"])
@login_required
@role_required("superadmin")
def activate_theme(theme_id: int):
    theme = SiteTheme.query.get_or_404(theme_id)
    SiteTheme.query.update({SiteTheme.is_active: False})
    theme.is_active = True
    db.session.commit()
    flash(f"เปลี่ยนธีมทั้งระบบเป็น “{theme.name}” แล้ว", "success")
    return redirect(url_for("manage_themes"))


@app.route("/admin/themes/<int:theme_id>/duplicate", methods=["POST"])
@login_required
@role_required("superadmin")
def duplicate_theme(theme_id: int):
    src = SiteTheme.query.get_or_404(theme_id)
    copy = SiteTheme(is_active=False, is_builtin=False)
    for col in SiteTheme.__table__.columns.keys():
        if col not in {"id", "is_active", "is_builtin", "created_at", "updated_at"}:
            setattr(copy, col, getattr(src, col))
    copy.name = (f"{src.name} (สำเนา)")[:120]
    db.session.add(copy)
    db.session.flush()
    for a in src.assets:
        db.session.add(SiteThemeAsset(theme_id=copy.id, kind=a.kind, mimetype=a.mimetype, data=a.data))
    db.session.commit()
    flash(f"คัดลอกเป็น “{copy.name}” แล้ว", "success")
    return redirect(url_for("edit_theme", theme_id=copy.id))


@app.route("/admin/themes/<int:theme_id>/delete", methods=["POST"])
@login_required
@role_required("superadmin")
def delete_theme(theme_id: int):
    theme = SiteTheme.query.get_or_404(theme_id)
    if theme.is_active:
        flash("ลบธีมที่กำลังใช้งานไม่ได้ ให้เปิดใช้ธีมอื่นก่อน", "warning")
    elif theme.is_builtin:
        flash("ธีมที่มากับระบบลบไม่ได้ (แก้ไขหรือคัดลอกได้)", "warning")
    else:
        name = theme.name
        db.session.delete(theme)
        db.session.commit()
        flash(f"ลบธีม “{name}” แล้ว", "info")
    return redirect(url_for("manage_themes"))


def init_database_for_deploy() -> None:
    """Create database tables when running under gunicorn/Railway.

    Flask code inside __main__ is not executed by `gunicorn app:app`,
    so Railway needs this initialization during import.
    """
    os.makedirs(os.path.join(BASE_DIR, "instance"), exist_ok=True)
    with app.app_context():
        db.create_all()
        ensure_schema()
        seed_defaults()
        # gunicorn --preload เปิด connection ใน master ก่อน fork; ต้องปิดทิ้งไม่ให้ worker ใช้ connection ร่วมกัน
        db.engine.dispose()


# ให้ Railway/gunicorn สร้างตารางและ user ตั้งต้นทันทีตอน import app
init_database_for_deploy()


if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 8001)),
        debug=os.environ.get("FLASK_DEBUG", "0") == "1"
    )
