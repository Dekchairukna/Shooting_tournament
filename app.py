from __future__ import annotations

import json
import os
from io import BytesIO
from datetime import date, datetime
from functools import wraps
from typing import Dict, List, Tuple
from flask import jsonify

from flask import (
    Flask,
    flash,
    jsonify,
    redirect,
    render_template,
    request,
    send_file,
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
from openpyxl import Workbook, load_workbook

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, "instance", "shooting.db")

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "dev-secret-change-me")

database_url = os.environ.get("DATABASE_URL")
if database_url:
    if database_url.startswith("postgres://"):
        database_url = database_url.replace("postgres://", "postgresql://", 1)
    app.config["SQLALCHEMY_DATABASE_URI"] = database_url
else:
    app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DB_PATH}"

app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

db = SQLAlchemy(app)
login_manager = LoginManager(app)
login_manager.login_view = "login"

LANG_LABELS = {
    "th": {
        "shooting_title": "ประเภทสุดยอดความแม่นยำ (SHOOTING)",
    },
    "en": {
        "shooting_title": "Precision Shooting",
    },
    "fr": {
        "shooting_title": "Tir de précision",
    },
}

ROUND_LABELS = {
    1: "รอบที่ 1",
    2: "รอบที่ 2",
}

SCORECARD_ROUND_LABELS = [
    "รอบที่ 1",
    "รอบที่ 2",
    "รอบ 8 คน",
    "รอบรองชนะเลิศ",
    "รอบชิงชนะเลิศ",
]

STATIONS = [1, 2, 3, 4, 5]
DISTANCES = [6, 7, 8, 9]
MAX_RED_CARDS = 2


class User(UserMixin, db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False, default="user")
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def set_password(self, password: str) -> None:
        self.password_hash = generate_password_hash(password)

    def check_password(self, password: str) -> bool:
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
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    entries = db.relationship("ScoreEntry", backref="athlete", cascade="all, delete-orphan", lazy=True)
    signatures = db.relationship("ScoreSignature", backref="athlete", cascade="all, delete-orphan", lazy=True)
    tiebreaks = db.relationship("TieBreakEntry", backref="athlete", cascade="all, delete-orphan", lazy=True)


class ScoreEntry(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    athlete_id = db.Column(db.Integer, db.ForeignKey("athlete.id"), nullable=False)
    round_no = db.Column(db.Integer, nullable=False, default=1)
    station_no = db.Column(db.Integer, nullable=False)
    distance_m = db.Column(db.Integer, nullable=False)
    score = db.Column(db.Integer, nullable=False, default=0)
    is_red_card = db.Column(db.Boolean, nullable=False, default=False)
    updated_at = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)


class ScoreSignature(db.Model):
    id = db.Column(db.Integer, primary_key=True)
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


class TieBreakEntry(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    athlete_id = db.Column(db.Integer, db.ForeignKey("athlete.id"), nullable=False)
    round_no = db.Column(db.Integer, nullable=False)
    station_no = db.Column(db.Integer, nullable=False)
    score = db.Column(db.Integer, nullable=False, default=0)


class BracketMatch(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    event_id = db.Column(db.Integer, db.ForeignKey("event.id"), nullable=False)
    round_name = db.Column(db.String(20), nullable=False)
    match_no = db.Column(db.Integer, nullable=False)
    athlete_a_id = db.Column(db.Integer, nullable=True)
    athlete_b_id = db.Column(db.Integer, nullable=True)
    winner_id = db.Column(db.Integer, nullable=True)


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

#---------------คะแนนเรียลไทม์------------
def build_bracket_row_data(event, athlete, round_name):
    if not athlete:
        return {
            "athlete_id": None,
            "seed": "",
            "team": "",
            "name": "",
            "r1": "",
            "r1r2": "",
            "stations": [0, 0, 0, 0, 0],
            "total": 0,
        }

    round_map = {"QF": 3, "SF": 4, "F": 5}
    round_no = round_map.get(round_name)

    r1_summary = summarize_round(athlete.id, 1)
    r2_summary = summarize_round(athlete.id, 2) if event.has_round_two else {"total": 0}
    current_summary = summarize_round(athlete.id, round_no) if round_no else {"by_station": {}, "total": 0}

    stations = []
    by_station = current_summary.get("by_station", {})
    for station in [1, 2, 3, 4, 5]:
        station_data = by_station.get(station, {"total": 0})
        stations.append(station_data.get("total", 0))

    return {
        "athlete_id": athlete.id,
        "seed": "",
        "team": athlete.affiliation or "",
        "name": athlete.name or "",
        "r1": r1_summary["total"],
        "r1r2": r1_summary["total"] + r2_summary["total"],
        "stations": stations,
        "total": current_summary.get("total", 0),
    }

def ensure_schema() -> None:
    os.makedirs(os.path.join(BASE_DIR, "instance"), exist_ok=True)
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
        event_cols = {row[1] for row in conn.exec_driver_sql("PRAGMA table_info(event)").fetchall()}
        if "round_two_advancers" not in event_cols:
            conn.exec_driver_sql("ALTER TABLE event ADD COLUMN round_two_advancers INTEGER DEFAULT 4")
        tables = {row[0] for row in conn.exec_driver_sql("SELECT name FROM sqlite_master WHERE type='table'").fetchall()}
        if "bracket_match" not in tables:
            conn.exec_driver_sql("CREATE TABLE bracket_match (id INTEGER PRIMARY KEY AUTOINCREMENT, event_id INTEGER NOT NULL, round_name VARCHAR(20) NOT NULL, match_no INTEGER NOT NULL, athlete_a_id INTEGER, athlete_b_id INTEGER, winner_id INTEGER)")


def event_theme(category: str) -> str:
    return {
        "ชาย": "male",
        "หญิง": "female",
        "ผสม": "mixed",
    }.get(category, "mixed")


def scorecard_positions(round_no: int) -> dict:
    row_y = {1: 462, 2: 536}.get(round_no, 462)
    station_starts = {1: 170, 2: 374, 3: 580, 4: 785, 5: 990}
    distances = [6, 7, 8, 9]
    positions = {}
    for station, start_x in station_starts.items():
        for idx, distance in enumerate(distances):
            positions[(station, distance)] = {"left": start_x + (idx * 40), "top": row_y}
        positions[(station, "total")] = {"left": start_x + 160, "top": row_y}
    positions[("grand_total", "value")] = {"left": 1196, "top": row_y}
    positions[("rank", "value")] = {"left": 1268, "top": row_y}
    return positions


def seed_defaults() -> None:
    if not User.query.filter_by(username="superadmin").first():
        u = User(username="superadmin", role="superadmin")
        u.set_password("yagami1225")
        db.session.add(u)
    if not User.query.filter_by(username="admin").first():
        u = User(username="admin", role="admin")
        u.set_password("admin1234")
        db.session.add(u)
    if not User.query.filter_by(username="viewer").first():
        u = User(username="viewer", role="user")
        u.set_password("viewer1234")
        db.session.add(u)
    db.session.commit()


def get_round_score_map(athlete_id: int, round_no: int) -> Dict[Tuple[int, int], ScoreEntry]:
    entries = ScoreEntry.query.filter_by(athlete_id=athlete_id, round_no=round_no).all()
    return {(e.station_no, e.distance_m): e for e in entries}


def get_round_signature(athlete_id: int, round_no: int) -> ScoreSignature | None:
    return ScoreSignature.query.filter_by(athlete_id=athlete_id, round_no=round_no).first()


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
                ))
                changed = True
    if changed:
        db.session.commit()


def ensure_signature(athlete_id: int, round_no: int) -> ScoreSignature:
    signature = get_round_signature(athlete_id, round_no)
    if signature:
        return signature
    signature = ScoreSignature(athlete_id=athlete_id, round_no=round_no)
    db.session.add(signature)
    db.session.commit()
    return signature


def summarize_round(athlete_id: int, round_no: int) -> dict:
    entries = ScoreEntry.query.filter_by(athlete_id=athlete_id, round_no=round_no).all()
    total = sum(e.score for e in entries)
    count_5 = sum(1 for e in entries if e.score == 5)
    count_3 = sum(1 for e in entries if e.score == 3)
    red_cards = sum(1 for e in entries if e.is_red_card)
    by_station = {}
    for station_no in STATIONS:
        station_entries = [e for e in entries if e.station_no == station_no]
        by_station[station_no] = {
            "distances": {e.distance_m: e.score for e in station_entries},
            "total": sum(e.score for e in station_entries),
        }
    tiebreak_total = sum(e.score for e in TieBreakEntry.query.filter_by(athlete_id=athlete_id, round_no=round_no).all())
    return {
        "total": total,
        "count_5": count_5,
        "count_3": count_3,
        "red_cards": red_cards,
        "by_station": by_station,
        "tiebreak_total": tiebreak_total,
    }


def build_scorecard_template_data(athlete_id: int) -> dict:
    entries = ScoreEntry.query.filter_by(athlete_id=athlete_id).all()

    score_map: dict[tuple[int, int, int], dict] = {}
    station_totals: dict[tuple[int, int], int] = {}
    station_reds: dict[tuple[int, int], int] = {}
    round_totals: dict[int, int] = {}

    for round_no in [1, 2, 3, 4, 5]:
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


def ranking_key(item: dict):
    return (
        item["total"],
        item["count_5"],
        item["count_3"],
        item["tiebreak_total"],
    )


def build_round_ranking(event: Event, round_no: int) -> List[dict]:
    athletes = Athlete.query.filter_by(event_id=event.id).order_by(Athlete.start_order).all()
    rows = []
    round2_display_map = {}
    if round_no == 2:
        round1_rows = build_round_ranking(event, 1)
        round2_source_rows = [
            row for row in round1_rows
            if row["rank"] > event.direct_qualifiers and is_round_two_candidate(event, row["athlete"])
        ]
        round2_source_rows.sort(key=lambda row: (
            row["total"],
            row["count_5"],
            row["count_3"],
            row["tiebreak_total"],
            row["athlete"].start_order,
        ))
        lane_count = max(event.lane_count or 1, 1)
        for idx, source_row in enumerate(round2_source_rows, start=1):
            round2_display_map[source_row["athlete"].id] = {
                "display_order": idx,
                "display_lane_no": ((idx - 1) % lane_count) + 1,
                "display_lane_order": ((idx - 1) // lane_count) + 1,
            }

    for athlete in athletes:
        if round_no == 2:
            sig2 = get_round_signature(athlete.id, 2)
            if not sig2:
                continue
        ensure_round_entries(athlete.id, round_no)
        summary = summarize_round(athlete.id, round_no)
        row = {
            "athlete": athlete,
            "round_no": round_no,
            "total": summary["total"],
            "count_5": summary["count_5"],
            "count_3": summary["count_3"],
            "tiebreak_total": summary["tiebreak_total"],
            "status": athlete_round_status(athlete, round_no),
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
    rank = 1
    for idx, row in enumerate(rows):
        if idx == 0:
            row["rank"] = 1
            continue
        prev = rows[idx - 1]
        if ranking_key(row) == ranking_key(prev):
            row["rank"] = prev["rank"]
        else:
            row["rank"] = idx + 1
    return rows


def is_round_two_candidate(event: Event, athlete: Athlete) -> bool:
    if not event.has_round_two or not event.round_two_cutoff_rank:
        return False
    round1_rows = build_round_ranking(event, 1)
    if not round1_rows:
        return False
    cutoff = event.round_two_cutoff_rank
    direct = event.direct_qualifiers
    if cutoff <= direct:
        return False
    eligible_rows = [r for r in round1_rows if r["rank"] > direct]
    if not eligible_rows:
        return False
    target_score = None
    for row in round1_rows:
        if row["rank"] == cutoff:
            target_score = row["total"]
            break
    if target_score is None:
        return False
    athlete_row = next((r for r in round1_rows if r["athlete"].id == athlete.id), None)
    if not athlete_row or athlete_row["rank"] <= direct:
        return False
    return athlete_row["total"] >= target_score


def sync_round_two_candidates(event: Event) -> None:
    if not event.has_round_two:
        return

    round1_rows = build_round_ranking(event, 1)
    if not round1_rows:
        return

    direct_ids = {row["athlete"].id for row in round1_rows if row["rank"] <= event.direct_qualifiers}
    candidate_ids = {
        row["athlete"].id
        for row in round1_rows
        if row["athlete"].id not in direct_ids and is_round_two_candidate(event, row["athlete"])
    }

    for athlete in event.athletes:
        sig2 = get_round_signature(athlete.id, 2)
        if athlete.id in candidate_ids:
            ensure_round_entries(athlete.id, 2)
            sig2 = ensure_signature(athlete.id, 2)
            if not sig2.started_at and not sig2.finished_at:
                athlete.status = "waiting"
        elif sig2 and not sig2.started_at and not sig2.finished_at:
            ScoreEntry.query.filter_by(athlete_id=athlete.id, round_no=2).delete()
            TieBreakEntry.query.filter_by(athlete_id=athlete.id, round_no=2).delete()
            ScoreSignature.query.filter_by(athlete_id=athlete.id, round_no=2).delete()

    db.session.commit()

def build_round_two_start_list(event: Event) -> list[dict]:
    round1_rows = build_round_ranking(event, 1)

    direct_ids = {
        r["athlete"].id
        for r in round1_rows
        if r["rank"] <= event.direct_qualifiers
    }

    candidates = [
        r for r in round1_rows
        if is_round_two_candidate(event, r["athlete"])
        and r["athlete"].id not in direct_ids
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

def build_combined_qualifiers(event: Event) -> List[dict]:
    round1_rows = build_round_ranking(event, 1)
    direct_rows = [r for r in round1_rows if r["rank"] <= event.direct_qualifiers]
    if not event.has_round_two:
        return direct_rows

    round2_rows = build_round_ranking(event, 2)
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
            "tiebreak_total": (round1_row["tiebreak_total"] if round1_row else 0) + row["tiebreak_total"],
            "round1_total": round1_total,
            "round2_total": row["total"],
        })
    combined_rows.sort(key=ranking_key, reverse=True)
    for idx, row in enumerate(combined_rows):
        if idx == 0:
            row["rank"] = 1
        else:
            prev = combined_rows[idx - 1]
            row["rank"] = prev["rank"] if ranking_key(row) == ranking_key(prev) else idx + 1
    for idx, row in enumerate(combined_rows[:event.round_two_advancers]):
        row["seed"] = event.direct_qualifiers + idx + 1
    return direct_rows + combined_rows


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
            3: {"top": 512},  # รอบ 8 คน
            4: {"top": 587},  # รอบรองชนะเลิศ
            5: {"top": 662},  # รอบชิงชนะเลิศ
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
        },
    }



def get_progression_groups(event: Event) -> dict:
    round1_rows = build_round_ranking(event, 1)
    direct_ids = {r["athlete"].id for r in round1_rows if r["rank"] <= event.direct_qualifiers}
    round2_candidate_ids = set()
    passed_round2_ids = set()
    eliminated_ids = set()
    if event.has_round_two:
        candidates = [r for r in round1_rows if is_round_two_candidate(event, r["athlete"])]
        round2_candidate_ids = {r["athlete"].id for r in candidates} - direct_ids
        combined = build_combined_qualifiers(event)
        advancers = combined[event.direct_qualifiers:event.direct_qualifiers + event.round_two_advancers]
        passed_round2_ids = {r["athlete"].id for r in advancers}
    for athlete in event.athletes:
        aid = athlete.id
        if aid in direct_ids:
            continue
        if aid in round2_candidate_ids and aid not in passed_round2_ids:
            eliminated_ids.add(aid)
    return {"direct": direct_ids, "round2_candidates": round2_candidate_ids, "round2_passed": passed_round2_ids, "eliminated": eliminated_ids}


def compute_round_ranks(event: Event) -> dict[int, dict[int,int]]:
    result = {}
    for rn in [1,2]:
        result[rn] = {}
        for row in build_round_ranking(event, rn):
            result[rn][row["athlete"].id] = row["rank"]
    return result


def ensure_bracket(event: Event) -> list[BracketMatch]:
    matches = BracketMatch.query.filter_by(event_id=event.id).order_by(BracketMatch.round_name, BracketMatch.match_no).all()
    if matches:
        return matches
    qualifiers = build_combined_qualifiers(event)
    seeds = qualifiers[:event.direct_qualifiers + event.round_two_advancers]
    if len(seeds) >= 8:
        order = [(1,8),(4,5),(3,6),(2,7)]
        for idx,(a,b) in enumerate(order, start=1):
            arow = seeds[a-1] if len(seeds) >= a else None
            brow = seeds[b-1] if len(seeds) >= b else None
            db.session.add(BracketMatch(event_id=event.id, round_name="QF", match_no=idx, athlete_a_id=arow["athlete"].id if arow else None, athlete_b_id=brow["athlete"].id if brow else None))
    elif len(seeds) >= 4:
        order = [(1,4),(2,3)]
        for idx,(a,b) in enumerate(order, start=1):
            arow = seeds[a-1] if len(seeds) >= a else None
            brow = seeds[b-1] if len(seeds) >= b else None
            db.session.add(BracketMatch(event_id=event.id, round_name="SF", match_no=idx, athlete_a_id=arow["athlete"].id if arow else None, athlete_b_id=brow["athlete"].id if brow else None))
    db.session.commit()
    return BracketMatch.query.filter_by(event_id=event.id).order_by(BracketMatch.round_name, BracketMatch.match_no).all()


def maybe_advance_bracket(event: Event) -> None:
    matches = BracketMatch.query.filter_by(event_id=event.id).all()
    qf = sorted([m for m in matches if m.round_name == "QF"], key=lambda m: m.match_no)
    sf = sorted([m for m in matches if m.round_name == "SF"], key=lambda m: m.match_no)
    fn = sorted([m for m in matches if m.round_name == "F"], key=lambda m: m.match_no)
    changed = False
    if qf and all(m.winner_id for m in qf) and not sf:
        pairings = [(qf[0].winner_id, qf[1].winner_id), (qf[2].winner_id, qf[3].winner_id)]
        for idx,(a,b) in enumerate(pairings, start=1):
            db.session.add(BracketMatch(event_id=event.id, round_name="SF", match_no=idx, athlete_a_id=a, athlete_b_id=b))
        changed = True
    if sf and all(m.winner_id for m in sf) and not fn:
        db.session.add(BracketMatch(event_id=event.id, round_name="F", match_no=1, athlete_a_id=sf[0].winner_id, athlete_b_id=sf[1].winner_id))
        changed = True
    if changed:
        db.session.commit()


def bracket_round_to_scorecard_round(round_name: str) -> int:
    return {"QF": 3, "SF": 4, "F": 5}[round_name]


def sync_match_winner_from_scores(match: BracketMatch) -> None:
    round_no = bracket_round_to_scorecard_round(match.round_name)
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
        base = {"athlete": None, "team": "-", "name": "-", "r1": "-", "r1r2": "-", "stations": ["-"]*5, "total": "-", "seed": ""}
        return base
    r1 = summarize_round(athlete.id, 1)["total"]
    r2 = summarize_round(athlete.id, 2)["total"] if event.has_round_two else 0
    round_no = bracket_round_to_scorecard_round(round_name)
    current = summarize_round(athlete.id, round_no)
    return {
        "athlete": athlete,
        "team": athlete.affiliation,
        "name": athlete.name,
        "r1": r1,
        "r1r2": r1 + r2,
        "stations": [current["by_station"][station]["total"] for station in STATIONS],
        "total": current["total"],
        "seed": seed_map.get(athlete.id, ""),
        "round_no": round_no,
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
    return {"now": datetime.now(), "MAX_RED_CARDS": MAX_RED_CARDS}


@app.route("/")
def index():
    stats = dashboard_stats()
    events = Event.query.order_by(Event.competition_date.desc(), Event.id.desc()).all()
    return render_template("index.html", events=events, stats=stats)


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        user = User.query.filter_by(username=username).first()
        if user and user.check_password(password):
            login_user(user)
            flash("เข้าสู่ระบบสำเร็จ", "success")
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
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "").strip()
        role = request.form.get("role", "user")
        if username and password and not User.query.filter_by(username=username).first():
            user = User(username=username, role=role)
            user.set_password(password)
            db.session.add(user)
            db.session.commit()
            flash("สร้างผู้ใช้สำเร็จ", "success")
        else:
            flash("สร้างผู้ใช้ไม่สำเร็จ กรุณาตรวจสอบข้อมูล", "danger")
    users = User.query.order_by(User.id).all()
    return render_template("users.html", users=users)


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
            next_round_label=request.form["next_round_label"],
            round_two_advancers=int(request.form.get("round_two_advancers") or 4),
            created_by=current_user.id,
        )
        db.session.add(event)
        db.session.commit()
        flash("สร้างอีเวนต์สำเร็จ", "success")
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
        event.next_round_label = request.form["next_round_label"]
        event.round_two_advancers = int(request.form.get("round_two_advancers") or 4)
        recalculate_event_orders(event)
        db.session.commit()
        flash("แก้ไขอีเวนต์สำเร็จ", "success")
        return redirect(url_for("event_overview", event_id=event.id, round=1))
    return render_template("event_form.html", event=event, is_edit=True)


@app.route("/events/<int:event_id>/delete", methods=["POST"])
@login_required
@role_required("admin", "superadmin")
def delete_event(event_id: int):
    event = Event.query.get_or_404(event_id)
    BracketMatch.query.filter_by(event_id=event.id).delete()
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
    return render_template("athletes.html", event=event, athletes=athletes)


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
    import random

    event = Event.query.get_or_404(event_id)
    athletes = Athlete.query.filter_by(event_id=event.id).all()
    random.shuffle(athletes)
    for idx, athlete in enumerate(athletes, start=1):
        athlete.start_order = idx
        athlete.lane_no = ((idx - 1) % event.lane_count) + 1
        athlete.lane_order = ((idx - 1) // event.lane_count) + 1
    db.session.commit()
    flash("สุ่มลำดับใหม่แล้ว", "success")
    return redirect(url_for("manage_athletes", event_id=event.id))


@app.route("/events/<int:event_id>/overview")
def event_overview(event_id: int):
    event = Event.query.get_or_404(event_id)
    round_no = int(request.args.get("round", 1))
    if round_no == 2 and event.has_round_two:
        sync_round_two_candidates(event)
    rows = build_round_ranking(event, round_no)
    groups = get_progression_groups(event)
    for row in rows:
        aid = row["athlete"].id
        row["progress_class"] = ("qualified-direct" if aid in groups["direct"] else ("qualified-round2" if aid in groups["round2_passed"] else ("round2-candidate" if aid in groups["round2_candidates"] else ("eliminated" if aid in groups["eliminated"] else ""))))
    combined_rows = build_combined_qualifiers(event) if round_no == 2 and event.has_round_two else []
    return render_template(
        "overview.html",
        event=event,
        round_no=round_no,
        rows=rows,
        combined_rows=combined_rows,
        theme=event_theme(event.category),
        station_images=[f"station_{i}.png" for i in STATIONS],
        
    )


@app.route("/events/<int:event_id>/overview-data")
def overview_data(event_id: int):
    event = Event.query.get_or_404(event_id)
    round_no = int(request.args.get("round", 1))
    if round_no == 2 and event.has_round_two:
        sync_round_two_candidates(event)
    rows = build_round_ranking(event, round_no)
    groups = get_progression_groups(event)
    payload = []
    for row in rows:
        stations = {}
        for station in STATIONS:
            stations[str(station)] = {
                "6": row["by_station"][station]["distances"].get(6, 0),
                "7": row["by_station"][station]["distances"].get(7, 0),
                "8": row["by_station"][station]["distances"].get(8, 0),
                "9": row["by_station"][station]["distances"].get(9, 0),
                "total": row["by_station"][station]["total"],
            }
        aid = row["athlete"].id
        payload.append({
            "rank": row["rank"],
            "name": row["athlete"].name,
            "affiliation": row["athlete"].affiliation,
            "start_order": row["athlete"].start_order,
            "lane_no": row["athlete"].lane_no,
            "lane_order": row["athlete"].lane_order,
            "display_order": row.get("display_order", row["athlete"].start_order),
            "display_lane_no": row.get("display_lane_no", row["athlete"].lane_no),
            "display_lane_order": row.get("display_lane_order", row["athlete"].lane_order),
            "status": row["status"],
            "total": row["total"],
            "athlete_id": row["athlete"].id,
            "progress_class": ("qualified-direct" if aid in groups["direct"] else ("qualified-round2" if aid in groups["round2_passed"] else ("round2-candidate" if aid in groups["round2_candidates"] else ("eliminated" if aid in groups["eliminated"] else "")))),
            "stations": stations,
        })
    return jsonify(payload)


@app.route("/api/scorecard/<int:athlete_id>/autosave", methods=["POST"])
@login_required
@role_required("admin", "superadmin")
def autosave_scorecard(athlete_id: int):
    athlete = Athlete.query.get_or_404(athlete_id)
    payload = request.get_json() or {}
    round_no = int(payload.get("round_no", 1))
    station_no = int(payload.get("station_no", 1))
    distance_m = int(payload.get("distance_m", 6))
    score_value = str(payload.get("score", "")).strip()
    red = bool(payload.get("red", False))
    ensure_round_entries(athlete.id, round_no)
    signature = ensure_signature(athlete.id, round_no)
    if not signature.started_at:
        signature.started_at = datetime.utcnow()
    athlete.status = "active"
    entry = ScoreEntry.query.filter_by(athlete_id=athlete.id, round_no=round_no, station_no=station_no, distance_m=distance_m).first()
    if entry:
        value = 0 if score_value == "" else int(score_value)
        value = max(0, min(5, value))
        entry.is_red_card = red
        entry.score = 0 if red else value
    db.session.commit()
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
        "round_total": summary["total"]
    })


@app.route("/athletes/<int:athlete_id>/scorecard", methods=["GET", "POST"])
@login_required
@role_required("admin", "superadmin")
def scorecard(athlete_id: int):
    athlete = Athlete.query.get_or_404(athlete_id)
    event = athlete.event
    round_no = int(request.args.get("round", 1))
    if round_no == 2 and not is_round_two_candidate(event, athlete):
        flash("นักกีฬาคนนี้ไม่มีสิทธิ์ตีรอบ 2", "warning")
        return redirect(url_for("event_overview", event_id=event.id, round=1))

    ensure_round_entries(athlete.id, 1)
    if round_no == 2:
        ensure_round_entries(athlete.id, 2)

    signature = ensure_signature(athlete.id, round_no)
    if not signature.started_at:
        signature.started_at = datetime.utcnow()
        athlete.status = "active"
        db.session.commit()

    if request.method == "POST":
        referee_name_key = f"referee_name_{round_no}"
        recorder_name_key = f"recorder_name_{round_no}"
        athlete_name_key = f"athlete_name_{round_no}"
        referee_sig_key = f"ref_sig_{round_no}"
        recorder_sig_key = f"rec_sig_{round_no}"
        athlete_sig_key = f"ath_sig_{round_no}"

        signature.recorder_name = request.form.get(recorder_name_key, "").strip() or signature.recorder_name
        signature.referee_name = request.form.get(referee_name_key, "").strip() or signature.referee_name
        signature.athlete_name = request.form.get(athlete_name_key, "").strip() or signature.athlete_name

        signature.recorder_signature = request.form.get(recorder_sig_key, "").strip() or signature.recorder_signature
        signature.referee_signature = request.form.get(referee_sig_key, "").strip() or signature.referee_signature
        signature.athlete_signature = request.form.get(athlete_sig_key, "").strip() or signature.athlete_signature

        bypass_code = request.form.get("bypass_code", "").strip()
        bypass_ok = current_user.role == "superadmin" or (current_user.role == "admin" and bypass_code == "7929")
        signed_ok = all([
            bool(signature.recorder_name or signature.recorder_signature),
            bool(signature.referee_name or signature.referee_signature),
            bool(signature.athlete_name or signature.athlete_signature),
        ])

        if bypass_ok or signed_ok:
            signature.bypass_signed = bypass_ok
            signature.finished_at = datetime.utcnow()
            athlete.status = "finished"
            db.session.commit()
            if round_no == 1 and event.has_round_two:
                sync_round_two_candidates(event)
                reset_event_bracket(event)
            elif round_no == 2 and event.has_round_two:
                reset_event_bracket(event)
            elif round_no in [3, 4, 5]:
                round_map = {3: "QF", 4: "SF", 5: "F"}
                match = BracketMatch.query.filter_by(
                    event_id=event.id,
                    round_name=round_map[round_no]
                ).filter(
                    (BracketMatch.athlete_a_id == athlete.id) | (BracketMatch.athlete_b_id == athlete.id)
                ).first()
                if match:
                    sync_match_winner_from_scores(match)
                    maybe_advance_bracket(event)
                flash("จบการตีเรียบร้อย", "success")
                return redirect(url_for("bracket", event_id=event.id))
            flash("จบการตีเรียบร้อย", "success")
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
    }

    round_station_running_totals = {}
    for rn in [1, 2, 3, 4, 5]:
        running = {}
        acc = 0
        for st in [1, 2, 3, 4, 5]:
            val = template_data["station_totals"].get((rn, st), 0)
            acc += val
            running[st] = acc
        round_station_running_totals[rn] = running

    round_signatures = {}
    for rn in [1, 2, 3, 4, 5]:
        round_signatures[rn] = get_round_signature(athlete.id, rn)

    combined_rows = build_combined_qualifiers(event) if event.has_round_two else []
    current_combined = next((r for r in combined_rows if r.get("athlete") and r["athlete"].id == athlete.id), None)
    display_order = athlete.start_order
    display_lane_no = athlete.lane_no
    display_lane_order = athlete.lane_order
    current_round_rows = build_round_ranking(event, round_no)
    current_row = next((row for row in current_round_rows if row["athlete"].id == athlete.id), None)
    if current_row:
        display_order = current_row.get("display_order", display_order)
        display_lane_no = current_row.get("display_lane_no", display_lane_no)
        display_lane_order = current_row.get("display_lane_order", display_lane_order)

    return render_template(
        "scorecard.html",
        athlete=athlete,
        event=event,
        round_no=round_no,
        round_labels=SCORECARD_ROUND_LABELS,
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
    )


def build_scorecard_print_context(athlete: Athlete, round_no: int) -> dict:
    event = athlete.event

    template_data = build_scorecard_template_data(athlete.id)
    ranks = compute_round_ranks(event)

    round_ranks = {
        1: ranks.get(1, {}).get(athlete.id, ""),
        2: ranks.get(2, {}).get(athlete.id, ""),
        3: "",
        4: "",
        5: "",
    }

    round_station_running_totals = {}
    for rn in [1, 2, 3, 4, 5]:
        running = {}
        acc = 0
        for st in STATIONS:
            val = template_data["station_totals"].get((rn, st), 0)
            acc += val
            running[st] = acc
        round_station_running_totals[rn] = running

    round_signatures = {}
    for rn in [1, 2, 3, 4, 5]:
        round_signatures[rn] = get_round_signature(athlete.id, rn)

    combined_rows = build_combined_qualifiers(event) if event.has_round_two else []
    current_combined = next(
        (r for r in combined_rows if r.get("athlete") and r["athlete"].id == athlete.id),
        None
    )

    display_order = athlete.start_order
    display_lane_no = athlete.lane_no
    display_lane_order = athlete.lane_order

    current_round_rows = build_round_ranking(event, round_no)
    current_row = next((row for row in current_round_rows if row["athlete"].id == athlete.id), None)
    if current_row:
        display_order = current_row.get("display_order", display_order)
        display_lane_no = current_row.get("display_lane_no", display_lane_no)
        display_lane_order = current_row.get("display_lane_order", display_lane_order)

    return {
        "athlete": athlete,
        "event": event,
        "round_no": round_no,
        "round_labels": SCORECARD_ROUND_LABELS,
        "score_map": template_data["score_map"],
        "station_totals": template_data["station_totals"],
        "station_reds": template_data["station_reds"],
        "round_totals": template_data["round_totals"],
        "round_ranks": round_ranks,
        "round_signatures": round_signatures,
        "round_station_running_totals": round_station_running_totals,
        "current_combined": current_combined,
        "combined_rows": combined_rows,
        "display_order": display_order,
        "display_lane_no": display_lane_no,
        "display_lane_order": display_lane_order,
        "positions": scorecard_print_positions(),
        "station_images": [f"station_{i}.png" for i in STATIONS],
    }


@app.route("/events/<int:event_id>/scorecards-print-select", methods=["GET", "POST"])
@login_required
@role_required("admin", "superadmin")
def scorecards_print_select(event_id: int):
    event = Event.query.get_or_404(event_id)
    round_no = int(request.values.get("round", 1))

    athletes = Athlete.query.filter_by(event_id=event.id).order_by(
        Athlete.start_order,
        Athlete.id
    ).all()

    if round_no == 2 and event.has_round_two:
        athletes = [a for a in athletes if is_round_two_candidate(event, a)]

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
        round_labels=SCORECARD_ROUND_LABELS,
    )


@app.route("/events/<int:event_id>/scorecards-print-bulk")
@login_required
@role_required("admin", "superadmin")
def scorecards_print_bulk(event_id: int):
    event = Event.query.get_or_404(event_id)
    round_no = int(request.args.get("round", 1))

    ids_text = request.args.get("ids", "").strip()
    selected_ids = []
    if ids_text:
        for raw_id in ids_text.split(","):
            raw_id = raw_id.strip()
            if raw_id.isdigit():
                selected_ids.append(int(raw_id))

    query = Athlete.query.filter_by(event_id=event.id)
    if selected_ids:
        query = query.filter(Athlete.id.in_(selected_ids))

    athletes = query.order_by(Athlete.start_order, Athlete.id).all()

    if round_no == 2 and event.has_round_two:
        athletes = [a for a in athletes if is_round_two_candidate(event, a)]

    print_items = [
        build_scorecard_print_context(athlete, round_no)
        for athlete in athletes
    ]

    return render_template(
        "scorecards_print_bulk.html",
        event=event,
        round_no=round_no,
        round_labels=SCORECARD_ROUND_LABELS,
        print_items=print_items,
        station_images=[f"station_{i}.png" for i in STATIONS],
    )

@app.route("/athletes/<int:athlete_id>/scorecard-print")
@login_required
@role_required("admin", "superadmin")
def scorecard_print(athlete_id: int):
    athlete = Athlete.query.get_or_404(athlete_id)
    event = athlete.event
    round_no = int(request.args.get("round", 1))

    if round_no == 2 and event.has_round_two and not is_round_two_candidate(event, athlete):
        flash("นักกีฬาคนนี้ไม่มีสิทธิ์ตีรอบ 2", "warning")
        return redirect(url_for("event_overview", event_id=event.id, round=1))

    template_data = build_scorecard_template_data(athlete.id)
    ranks = compute_round_ranks(event)

    round_ranks = {
        1: ranks.get(1, {}).get(athlete.id, ""),
        2: ranks.get(2, {}).get(athlete.id, ""),
        3: "",
        4: "",
        5: "",
    }

    round_station_running_totals = {}
    for rn in [1, 2, 3, 4, 5]:
        running = {}
        acc = 0
        for st in STATIONS:
            val = template_data["station_totals"].get((rn, st), 0)
            acc += val
            running[st] = acc
        round_station_running_totals[rn] = running

    round_signatures = {}
    for rn in [1, 2, 3, 4, 5]:
        round_signatures[rn] = get_round_signature(athlete.id, rn)

    combined_rows = build_combined_qualifiers(event) if event.has_round_two else []
    current_combined = next(
        (r for r in combined_rows if r.get("athlete") and r["athlete"].id == athlete.id),
        None
    )

    display_order = athlete.start_order
    display_lane_no = athlete.lane_no
    display_lane_order = athlete.lane_order
    current_round_rows = build_round_ranking(event, round_no)
    current_row = next((row for row in current_round_rows if row["athlete"].id == athlete.id), None)
    if current_row:
        display_order = current_row.get("display_order", display_order)
        display_lane_no = current_row.get("display_lane_no", display_lane_no)
        display_lane_order = current_row.get("display_lane_order", display_lane_order)

    positions = scorecard_print_positions()

    return render_template(
        "scorecard_print.html",
        athlete=athlete,
        event=event,
        round_no=round_no,
        round_labels=SCORECARD_ROUND_LABELS,
        score_map=template_data["score_map"],
        station_totals=template_data["station_totals"],
        station_reds=template_data["station_reds"],
        round_totals=template_data["round_totals"],
        round_ranks=round_ranks,
        round_signatures=round_signatures,
        round_station_running_totals=round_station_running_totals,
        current_combined=current_combined,
        combined_rows=combined_rows,
        display_order=display_order,
        display_lane_no=display_lane_no,
        display_lane_order=display_lane_order,
        positions=positions,
        station_images=[f"station_{i}.png" for i in STATIONS],
    )

@app.route("/athletes/<int:athlete_id>/activate", methods=["POST"])
@login_required
@role_required("admin", "superadmin")
def activate_scorecard(athlete_id: int):
    athlete = Athlete.query.get_or_404(athlete_id)
    round_no = int(request.args.get("round", 1))
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
    round_no = int(request.args.get("round", 1))
    if request.method == "POST":
        TieBreakEntry.query.filter_by(athlete_id=athlete.id, round_no=round_no).delete()
        for station_no in STATIONS:
            score = int(request.form.get(f"tb_{station_no}", 0) or 0)
            db.session.add(TieBreakEntry(athlete_id=athlete.id, round_no=round_no, station_no=station_no, score=score))
        db.session.commit()
        flash("บันทึกผลเที่ยวพิเศษแล้ว", "success")
        return redirect(url_for("event_overview", event_id=athlete.event_id, round=round_no))
    existing = {e.station_no: e.score for e in TieBreakEntry.query.filter_by(athlete_id=athlete.id, round_no=round_no).all()}
    return render_template("tiebreak.html", athlete=athlete, round_no=round_no, existing=existing)


@app.route("/events/<int:event_id>/bracket")
def bracket(event_id: int):
    event = Event.query.get_or_404(event_id)
    qualifiers = build_combined_qualifiers(event)
    ensure_bracket(event)
    maybe_advance_bracket(event)
    matches = BracketMatch.query.filter_by(event_id=event.id).order_by(BracketMatch.round_name, BracketMatch.match_no).all()
    athlete_map = {a.id: a for a in event.athletes}
    seed_map = {row["athlete"].id: idx for idx, row in enumerate(qualifiers, start=1) if row.get("athlete")}
    grouped = {"QF": [], "SF": [], "F": []}
    for m in matches:
        a = athlete_map.get(m.athlete_a_id)
        b = athlete_map.get(m.athlete_b_id)
        grouped.setdefault(m.round_name, []).append({
            "match": m,
            "a": build_bracket_match_row(event, a, m.round_name, seed_map),
            "b": build_bracket_match_row(event, b, m.round_name, seed_map),
            "winner": athlete_map.get(m.winner_id),
            "round_no": bracket_round_to_scorecard_round(m.round_name),
        })
    return render_template("bracket.html", event=event, grouped=grouped)


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
def bracket_data(event_id: int):
    event = Event.query.get_or_404(event_id)

    maybe_advance_bracket(event)

    matches = BracketMatch.query.filter_by(event_id=event.id).order_by(
        BracketMatch.round_name,
        BracketMatch.match_no
    ).all()

    athlete_map = {a.id: a for a in event.athletes}

    grouped = {"QF": [], "SF": [], "F": []}
    for m in matches:
        grouped.setdefault(m.round_name, []).append({
            "match_no": m.match_no,
            "winner_id": m.winner_id,
            "a": build_bracket_row_data(event, athlete_map.get(m.athlete_a_id), m.round_name),
            "b": build_bracket_row_data(event, athlete_map.get(m.athlete_b_id), m.round_name),
        })

    return jsonify(grouped)

@app.route("/events/<int:event_id>/stats")
def event_stats(event_id: int):
    event = Event.query.get_or_404(event_id)
    round1_rows = build_round_ranking(event, 1)
    best = round1_rows[0] if round1_rows else None
    return render_template("stats.html", event=event, best=best, round1_rows=round1_rows)


if __name__ == "__main__":
    os.makedirs(os.path.join(BASE_DIR, "instance"), exist_ok=True)
    with app.app_context():
        db.create_all()
        ensure_schema()
        seed_defaults()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))