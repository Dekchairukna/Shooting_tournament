#!/usr/bin/env python3
"""
Import the two local DEMO events from SQLite into the website's PostgreSQL database.

Safe defaults:
- Imports only events whose event_group == "DEMO" or whose name starts with "DEMO •"
- Does NOT overwrite an existing event with the same name unless --replace is used.
- Remaps Event/Athlete IDs so existing PostgreSQL IDs are never overwritten.
- Does not copy local user IDs into PostgreSQL.

Recommended local usage with Railway:
    railway run python import_demo_to_postgres.py \
      --sqlite instance/shooting.db \
      --use-public-url

If the demo events already exist and you want to replace them:
    railway run python import_demo_to_postgres.py \
      --sqlite instance/shooting.db \
      --use-public-url \
      --replace
"""
from __future__ import annotations

import argparse
import os
import sqlite3
import sys
from pathlib import Path


def choose_target_url(use_public_url: bool) -> str | None:
    """Choose a PostgreSQL URL supplied by Railway or the shell."""
    if os.environ.get("TARGET_DATABASE_URL"):
        return os.environ["TARGET_DATABASE_URL"]

    if use_public_url and os.environ.get("DATABASE_PUBLIC_URL"):
        return os.environ["DATABASE_PUBLIC_URL"]

    return os.environ.get("DATABASE_URL") or os.environ.get("DATABASE_PUBLIC_URL")


def normalize_pg_url(url: str) -> str:
    if url.startswith("postgres://"):
        return "postgresql://" + url[len("postgres://"):]
    return url


def qrows(conn: sqlite3.Connection, sql: str, params=()):
    return [dict(row) for row in conn.execute(sql, params).fetchall()]


def qrow(conn: sqlite3.Connection, sql: str, params=()):
    row = conn.execute(sql, params).fetchone()
    return dict(row) if row else None


def table_columns(conn: sqlite3.Connection, table: str) -> set[str]:
    return {row["name"] for row in conn.execute(f"PRAGMA table_info({table})").fetchall()}


def getv(row: dict, key: str, default=None):
    value = row.get(key, default)
    return default if value is None and default is not None else value


def main() -> int:
    parser = argparse.ArgumentParser(description="Import local demo events from SQLite to PostgreSQL.")
    parser.add_argument("--sqlite", default="instance/shooting.db", help="Source SQLite DB path.")
    parser.add_argument("--replace", action="store_true",
                        help="Delete same-named DEMO events in PostgreSQL before importing.")
    parser.add_argument("--use-public-url", action="store_true",
                        help="Prefer DATABASE_PUBLIC_URL (recommended when running locally).")
    parser.add_argument("--dry-run", action="store_true",
                        help="Show which demo events would be imported without writing PostgreSQL.")
    args = parser.parse_args()

    sqlite_path = Path(args.sqlite).expanduser().resolve()
    if not sqlite_path.exists():
        print(f"ERROR: SQLite file not found: {sqlite_path}")
        return 2

    src = sqlite3.connect(str(sqlite_path))
    src.row_factory = sqlite3.Row

    events = qrows(
        src,
        """
        SELECT *
        FROM event
        WHERE event_group = 'DEMO' OR name LIKE 'DEMO •%'
        ORDER BY id
        """
    )

    if not events:
        print("No DEMO events were found in the SQLite database.")
        return 3

    print("DEMO events found in SQLite:")
    for e in events:
        athletes = qrow(src, "SELECT COUNT(*) AS n FROM athlete WHERE event_id=?", (e["id"],))["n"]
        scores = qrow(
            src,
            """
            SELECT COUNT(*) AS n
            FROM score_entry se
            JOIN athlete a ON a.id = se.athlete_id
            WHERE a.event_id=?
            """,
            (e["id"],)
        )["n"]
        matches = qrow(src, "SELECT COUNT(*) AS n FROM bracket_match WHERE event_id=?", (e["id"],))["n"]
        print(f"  - [{e['id']}] {e['name']} | athletes={athletes}, scores={scores}, bracket={matches}")

    if args.dry_run:
        print("\nDry run only. Nothing was written.")
        return 0

    target_url = choose_target_url(args.use_public_url)
    if not target_url:
        print(
            "\nERROR: PostgreSQL URL not found.\n"
            "Use Railway CLI (railway run ...) or set TARGET_DATABASE_URL / DATABASE_PUBLIC_URL."
        )
        return 4

    target_url = normalize_pg_url(target_url)
    if not target_url.startswith(("postgresql://", "postgresql+")):
        print("ERROR: Target URL is not PostgreSQL. Refusing to continue.")
        return 5

    # Make app.py point to the target PostgreSQL database before importing it.
    os.environ["DATABASE_URL"] = target_url

    # Import the application's current SQLAlchemy models.
    try:
        from app import (
            app, db,
            User, Event, Athlete, ScoreEntry, ScoreSignature,
            ScoreEditLog, TieBreakEntry, BracketMatch, ResultsApprovedSetting,
        )
    except Exception as exc:
        print(f"ERROR importing app/models: {exc}")
        print("Run this script from the project root, next to app.py.")
        return 6

    event_cols = table_columns(src, "event")
    athlete_cols = table_columns(src, "athlete")
    score_cols = table_columns(src, "score_entry")
    sig_cols = table_columns(src, "score_signature")
    edit_cols = table_columns(src, "score_edit_log")
    tie_cols = table_columns(src, "tie_break_entry")
    bracket_cols = table_columns(src, "bracket_match")
    settings_cols = table_columns(src, "results_approved_setting")

    def src_value(row, cols, name, default=None):
        return row.get(name, default) if name in cols else default

    with app.app_context():
        if db.engine.dialect.name != "postgresql":
            print(f"ERROR: Target dialect is {db.engine.dialect.name!r}, not PostgreSQL.")
            return 7

        print("\nConnected to PostgreSQL.")

        imported = []
        skipped = []

        try:
            for e in events:
                existing = Event.query.filter_by(name=e["name"]).first()

                if existing and not args.replace:
                    print(f"SKIP: already exists: {e['name']}")
                    skipped.append(e["name"])
                    continue

                if existing and args.replace:
                    print(f"REPLACE: deleting existing destination event: {e['name']}")
                    athlete_ids = [a.id for a in Athlete.query.filter_by(event_id=existing.id).all()]

                    # Break user -> event references first.
                    User.query.filter_by(court_event_id=existing.id).update(
                        {User.court_event_id: None},
                        synchronize_session=False,
                    )

                    if athlete_ids:
                        ScoreEditLog.query.filter(ScoreEditLog.athlete_id.in_(athlete_ids)).delete(
                            synchronize_session=False
                        )

                    ResultsApprovedSetting.query.filter_by(event_id=existing.id).delete(
                        synchronize_session=False
                    )
                    BracketMatch.query.filter_by(event_id=existing.id).delete(
                        synchronize_session=False
                    )

                    db.session.delete(existing)
                    db.session.flush()

                new_event = Event(
                    name=e["name"],
                    event_group=getv(e, "event_group", "DEMO"),
                    category=getv(e, "category", "men"),
                    competition_date=e["competition_date"],
                    location=getv(e, "location", "-"),
                    lane_count=getv(e, "lane_count", 1),
                    direct_qualifiers=getv(e, "direct_qualifiers", 0),
                    has_round_two=bool(getv(e, "has_round_two", 0)),
                    round_two_cutoff_rank=src_value(e, event_cols, "round_two_cutoff_rank", None),
                    next_round_label=getv(e, "next_round_label", "รอบ 8 คน"),
                    round_two_advancers=getv(e, "round_two_advancers", 4),
                    created_at=src_value(e, event_cols, "created_at", None),
                    created_by=None,
                )
                db.session.add(new_event)
                db.session.flush()

                old_to_new_athlete: dict[int, int] = {}

                src_athletes = qrows(
                    src,
                    "SELECT * FROM athlete WHERE event_id=? ORDER BY id",
                    (e["id"],),
                )

                for a in src_athletes:
                    na = Athlete(
                        event_id=new_event.id,
                        bib_no=a["bib_no"],
                        name=a["name"],
                        affiliation=a["affiliation"],
                        start_order=a["start_order"],
                        lane_no=a["lane_no"],
                        lane_order=a["lane_order"],
                        status=getv(a, "status", "waiting"),
                        red_card_count=getv(a, "red_card_count", 0),
                        round_two_disabled=bool(src_value(a, athlete_cols, "round_two_disabled", 0)),
                        round_two_disabled_at=src_value(a, athlete_cols, "round_two_disabled_at", None),
                        round_two_disabled_by=None,
                        created_at=src_value(a, athlete_cols, "created_at", None),
                    )
                    db.session.add(na)
                    db.session.flush()
                    old_to_new_athlete[a["id"]] = na.id

                for old_aid, new_aid in old_to_new_athlete.items():
                    for r in qrows(src, "SELECT * FROM score_entry WHERE athlete_id=? ORDER BY id", (old_aid,)):
                        db.session.add(ScoreEntry(
                            athlete_id=new_aid,
                            round_no=getv(r, "round_no", 1),
                            station_no=r["station_no"],
                            distance_m=r["distance_m"],
                            score=getv(r, "score", 0),
                            is_red_card=bool(getv(r, "is_red_card", 0)),
                            is_scored=bool(src_value(r, score_cols, "is_scored", 0)),
                            updated_at=src_value(r, score_cols, "updated_at", None),
                        ))

                    for r in qrows(src, "SELECT * FROM score_signature WHERE athlete_id=? ORDER BY id", (old_aid,)):
                        db.session.add(ScoreSignature(
                            athlete_id=new_aid,
                            round_no=r["round_no"],
                            recorder_name=src_value(r, sig_cols, "recorder_name", None),
                            referee_name=src_value(r, sig_cols, "referee_name", None),
                            athlete_name=src_value(r, sig_cols, "athlete_name", None),
                            recorder_signature=src_value(r, sig_cols, "recorder_signature", None),
                            referee_signature=src_value(r, sig_cols, "referee_signature", None),
                            athlete_signature=src_value(r, sig_cols, "athlete_signature", None),
                            bypass_signed=bool(src_value(r, sig_cols, "bypass_signed", 0)),
                            started_at=src_value(r, sig_cols, "started_at", None),
                            finished_at=src_value(r, sig_cols, "finished_at", None),
                            stopped_by_red=bool(src_value(r, sig_cols, "stopped_by_red", 0)),
                        ))

                    for r in qrows(src, "SELECT * FROM tie_break_entry WHERE athlete_id=? ORDER BY id", (old_aid,)):
                        db.session.add(TieBreakEntry(
                            athlete_id=new_aid,
                            round_no=r["round_no"],
                            station_no=r["station_no"],
                            score=getv(r, "score", 0),
                        ))

                    for r in qrows(src, "SELECT * FROM score_edit_log WHERE athlete_id=? ORDER BY id", (old_aid,)):
                        db.session.add(ScoreEditLog(
                            athlete_id=new_aid,
                            round_no=r["round_no"],
                            station_no=r["station_no"],
                            distance_m=r["distance_m"],
                            old_score=getv(r, "old_score", 0),
                            new_score=getv(r, "new_score", 0),
                            old_red=bool(getv(r, "old_red", 0)),
                            new_red=bool(getv(r, "new_red", 0)),
                            old_played=bool(src_value(r, edit_cols, "old_played", 0)),
                            new_played=bool(src_value(r, edit_cols, "new_played", 0)),
                            edited_by=None,
                            editor_username=src_value(r, edit_cols, "editor_username", None),
                            editor_court_no=src_value(r, edit_cols, "editor_court_no", None),
                            editor_signature=src_value(r, edit_cols, "editor_signature", None),
                            reason=src_value(r, edit_cols, "reason", None),
                            edited_at=src_value(r, edit_cols, "edited_at", None),
                        ))

                # Bracket matches
                for m in qrows(
                    src,
                    "SELECT * FROM bracket_match WHERE event_id=? ORDER BY id",
                    (e["id"],),
                ):
                    db.session.add(BracketMatch(
                        event_id=new_event.id,
                        round_name=m["round_name"],
                        match_no=m["match_no"],
                        athlete_a_id=old_to_new_athlete.get(m.get("athlete_a_id")),
                        athlete_b_id=old_to_new_athlete.get(m.get("athlete_b_id")),
                        winner_id=old_to_new_athlete.get(m.get("winner_id")),
                    ))

                # Results/approved page settings
                settings = qrow(
                    src,
                    "SELECT * FROM results_approved_setting WHERE event_id=?",
                    (e["id"],),
                )
                if settings:
                    db.session.add(ResultsApprovedSetting(
                        event_id=new_event.id,
                        competition_title=src_value(settings, settings_cols, "competition_title", None),
                        host_line=src_value(settings, settings_cols, "host_line", None),
                        date_line=src_value(settings, settings_cols, "date_line", None),
                        location_line=src_value(settings, settings_cols, "location_line", None),
                        country_label=src_value(settings, settings_cols, "country_label", "COUNTRY") or "COUNTRY",
                        president_title=src_value(settings, settings_cols, "president_title", None),
                        president_name=src_value(settings, settings_cols, "president_name", None),
                        technical_title=src_value(settings, settings_cols, "technical_title", None),
                        technical_name=src_value(settings, settings_cols, "technical_name", None),
                        umpires_text=src_value(settings, settings_cols, "umpires_text", None),
                        approved_text=src_value(settings, settings_cols, "approved_text", "……………………………APPROVED")
                                      or "……………………………APPROVED",
                        show_official_pages=bool(src_value(settings, settings_cols, "show_official_pages", 1)),
                        cover_main_logo_path=src_value(settings, settings_cols, "cover_main_logo_path", None),
                        cover_bottom_logo_1_path=src_value(settings, settings_cols, "cover_bottom_logo_1_path", None),
                        cover_bottom_logo_2_path=src_value(settings, settings_cols, "cover_bottom_logo_2_path", None),
                        cover_bottom_logo_3_path=src_value(settings, settings_cols, "cover_bottom_logo_3_path", None),
                        header_logo_1_path=src_value(settings, settings_cols, "header_logo_1_path", None),
                        header_logo_2_path=src_value(settings, settings_cols, "header_logo_2_path", None),
                        header_logo_3_path=src_value(settings, settings_cols, "header_logo_3_path", None),
                        header_logo_4_path=src_value(settings, settings_cols, "header_logo_4_path", None),
                        side_logo_path=src_value(settings, settings_cols, "side_logo_path", None),
                        updated_at=src_value(settings, settings_cols, "updated_at", None),
                    ))

                db.session.flush()

                n_athletes = Athlete.query.filter_by(event_id=new_event.id).count()
                n_bracket = BracketMatch.query.filter_by(event_id=new_event.id).count()
                n_scores = (
                    db.session.query(ScoreEntry)
                    .join(Athlete, Athlete.id == ScoreEntry.athlete_id)
                    .filter(Athlete.event_id == new_event.id)
                    .count()
                )

                print(
                    f"IMPORTED: {e['name']} -> PostgreSQL event_id={new_event.id} "
                    f"(athletes={n_athletes}, scores={n_scores}, bracket={n_bracket})"
                )
                imported.append(e["name"])

            db.session.commit()

        except Exception as exc:
            db.session.rollback()
            print(f"\nERROR: import rolled back: {exc}")
            raise

        print("\nDone.")
        print(f"Imported: {len(imported)}")
        print(f"Skipped:  {len(skipped)}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
