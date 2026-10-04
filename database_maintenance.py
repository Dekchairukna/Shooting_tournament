"""Explicit, non-destructive schema maintenance (never run on app import)."""
from sqlalchemy import inspect, text


def harden_schema(db):
    unique_indexes = {
        "uq_score_cell": ("score_entry", "athlete_id, round_no, station_no, distance_m"),
        "uq_score_signature": ("score_signature", "athlete_id, round_no"),
        "uq_bracket_match": ("bracket_match", "event_id, round_name, match_no"),
    }
    with db.engine.begin() as conn:
        # Refuse ambiguous old data rather than silently deleting scores.
        for name, (table, columns) in unique_indexes.items():
            duplicates = conn.execute(text(
                f"SELECT {columns}, COUNT(*) FROM {table} GROUP BY {columns} HAVING COUNT(*) > 1 LIMIT 5"
            )).fetchall()
            if duplicates:
                raise RuntimeError(f"Duplicate records in {table}; review before migration: {duplicates}")
        if "score_version" not in {c["name"] for c in inspect(conn).get_columns("athlete")}:
            conn.execute(text("ALTER TABLE athlete ADD COLUMN score_version INTEGER NOT NULL DEFAULT 0"))
        for name, (table, columns) in unique_indexes.items():
            conn.execute(text(f"CREATE UNIQUE INDEX IF NOT EXISTS {name} ON {table} ({columns})"))
        for name, table, columns in (
            ("ix_athlete_event", "athlete", "event_id"),
            ("ix_tiebreak_round", "tie_break_entry", "athlete_id, round_no"),
            ("ix_score_log_round", "score_edit_log", "athlete_id, round_no"),
            ("ix_court_event", '"user"', "court_event_id, court_no"),
        ):
            conn.execute(text(f"CREATE INDEX IF NOT EXISTS {name} ON {table} ({columns})"))
