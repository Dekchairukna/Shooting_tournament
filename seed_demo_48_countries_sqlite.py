"""Seed a complete 48-country demo tournament directly into instance/shooting.db.

Safe to rerun: it deletes only the event named EVENT_NAME, then recreates it.
Demo flow:
Round 1 (48) -> Top 4 direct -> Class 5-16 Round 2 -> Top 4 from R2 ->
Quarter Final (8) -> Semi Final (4) -> Final (2) -> Champion.
"""
from __future__ import annotations

import os
import sqlite3
from datetime import datetime

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
DB_PATH = os.path.join(BASE_DIR, "instance", "shooting.db")
EVENT_NAME = "DEMO • FINISHED • 48 NATIONS • PETANQUE SHOOTING 2026"
NOW = "2026-09-15 09:00:00"
FINISHED = "2026-09-15 09:10:00"

COUNTRIES = [
    "BELGIUM", "SPAIN", "ITALY", "MADAGASCAR", "TÜRKIYE", "MONACO",
    "MALAYSIA", "SWITZERLAND", "JAPAN", "GERMANY", "NORWAY", "ENGLAND",
    "FRENCH POLYNESIA", "CANADA", "SLOVAKIA", "SWEDEN", "THAILAND", "FRANCE",
    "HUNGARY", "CAMBODIA", "VIETNAM", "TUNISIA", "NETHERLANDS", "LUXEMBOURG",
    "ALGERIA", "LITHUANIA", "AUSTRALIA", "MEXICO", "UNITED STATES", "UKRAINE",
    "ESTONIA", "FINLAND", "WALES", "SCOTLAND", "BULGARIA", "SOUTH KOREA",
    "NEW ZEALAND", "MOROCCO", "SENEGAL", "LAOS", "SINGAPORE", "INDONESIA",
    "CZECH REPUBLIC", "AUSTRIA", "POLAND", "PORTUGAL", "CHINESE TAIPEI", "ISRAEL",
]


R1_TOTALS = [
    62, 59, 57, 55, 52, 50, 48, 46, 44, 42, 40, 38, 36, 34, 32, 30,
    28, 27, 26, 25, 24, 23, 22, 21, 20, 19, 18, 17, 16, 15, 14, 13,
    12, 11, 10, 9, 8, 7, 6, 5, 4, 3, 3, 2, 2, 1, 1, 0,
]

R2_BY_COUNTRY = {
    "TÜRKIYE": 20, "MONACO": 19, "MALAYSIA": 52, "SWITZERLAND": 55,
    "JAPAN": 53, "GERMANY": 54, "NORWAY": 24, "ENGLAND": 23,
    "FRENCH POLYNESIA": 22, "CANADA": 21, "SLOVAKIA": 20, "SWEDEN": 19,
}


QF_SCORES = {
    "BELGIUM": 44, "GERMANY": 32,
    "MADAGASCAR": 40, "SWITZERLAND": 36,
    "ITALY": 39, "MALAYSIA": 34,
    "SPAIN": 45, "JAPAN": 37,
}
SF_SCORES = {"BELGIUM": 46, "MADAGASCAR": 35, "ITALY": 38, "SPAIN": 43}
FINAL_SCORES = {"BELGIUM": 48, "SPAIN": 42}



def score_vector(total: int) -> list[int]:
    vals: list[int] = []
    remaining = int(total)
    while remaining >= 5 and len(vals) < 20:
        vals.append(5)
        remaining -= 5
    if remaining == 4:
        vals += [3, 1]
    elif remaining == 3:
        vals += [3]
    elif remaining == 2:
        vals += [1, 1]
    elif remaining == 1:
        vals += [1]
    if len(vals) > 20:
        raise ValueError(total)
    vals += [0] * (20 - len(vals))
    # Spread nonzero values across 5 ateliers for a natural-looking table.
    order = [0,4,8,12,16, 1,5,9,13,17, 2,6,10,14,18, 3,7,11,15,19]
    out = [0] * 20
    for src, dst in enumerate(order):
        out[dst] = vals[src]
    return out


def next_id(cur, table: str) -> int:
    return int(cur.execute(f"SELECT COALESCE(MAX(id),0)+1 FROM {table}").fetchone()[0])


def add_round(cur, athlete_id: int, athlete_name: str, round_no: int, total: int, signed=True):
    vals = score_vector(total)
    i = 0
    for station in range(1, 6):
        for distance in (6, 7, 8, 9):
            cur.execute(
                """INSERT INTO score_entry
                   (athlete_id,round_no,station_no,distance_m,score,is_red_card,updated_at,is_scored)
                   VALUES (?,?,?,?,?,?,?,?)""",
                (athlete_id, round_no, station, distance, vals[i], 0, FINISHED, 1),
            )
            i += 1
    cur.execute(
        """INSERT INTO score_signature
           (athlete_id,round_no,recorder_name,referee_name,athlete_name,
            recorder_signature,referee_signature,athlete_signature,bypass_signed,
            started_at,finished_at,stopped_by_red)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?)""",
        (
            athlete_id, round_no, "DEMO RECORDER", "DEMO REFEREE", athlete_name,
            "demo-recorder" if signed else None,
            "demo-referee" if signed else None,
            "demo-athlete" if signed else None,
            0 if signed else 1,
            NOW, FINISHED, 0,
        ),
    )


def delete_demo(cur):
    row = cur.execute("SELECT id FROM event WHERE name=?", (EVENT_NAME,)).fetchone()
    if not row:
        return
    event_id = row[0]
    athlete_ids = [r[0] for r in cur.execute("SELECT id FROM athlete WHERE event_id=?", (event_id,)).fetchall()]
    if athlete_ids:
        marks = ",".join("?" for _ in athlete_ids)
        for table in ("score_entry", "score_signature", "tie_break_entry", "score_edit_log"):
            cur.execute(f"DELETE FROM {table} WHERE athlete_id IN ({marks})", athlete_ids)
    cur.execute("DELETE FROM bracket_match WHERE event_id=?", (event_id,))
    cur.execute("DELETE FROM results_approved_setting WHERE event_id=?", (event_id,))
    cur.execute("DELETE FROM athlete WHERE event_id=?", (event_id,))
    cur.execute("DELETE FROM event WHERE id=?", (event_id,))


def main():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("PRAGMA foreign_keys=OFF")
    delete_demo(cur)

    creator = cur.execute("SELECT id FROM user WHERE role='superadmin' ORDER BY id LIMIT 1").fetchone()
    event_id = next_id(cur, "event")
    cur.execute(
        """INSERT INTO event
           (id,name,event_group,category,competition_date,location,lane_count,direct_qualifiers,
            has_round_two,round_two_cutoff_rank,next_round_label,round_two_advancers,created_at,created_by,bracket_qualifiers)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (
            event_id, EVENT_NAME, "DEMO", "men", "2026-09-15", "Khon Kaen University, Thailand",
            8, 4, 1, 16, "รอบ 8 คน", 4, NOW, creator[0] if creator else None, 8,
        ),
    )

    athlete_ids: dict[str, int] = {}
    aid = next_id(cur, "athlete")
    for idx, country in enumerate(COUNTRIES, 1):
        athlete_id = aid + idx - 1
        athlete_ids[country] = athlete_id
        cur.execute(
            """INSERT INTO athlete
               (id,event_id,bib_no,name,affiliation,start_order,lane_no,lane_order,status,red_card_count,
                created_at,round_two_disabled,round_two_disabled_at,round_two_disabled_by)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
            (
                athlete_id, event_id, str(idx), country, country, idx,
                ((idx - 1) % 8) + 1, ((idx - 1) // 8) + 1,
                "finished", 0, NOW, 0, None, None,
            ),
        )
        # 2/3 use full signatures; 1/3 demonstrate Superadmin approval.
        add_round(cur, athlete_id, country, 1, R1_TOTALS[idx - 1], signed=(idx % 3 != 0))

    # Round 2 = everybody with Class 5-16, exactly 12 people in this deterministic demo.
    for country, total in R2_BY_COUNTRY.items():
        add_round(cur, athlete_ids[country], country, 2, total, signed=True)

    # Seed order produced by the app logic for this dataset.
    seeds = {
        1: "BELGIUM", 2: "SPAIN", 3: "ITALY", 4: "MADAGASCAR",
        5: "SWITZERLAND", 6: "MALAYSIA", 7: "JAPAN", 8: "GERMANY",
    }

    # Quarter Final pairings: 1v8, 4v5, 3v6, 2v7.
    qf_pairs = [(1, 8), (4, 5), (3, 6), (2, 7)]
    qf_winners = []
    for match_no, (sa, sb) in enumerate(qf_pairs, 1):
        ca, cb = seeds[sa], seeds[sb]
        add_round(cur, athlete_ids[ca], ca, 3, QF_SCORES[ca], signed=True)
        add_round(cur, athlete_ids[cb], cb, 3, QF_SCORES[cb], signed=True)
        winner = ca if QF_SCORES[ca] > QF_SCORES[cb] else cb
        qf_winners.append(winner)
        cur.execute(
            "INSERT INTO bracket_match(event_id,round_name,match_no,athlete_a_id,athlete_b_id,winner_id) VALUES (?,?,?,?,?,?)",
            (event_id, "QF", match_no, athlete_ids[ca], athlete_ids[cb], athlete_ids[winner]),
        )

    # Semi Final: winner QF1 v QF2, winner QF3 v QF4.
    sf_pairs = [(qf_winners[0], qf_winners[1]), (qf_winners[2], qf_winners[3])]
    sf_winners = []
    for match_no, (ca, cb) in enumerate(sf_pairs, 1):
        add_round(cur, athlete_ids[ca], ca, 4, SF_SCORES[ca], signed=True)
        add_round(cur, athlete_ids[cb], cb, 4, SF_SCORES[cb], signed=True)
        winner = ca if SF_SCORES[ca] > SF_SCORES[cb] else cb
        sf_winners.append(winner)
        cur.execute(
            "INSERT INTO bracket_match(event_id,round_name,match_no,athlete_a_id,athlete_b_id,winner_id) VALUES (?,?,?,?,?,?)",
            (event_id, "SF", match_no, athlete_ids[ca], athlete_ids[cb], athlete_ids[winner]),
        )

    # Final.
    ca, cb = sf_winners
    add_round(cur, athlete_ids[ca], ca, 5, FINAL_SCORES[ca], signed=True)
    add_round(cur, athlete_ids[cb], cb, 5, FINAL_SCORES[cb], signed=True)
    champion = ca if FINAL_SCORES[ca] > FINAL_SCORES[cb] else cb
    cur.execute(
        "INSERT INTO bracket_match(event_id,round_name,match_no,athlete_a_id,athlete_b_id,winner_id) VALUES (?,?,?,?,?,?)",
        (event_id, "F", 1, athlete_ids[ca], athlete_ids[cb], athlete_ids[champion]),
    )

    cur.execute(
        """INSERT INTO results_approved_setting
           (event_id,competition_title,host_line,date_line,location_line,country_label,
            president_title,president_name,technical_title,technical_name,umpires_text,
            approved_text,show_official_pages,updated_at)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (
            event_id,
            "52nd PÉTANQUE WORLD CHAMPIONSHIP 2026 • DEMO EVENT",
            "Khon Kaen University, Thailand",
            "15 September 2026",
            "Khon Kaen University, Thailand",
            "COUNTRY",
            "DEMO PRESIDENT", "DEMO OFFICIAL",
            "TECHNICAL DELEGATE", "DEMO TECHNICAL",
            "DEMO UMPIRE TEAM",
            "OFFICIAL DEMO RESULTS • APPROVED",
            1, FINISHED,
        ),
    )

    con.commit()

    # Independent verification using SQL only.
    athlete_count = cur.execute("SELECT COUNT(*) FROM athlete WHERE event_id=?", (event_id,)).fetchone()[0]
    r1_count = cur.execute(
        "SELECT COUNT(DISTINCT athlete_id) FROM score_signature s JOIN athlete a ON a.id=s.athlete_id WHERE a.event_id=? AND s.round_no=1 AND s.finished_at IS NOT NULL",
        (event_id,),
    ).fetchone()[0]
    r2_count = cur.execute(
        "SELECT COUNT(DISTINCT athlete_id) FROM score_signature s JOIN athlete a ON a.id=s.athlete_id WHERE a.event_id=? AND s.round_no=2 AND s.finished_at IS NOT NULL",
        (event_id,),
    ).fetchone()[0]
    bracket_count = cur.execute("SELECT COUNT(*) FROM bracket_match WHERE event_id=?", (event_id,)).fetchone()[0]
    print(f"DEMO_EVENT_ID={event_id}")
    print(f"COUNTRIES={athlete_count}")
    print(f"ROUND1_FINISHED={r1_count}")
    print(f"ROUND2_FINISHED={r2_count}")
    print(f"BRACKET_MATCHES={bracket_count}")
    print(f"CHAMPION={champion}")
    con.close()


if __name__ == "__main__":
    main()
