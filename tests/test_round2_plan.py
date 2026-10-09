"""ผังลงสนามรอบ 2. Run: python -m unittest tests.test_round2_plan -v"""
import os, re, tempfile, unittest
from datetime import date, datetime

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-r2plan-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from flask import g
from app import (app, db, User, Event, Athlete, ScoreEntry, ScoreSignature, ensure_round_entries,
                 invalidate_poll_cache, round2_lane_slot)


class Round2PlanTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()
        self.admin = User(username="adm", role="admin", password_hash="!UNSET!")
        db.session.add(self.admin); db.session.commit()
        self.events = []
        for name in ["12 ชาย", "12 หญิง"]:
            e = Event(name=name, event_group="12", category="ชาย", competition_date=date(2026, 10, 9), location="",
                      lane_count=8, direct_qualifiers=4, has_round_two=True, round_two_cutoff_rank=15,
                      round_two_mode="next", next_round_label="รอบ 8 คน", round_two_advancers=4)
            db.session.add(e); db.session.flush(); self.events.append(e)
            for i in range(1, 26):
                a = Athlete(event_id=e.id, bib_no=str(i), name=f"{name}-T{i}", affiliation=f"{name}-T{i}", start_order=i,
                            lane_no=(i - 1) % 8 + 1, lane_order=(i - 1) // 8 + 1, status="finished")
                db.session.add(a); db.session.flush()
                ensure_round_entries(a.id, 1)
                # คะแนนไม่ซ้ำกัน: แต้ม 1..25 (ใช้ช่อง 1 แต้ม)
                ens = ScoreEntry.query.filter_by(athlete_id=a.id, round_no=1).order_by(ScoreEntry.id).all()
                for k in range(i // 5):
                    ens[k].score, ens[k].is_scored = 5, True
                for k in range(i // 5, i // 5 + i % 5):
                    ens[k].score, ens[k].is_scored = 1, True
                db.session.add(ScoreSignature(athlete_id=a.id, round_no=1, started_at=datetime.utcnow(),
                                              finished_at=datetime.utcnow(), bypass_signed=True))
        db.session.commit()

    def tearDown(self):
        db.session.remove(); self.ctx.pop()

    def client(self):
        g.pop("_login_user", None)
        c = app.test_client()
        with c.session_transaction() as s:
            s["_user_id"] = str(self.admin.id); s["_fresh"] = True
        page = c.get("/login").get_data(as_text=True)
        c.environ_base["HTTP_X_CSRF_TOKEN"] = re.search(r'name="csrf-token" content="([^"]+)"', page).group(1)
        return c

    def test_slot_back_fill_like_sheet(self):
        e = self.events[0]
        self.assertEqual(round2_lane_slot(e, 1, 15, "back"), (2, 1))   # คิว 1 สนาม 1 ว่าง
        self.assertEqual(round2_lane_slot(e, 7, 15, "back"), (8, 1))
        self.assertEqual(round2_lane_slot(e, 8, 15, "back"), (1, 2))
        self.assertEqual(round2_lane_slot(e, 15, 15, "back"), (8, 2))
        self.assertEqual(round2_lane_slot(e, 1, 15, "front"), (1, 1))

    def test_plan_page_and_apply(self):
        c = self.client()
        ids = ",".join(str(e.id) for e in self.events)
        html = c.get(f"/events/round2-plan?ids={ids}&fill=back").get_data(as_text=True)
        self.assertIn("12 ชาย", html); self.assertIn("12 หญิง", html)
        self.assertIn("(15 ทีม)", html)
        self.assertEqual(html.count("ลำดับที่ 2"), 2)
        r = c.post("/events/round2-plan", data={"ids": ids, "fill": "back"})
        self.assertEqual(r.status_code, 302)
        g.pop("shooting_request_cache", None)
        e = self.events[0]
        invalidate_poll_cache(e.id)
        rows = [r for r in c.get(f"/events/{e.id}/overview-data?round=2").get_json() if not r["is_round2_direct_placeholder"]]
        first = min(rows, key=lambda r: r["display_order"])
        self.assertEqual((first["display_lane_no"], first["display_lane_order"]), (2, 1))
        lane1 = [r for r in rows if r["display_lane_no"] == 1]
        self.assertEqual([r["display_lane_order"] for r in lane1], [2])
        # หน้าคิวสนาม 1 รอบ 2 เห็นแค่คนเดียว
        self.assertEqual(db.session.get(Event, e.id).round_two_fill, "back")


if __name__ == "__main__":
    unittest.main()
