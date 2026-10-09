"""รอบ 2 เอาแค่ N คน + Sudden death, ปุ่มปิดเส้นตัด, DQ ยังตีรอบถัดไปได้.
Run: python -m unittest tests.test_round2_exact -v"""
import os, re, tempfile, unittest
from datetime import date, datetime

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-r2exact-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from flask import g
from app import (app, db, User, Event, Athlete, ScoreEntry, ScoreSignature, TieBreakEntry, AthletePenalty,
                 ensure_round_entries, invalidate_poll_cache)


class Round2ExactTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()
        self.admin = User(username="adm", role="admin", password_hash="!UNSET!")
        db.session.add(self.admin); db.session.commit()
        self.e = Event(name="E", event_group="ทั่วไป", category="ชาย", competition_date=date(2026, 10, 9), location="",
                       lane_count=2, direct_qualifiers=1, has_round_two=True, round_two_cutoff_rank=2,
                       round_two_mode="next", next_round_label="รอบ 4 คน", round_two_advancers=1)
        db.session.add(self.e); db.session.flush()
        self.a = []
        # A1=25 ผ่านตรง · A2=20 · A3=15 · A4=15 (เท่ากันที่ลำดับสุดท้าย) · A5=10
        for i, fives in enumerate([5, 4, 3, 3, 2], start=1):
            a = Athlete(event_id=self.e.id, bib_no=str(i), name=f"A{i}", affiliation="T", start_order=i,
                        lane_no=(i - 1) % 2 + 1, lane_order=(i - 1) // 2 + 1, status="finished")
            db.session.add(a); db.session.flush(); self.a.append(a)
            ensure_round_entries(a.id, 1)
            for en in ScoreEntry.query.filter_by(athlete_id=a.id, round_no=1).limit(fives):
                en.score, en.is_scored = 5, True
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

    def overview(self, c, rnd=1):
        g.pop("shooting_request_cache", None)
        invalidate_poll_cache(self.e.id)
        return {r["name"]: r for r in c.get(f"/events/{self.e.id}/overview-data?round={rnd}").get_json()}

    def test_tie_at_last_place_needs_sudden_death(self):
        c = self.client()
        rows = self.overview(c)
        self.assertTrue(rows["A3"]["shoot_off_required"])
        self.assertTrue(rows["A4"]["shoot_off_required"])
        self.assertFalse(rows["A2"]["shoot_off_required"])
        r2 = self.overview(c, 2)
        self.assertIn("A2", r2)
        self.assertNotIn("A3", r2); self.assertNotIn("A4", r2)  # ยังไม่ตัดสิน

    def test_after_sudden_death_exactly_n(self):
        db.session.add(TieBreakEntry(athlete_id=self.a[2].id, round_no=1, station_no=1, score=5))
        db.session.add(TieBreakEntry(athlete_id=self.a[3].id, round_no=1, station_no=1, score=3))
        db.session.commit()
        c = self.client()
        rows = self.overview(c)
        self.assertFalse(rows["A3"]["shoot_off_required"])
        r2 = {k for k, v in self.overview(c, 2).items() if not v["is_round2_direct_placeholder"]}
        self.assertEqual(r2, {"A2", "A3"})

    def test_cutoff_mode_unchanged_takes_all_tied(self):
        self.e.round_two_mode = "cutoff"; self.e.round_two_cutoff_rank = 3; db.session.commit()
        c = self.client()
        r2 = {k for k, v in self.overview(c, 2).items() if not v["is_round2_direct_placeholder"]}
        self.assertEqual(r2, {"A2", "A3", "A4"})

    def test_cut_line_toggle(self):
        c = self.client()
        self.assertTrue(any(r["cut_line_after"] for r in self.overview(c).values()))
        c.post(f"/events/{self.e.id}/overview-cut-lines", data={"round": "1"})
        self.assertFalse(any(r["cut_line_after"] for r in self.overview(c).values()))
        self.assertIn("เส้นตัด: ปิด", c.get(f"/events/{self.e.id}/overview?round=1").get_data(as_text=True))

    def test_dq_can_still_play_next_round(self):
        self.e.round_two_cutoff_rank = 4; db.session.commit()
        db.session.add(AthletePenalty(athlete_id=self.a[1].id, round_no=1, level=2)); db.session.commit()
        c = self.client()
        rows = self.overview(c)
        self.assertTrue(rows["A2"]["disqualified"]); self.assertEqual(rows["A2"]["total"], 0)
        r2 = {k for k, v in self.overview(c, 2).items() if not v["is_round2_direct_placeholder"]}
        self.assertIn("A2", r2)


if __name__ == "__main__":
    unittest.main()
