"""บทลงโทษไม่มาแข่งขันตามเวลา. Run: python -m unittest tests.test_penalty -v"""
import os, re, tempfile, unittest
from datetime import date, datetime

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-penalty-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from flask import g
from app import (app, db, User, Event, Athlete, ScoreEntry, ScoreSignature, AthletePenalty, ScoreEditLog,
                 ensure_round_entries, invalidate_poll_cache)


class PenaltyTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()
        self.admin = User(username="adm", role="admin", password_hash="!UNSET!")
        db.session.add(self.admin); db.session.commit()
        self.e = Event(name="E", event_group="ทั่วไป", category="ชาย", competition_date=date(2026, 10, 9), location="",
                       lane_count=2, direct_qualifiers=1, has_round_two=True, round_two_cutoff_rank=3,
                       next_round_label="รอบ 8 คน", round_two_advancers=1)
        db.session.add(self.e); db.session.flush()
        self.a = []
        # A1=25, A2=20, A3=15, A4 ไม่มา
        for i, fives in enumerate([5, 4, 3, 0], start=1):
            a = Athlete(event_id=self.e.id, bib_no=str(i), name=f"A{i}", affiliation="T", start_order=i,
                        lane_no=(i - 1) % 2 + 1, lane_order=(i - 1) // 2 + 1, status="waiting")
            db.session.add(a); db.session.flush(); self.a.append(a)
            ensure_round_entries(a.id, 1)
            if fives:
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

    def pen(self, c, athlete, level, rnd=1):
        return c.post(f"/athletes/{athlete.id}/penalty", data={"round": str(rnd), "level": str(level), "note": "สาย"})

    def test_minus_five_then_dq(self):
        c = self.client()
        self.pen(c, self.a[0], 1)
        rows = self.overview(c)
        self.assertEqual(rows["A1"]["total"], 20)
        self.assertEqual(rows["A1"]["penalty"], 5)
        self.assertEqual(rows["A1"]["rank"], 1)  # 20 เท่า A2 แต่ 5 มากกว่า
        self.pen(c, self.a[0], 2)  # เกิน 10 นาที = DQ ไม่คำนวณคะแนน
        rows = self.overview(c)
        self.assertEqual(rows["A1"]["total"], 0)
        self.assertTrue(rows["A1"]["disqualified"])
        self.assertEqual(rows["A2"]["rank"], 1)
        self.assertEqual(AthletePenalty.query.count(), 1)
        self.assertEqual(ScoreEditLog.query.filter(ScoreEditLog.reason.like("บทลงโทษ%")).count(), 2)

    def test_disqualified_ranks_last_and_round_completes(self):
        c = self.client()
        rows = self.overview(c)
        self.assertEqual(rows["A4"]["status"], "waiting")
        self.pen(c, self.a[3], 2)
        rows = self.overview(c)
        self.assertTrue(rows["A4"]["disqualified"])
        self.assertEqual(rows["A4"]["status"], "finished")  # ไม่ค้างคิว รอบ 1 จบได้
        self.assertEqual(rows["A4"]["rank"], 4)
        # DQ ไม่ได้สิทธิ์รอบ 2 แม้ cutoff 3 จะครอบถึง
        r2 = self.overview(c, 2)
        self.assertNotIn("A4", r2)
        self.assertIn("A2", r2); self.assertIn("A3", r2)

    def test_dq_not_direct_even_with_high_score(self):
        c = self.client()
        self.pen(c, self.a[0], 2)
        rows = self.overview(c)
        self.assertEqual(rows["A1"]["rank"], 4)
        self.assertEqual(rows["A2"]["rank"], 1)
        r2 = self.overview(c, 2)
        self.assertNotIn("A1", r2)

    def test_cancel_penalty_restores(self):
        c = self.client()
        self.pen(c, self.a[0], 2)
        self.pen(c, self.a[0], 0)
        rows = self.overview(c)
        self.assertEqual(rows["A1"]["total"], 25)
        self.assertFalse(rows["A1"]["disqualified"])
        self.assertEqual(AthletePenalty.query.count(), 0)

    def test_scorecard_and_print_show_penalty(self):
        c = self.client()
        self.pen(c, self.a[1], 1)
        html = c.get(f"/athletes/{self.a[1].id}/scorecard?round=1").get_data(as_text=True)
        self.assertIn("คะแนนสุทธิ 15", html)
        html = c.get(f"/athletes/{self.a[1].id}/scorecard-print?round=1").get_data(as_text=True)
        self.assertIn("หัก −5", html)

    def test_old_level3_still_dq(self):
        db.session.add(AthletePenalty(athlete_id=self.a[3].id, round_no=1, level=3)); db.session.commit()
        rows = self.overview(self.client())
        self.assertTrue(rows["A4"]["disqualified"])

    def test_reset_clears_penalty(self):
        c = self.client()
        self.pen(c, self.a[0], 2)
        c.post(f"/athletes/{self.a[0].id}/reset-round", data={"round": "1", "reason": "คีย์ผิด"})
        self.assertEqual(AthletePenalty.query.count(), 0)


if __name__ == "__main__":
    unittest.main()
