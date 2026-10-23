"""ถอนตัว. Run: python -m unittest tests.test_withdraw -v"""
import os, re, tempfile, unittest
from datetime import date, datetime

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-withdraw-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from flask import g
from app import app, db, User, Event, Athlete, ScoreEntry, ScoreSignature, ensure_round_entries, invalidate_poll_cache


class WithdrawTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()
        self.admin = User(username="adm", role="admin", password_hash="!UNSET!"); db.session.add(self.admin)
        self.e = Event(name="E", event_group="ม.ปลาย", category="ชาย", competition_date=date(2026, 10, 24), location="",
                       lane_count=2, direct_qualifiers=1, has_round_two=True, round_two_cutoff_rank=3, round_two_mode="cutoff",
                       next_round_label="รอบ 8 คน", round_two_advancers=1)
        db.session.add(self.e); db.session.flush()
        self.a = []
        for i, fives in enumerate([4, 3, 2, 0], start=1):
            a = Athlete(event_id=self.e.id, bib_no=str(i), name=f"A{i}", affiliation="T", start_order=i,
                        lane_no=(i - 1) % 2 + 1, lane_order=(i - 1) // 2 + 1)
            db.session.add(a); db.session.flush(); self.a.append(a)
            if fives:
                ensure_round_entries(a.id, 1)
                for en in ScoreEntry.query.filter_by(athlete_id=a.id, round_no=1).limit(fives):
                    en.score, en.is_scored = 5, True
                db.session.add(ScoreSignature(athlete_id=a.id, round_no=1, started_at=datetime.utcnow(), finished_at=datetime.utcnow(), bypass_signed=True))
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

    def ov(self, c, rnd=1):
        g.pop("shooting_request_cache", None); invalidate_poll_cache(self.e.id)
        return {r["name"]: r for r in c.get(f"/events/{self.e.id}/overview-data?round={rnd}").get_json()}

    def test_withdraw_and_cancel(self):
        c = self.client()
        self.assertEqual(self.ov(c)["A4"]["status"], "waiting")
        c.post(f"/athletes/{self.a[3].id}/withdraw", data={"action": "withdraw"})
        rows = self.ov(c)
        self.assertTrue(rows["A4"]["withdrawn"]); self.assertEqual(rows["A4"]["status"], "finished")  # รอบ 1 จบได้
        # A2 ถอนตัว: คะแนน 15 ไม่คิด และไม่มีสิทธิ์รอบ 2 แม้ cutoff ครอบถึง
        c.post(f"/athletes/{self.a[1].id}/withdraw", data={"action": "withdraw"})
        rows = self.ov(c)
        self.assertEqual(rows["A2"]["total"], 0)
        r2 = {k for k, v in self.ov(c, 2).items() if not v["is_round2_direct_placeholder"]}
        self.assertNotIn("A2", r2); self.assertNotIn("A4", r2); self.assertIn("A3", r2)
        html = c.get(f"/events/{self.e.id}/athletes").get_data(as_text=True)
        self.assertIn("ยกเลิกถอนตัว", html)
        c.post(f"/athletes/{self.a[1].id}/withdraw", data={"action": "cancel"})
        rows = self.ov(c)
        self.assertFalse(rows["A2"]["withdrawn"]); self.assertEqual(rows["A2"]["total"], 15)

    def test_scorecard_shows_button(self):
        c = self.client()
        html = c.get(f"/athletes/{self.a[0].id}/scorecard?round=1").get_data(as_text=True)
        self.assertIn("toggle_athlete_withdrawn" if False else "/withdraw", html)


if __name__ == "__main__":
    unittest.main()
