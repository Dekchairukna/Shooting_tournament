"""ล้างการคีย์. Run: python -m unittest tests.test_reset_scores -v"""
import os, re, tempfile, unittest
from datetime import date, datetime

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-reset-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from flask import g
from app import (app, db, User, Event, Athlete, ScoreEntry, ScoreSignature, ScoreEditLog, TieBreakEntry,
                 BracketMatch, athlete_round_status, ensure_round_entries)


class ResetTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()
        self.boss = User(username="boss", role="superadmin", password_hash="!UNSET!")
        self.admin = User(username="adm", role="admin", password_hash="!UNSET!")
        db.session.add_all([self.boss, self.admin]); db.session.commit()
        self.e = Event(name="E", event_group="ทั่วไป", category="ชาย", competition_date=date(2026, 10, 9), location="",
                       lane_count=2, direct_qualifiers=4, has_round_two=True, round_two_cutoff_rank=16,
                       next_round_label="รอบ 8 คน", round_two_advancers=4)
        db.session.add(self.e); db.session.flush()
        self.a = []
        for i in range(1, 4):
            a = Athlete(event_id=self.e.id, bib_no=str(i), name=f"A{i}", affiliation="T", start_order=i,
                        lane_no=(i - 1) % 2 + 1, lane_order=(i - 1) // 2 + 1, status="finished", red_card_count=1)
            db.session.add(a); db.session.flush(); self.a.append(a)
            ensure_round_entries(a.id, 1)
            for en in ScoreEntry.query.filter_by(athlete_id=a.id, round_no=1).limit(5):
                en.score, en.is_scored = 5, True
            db.session.add(ScoreSignature(athlete_id=a.id, round_no=1, started_at=datetime.utcnow(),
                                          finished_at=datetime.utcnow(), bypass_signed=(i == 3)))
            db.session.add(TieBreakEntry(athlete_id=a.id, round_no=1, station_no=1, score=3))
        db.session.add(BracketMatch(event_id=self.e.id, round_name="QF", match_no=1, athlete_a_id=self.a[0].id, winner_id=self.a[0].id))
        db.session.commit()

    def tearDown(self):
        db.session.remove(); self.ctx.pop()

    def client(self, user):
        g.pop("_login_user", None)
        c = app.test_client()
        with c.session_transaction() as s:
            s["_user_id"] = str(user.id); s["_fresh"] = True
        page = c.get("/login").get_data(as_text=True)
        c.environ_base["HTTP_X_CSRF_TOKEN"] = re.search(r'name="csrf-token" content="([^"]+)"', page).group(1)
        return c

    def scored(self, a):
        return ScoreEntry.query.filter(ScoreEntry.athlete_id == a.id, ScoreEntry.is_scored.is_(True)).count()

    def test_reset_one_athlete(self):
        c = self.client(self.admin)
        r = c.post(f"/athletes/{self.a[0].id}/reset-round", data={"round": "1", "reason": "คีย์ก่อนเวลา"})
        self.assertEqual(r.status_code, 302)
        a0 = db.session.get(Athlete, self.a[0].id)
        self.assertEqual(self.scored(a0), 0)
        self.assertEqual(ScoreEntry.query.filter_by(athlete_id=a0.id, round_no=1).count(), 20)  # ช่องยังอยู่ ว่างเปล่า
        self.assertEqual(ScoreSignature.query.filter_by(athlete_id=a0.id).count(), 0)
        self.assertEqual(TieBreakEntry.query.filter_by(athlete_id=a0.id).count(), 0)
        self.assertEqual(athlete_round_status(a0, 1), "waiting")
        self.assertEqual((a0.status, a0.red_card_count), ("waiting", 0))
        logs = ScoreEditLog.query.filter_by(athlete_id=a0.id).all()
        self.assertEqual(len(logs), 5)
        self.assertEqual(logs[0].reason, "ล้างการคีย์: คีย์ก่อนเวลา")
        self.assertEqual(logs[0].editor_username, "adm")
        self.assertEqual(self.scored(self.a[1]), 5)  # คนอื่นไม่โดน
        self.assertEqual(BracketMatch.query.count(), 1)  # รายคนไม่ล้างสาย

    def test_reset_needs_reason(self):
        c = self.client(self.admin)
        c.post(f"/athletes/{self.a[0].id}/reset-round", data={"round": "1", "reason": " "})
        self.assertEqual(self.scored(self.a[0]), 5)

    def test_event_reset_skips_approved_and_clears_bracket(self):
        c = self.client(self.boss)
        c.post(f"/events/{self.e.id}/reset-scores", data={"round": "1", "reason": "ทดลองระบบ", "confirm_text": "ล้าง", "skip_approved": "yes"})
        self.assertEqual([self.scored(a) for a in self.a], [0, 0, 5])
        self.assertEqual(BracketMatch.query.count(), 0)

    def test_event_reset_requires_confirm_word_and_superadmin(self):
        self.client(self.boss).post(f"/events/{self.e.id}/reset-scores", data={"round": "1", "reason": "x", "confirm_text": "ok"})
        self.assertEqual(self.scored(self.a[0]), 5)
        self.client(self.admin).post(f"/events/{self.e.id}/reset-scores", data={"round": "1", "reason": "x", "confirm_text": "ล้าง"})
        self.assertEqual(self.scored(self.a[0]), 5)

    def test_buttons_visible(self):
        html = self.client(self.boss).get(f"/athletes/{self.a[0].id}/scorecard?round=1").get_data(as_text=True)
        self.assertIn("ล้างการคีย์รอบ 1", html)
        html = self.client(self.boss).get(f"/events/{self.e.id}/athletes").get_data(as_text=True)
        self.assertIn("ล้างการคีย์ทั้งอีเวนต์", html)
        html = self.client(self.admin).get(f"/events/{self.e.id}/athletes").get_data(as_text=True)
        self.assertNotIn("ล้างการคีย์ทั้งอีเวนต์", html)


if __name__ == "__main__":
    unittest.main()
