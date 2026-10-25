"""คู่ Knockout เสมอ → ตี 7 เมตรทุกสถานี → Sudden death. Run: python -m unittest tests.test_bracket_tie -v"""
import os, re, tempfile, unittest
from datetime import date, datetime

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-btie-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from flask import g
from app import app, db, User, Event, Athlete, ScoreEntry, ScoreSignature, BracketMatch, TieBreakEntry, ensure_round_entries


class BracketTieTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()
        self.admin = User(username="adm", role="admin", password_hash="!UNSET!"); db.session.add(self.admin)
        self.e = Event(name="E", event_group="x", category="หญิง", competition_date=date(2026, 10, 25), location="", lane_count=4,
                       direct_qualifiers=8, has_round_two=False, next_round_label="รอบ 8 คน", round_two_advancers=4)
        db.session.add(self.e); db.session.flush()
        for i in range(1, 9):
            a = Athlete(event_id=self.e.id, bib_no=str(i), name=f"A{i}", affiliation="T", start_order=i, lane_no=(i - 1) % 4 + 1, lane_order=1)
            db.session.add(a); db.session.flush()
            ensure_round_entries(a.id, 1)
            ens = ScoreEntry.query.filter_by(athlete_id=a.id, round_no=1).order_by(ScoreEntry.id).all()
            for k in range(i):  # คะแนนไม่ซ้ำกัน
                ens[k].score, ens[k].is_scored = 1, True
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

    def finish_qf(self, athlete_id, score):
        ensure_round_entries(athlete_id, 3)
        for k, en in enumerate(ScoreEntry.query.filter_by(athlete_id=athlete_id, round_no=3).order_by(ScoreEntry.id)):
            en.score, en.is_scored = (5 if k < score else 0), True
        db.session.add(ScoreSignature(athlete_id=athlete_id, round_no=3, started_at=datetime.utcnow(), finished_at=datetime.utcnow(), bypass_signed=True))
        db.session.commit()

    def test_tie_flow(self):
        c = self.client()
        c.get(f"/events/{self.e.id}/bracket")
        m = BracketMatch.query.filter_by(event_id=self.e.id, round_name="QF", match_no=1).one()
        a, b = m.athlete_a_id, m.athlete_b_id
        self.finish_qf(a, 3); self.finish_qf(b, 3)
        g.pop("shooting_request_cache", None)
        html = c.get(f"/events/{self.e.id}/bracket").get_data(as_text=True)
        self.assertIn("เสมอ · ตี 7 เมตรทุกสถานี", html)
        self.assertIn("ตี 7 เมตร →", html)
        url = f"/events/{self.e.id}/tiebreak?round=3&ids={a},{b}"
        data = {"action": "add", "ids": f"{a},{b}", "round": "3", "next": "bracket"}
        for x in (a, b):
            for st in range(1, 6):
                data[f"tb_{x}_{st}"] = "3" if st == 1 else "0"
        g.pop("shooting_request_cache", None)
        r = c.post(url, data=data)
        self.assertIn("/bracket", r.location)
        g.pop("shooting_request_cache", None)
        html = c.get(f"/events/{self.e.id}/bracket").get_data(as_text=True)
        self.assertIn("ยังเสมอ · Sudden death", html)
        self.assertIsNone(db.session.get(BracketMatch, m.id).winner_id)
        data = {"action": "add", "ids": f"{a},{b}", "round": "3", "next": "bracket"}
        for x, v in ((a, "5"), (b, "0")):
            for st in range(1, 6):
                data[f"tb_{x}_{st}"] = v if st == 1 else "0"
        g.pop("shooting_request_cache", None)
        c.post(url, data=data)
        db.session.expire_all()
        self.assertEqual(db.session.get(BracketMatch, m.id).winner_id, a)


if __name__ == "__main__":
    unittest.main()
