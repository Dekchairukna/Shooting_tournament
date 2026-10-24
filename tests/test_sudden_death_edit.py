"""แก้ไข Sudden death. Run: python -m unittest tests.test_sudden_death_edit -v"""
import os, re, tempfile, unittest
from datetime import date, datetime

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-sd-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from flask import g
from app import app, db, User, Event, Athlete, ScoreEntry, ScoreSignature, TieBreakEntry, ensure_round_entries, invalidate_poll_cache


class SuddenDeathEditTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()
        self.admin = User(username="adm", role="superadmin", password_hash="!UNSET!"); db.session.add(self.admin)
        self.e = Event(name="E", event_group="x", category="ชาย", competition_date=date(2026, 10, 24), location="", lane_count=2,
                       direct_qualifiers=1, has_round_two=False, next_round_label="รอบ 8 คน", round_two_advancers=4)
        db.session.add(self.e); db.session.flush()
        self.a = []
        for i in range(1, 4):
            a = Athlete(event_id=self.e.id, bib_no=str(i), name=f"A{i}", affiliation="T", start_order=i, lane_no=1, lane_order=i)
            db.session.add(a); db.session.flush(); self.a.append(a)
            ensure_round_entries(a.id, 1)
            for en in ScoreEntry.query.filter_by(athlete_id=a.id, round_no=1).limit(3 if i < 3 else 1):
                en.score, en.is_scored = 5, True
            db.session.add(ScoreSignature(athlete_id=a.id, round_no=1, started_at=datetime.utcnow(), finished_at=datetime.utcnow(), bypass_signed=True))
        for a, sc in ((self.a[0], 5), (self.a[1], 0)):  # คีย์ผิด: A1 ชนะ
            for st in range(1, 6):
                db.session.add(TieBreakEntry(athlete_id=a.id, round_no=1, station_no=st, score=sc if st == 1 else 0))
        db.session.commit()

    def tearDown(self):
        db.session.remove(); self.ctx.pop()

    def test_admin_cannot_edit(self):
        u = User(username="plain", role="admin", password_hash="!UNSET!"); db.session.add(u); db.session.commit()
        g.pop("_login_user", None)
        c = app.test_client()
        with c.session_transaction() as s:
            s["_user_id"] = str(u.id); s["_fresh"] = True
        page = c.get("/login").get_data(as_text=True)
        c.environ_base["HTTP_X_CSRF_TOKEN"] = re.search(r'name="csrf-token" content="([^"]+)"', page).group(1)
        html = c.get(f"/events/{self.e.id}/overview?round=1").get_data(as_text=True)
        self.assertNotIn("tiebreak?round=1&amp;ids=", html)
        self.assertIn("const CAN_EDIT_SD = false", html)
        ids = f"{self.a[0].id},{self.a[1].id}"
        c.post(f"/events/{self.e.id}/tiebreak?round=1&ids={ids}", data={"action": "clear", "ids": ids, "round": "1"})
        self.assertEqual(TieBreakEntry.query.count(), 10)

    def test_edit_and_clear(self):
        g.pop("_login_user", None)
        c = app.test_client()
        with c.session_transaction() as s:
            s["_user_id"] = str(self.admin.id); s["_fresh"] = True
        page = c.get("/login").get_data(as_text=True)
        c.environ_base["HTTP_X_CSRF_TOKEN"] = re.search(r'name="csrf-token" content="([^"]+)"', page).group(1)
        ov = c.get(f"/events/{self.e.id}/overview?round=1").get_data(as_text=True)
        self.assertIn("tiebreak?round=1&amp;ids=", ov)
        ids = f"{self.a[0].id},{self.a[1].id}"
        html = c.get(f"/events/{self.e.id}/tiebreak?round=1&ids={ids}").get_data(as_text=True)
        self.assertIn("แก้ไขผล Sudden death", html)
        data = {"action": "replace", "ids": ids, "round": "1"}
        for a, first in ((self.a[0], 0), (self.a[1], 3)):
            for st in range(1, 6):
                data[f"tb_{a.id}_{st}"] = str(first if st == 1 else 0)
        c.post(f"/events/{self.e.id}/tiebreak?round=1&ids={ids}", data=data)
        g.pop("shooting_request_cache", None); invalidate_poll_cache(self.e.id)
        rows = {r["name"]: r for r in c.get(f"/events/{self.e.id}/overview-data?round=1").get_json()}
        self.assertEqual(rows["A2"]["rank"], 1)
        self.assertEqual(TieBreakEntry.query.count(), 10)
        c.post(f"/events/{self.e.id}/tiebreak?round=1&ids={ids}", data={"action": "clear", "ids": ids, "round": "1"})
        self.assertEqual(TieBreakEntry.query.count(), 0)


if __name__ == "__main__":
    unittest.main()
