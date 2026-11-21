"""สถานะ: ตีแล้วหลังกด "จบการตี" · ชื่อที่พิมพ์นับเป็นการลงชื่อ · APPROVED เมื่อลงชื่อครบ 3 ฝ่าย.
Run: python -m unittest tests.test_finish_status -v"""
import os, re, tempfile, unittest
from datetime import date, datetime

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-finish-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from flask import g
from app import app, db, User, Event, Athlete, ScoreEntry, ScoreSignature, ensure_round_entries, invalidate_poll_cache


class FinishStatusTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()
        self.admin = User(username="adm", role="admin", password_hash="!UNSET!")
        db.session.add(self.admin); db.session.flush()
        self.e = Event(name="E", event_group="ทั่วไป", category="ชาย", competition_date=date(2026, 11, 21), location="",
                       lane_count=2, direct_qualifiers=1, has_round_two=False, next_round_label="รอบ 4 คน")
        db.session.add(self.e); db.session.flush()
        self.a = Athlete(event_id=self.e.id, bib_no="1", name="A1", affiliation="T", start_order=1, lane_no=1, lane_order=1, status="active")
        db.session.add(self.a); db.session.flush()
        ensure_round_entries(self.a.id, 1)
        for en in ScoreEntry.query.filter_by(athlete_id=self.a.id, round_no=1):
            en.score, en.is_scored = 3, True
        db.session.add(ScoreSignature(athlete_id=self.a.id, round_no=1, started_at=datetime.utcnow()))
        db.session.commit()
        g.pop("_login_user", None)
        self.c = app.test_client()
        with self.c.session_transaction() as s:
            s["_user_id"] = str(self.admin.id); s["_fresh"] = True
        page = self.c.get("/login").get_data(as_text=True)
        self.c.environ_base["HTTP_X_CSRF_TOKEN"] = re.search(r'name="csrf-token" content="([^"]+)"', page).group(1)

    def tearDown(self):
        db.session.remove(); self.ctx.pop()

    def row(self):
        g.pop("shooting_request_cache", None)
        invalidate_poll_cache(self.e.id)
        return self.c.get(f"/events/{self.e.id}/overview-data?round=1").get_json()[0]

    def test_all_shots_but_not_finished_is_waiting_for_finish(self):
        r = self.row()
        self.assertTrue(r["shots_complete"]); self.assertFalse(r["score_complete"]); self.assertFalse(r["approved"])

    def test_admin_without_names_or_code_cannot_finish(self):
        self.c.post(f"/athletes/{self.a.id}/scorecard?round=1", data={"action": "finish"})
        self.assertFalse(self.row()["score_complete"])

    def test_typed_names_finish_and_approve(self):
        self.c.post(f"/athletes/{self.a.id}/scorecard?round=1", data={
            "action": "finish", "recorder_name_1": "Rec", "referee_name_1": "Ref", "athlete_name_1": "Ath"})
        r = self.row()
        self.assertTrue(r["score_complete"]); self.assertTrue(r["approved"])

    def test_bypass_code_finishes_without_approval(self):
        self.c.post(f"/athletes/{self.a.id}/scorecard?round=1", data={"action": "finish", "bypass_code": "7929"})
        r = self.row()
        self.assertTrue(r["score_complete"]); self.assertFalse(r["approved"])


if __name__ == "__main__":
    unittest.main()
