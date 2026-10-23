"""รูปแบบใบบันทึกคะแนน. Run: python -m unittest tests.test_scorecard_layout -v"""
import os, re, tempfile, unittest
from datetime import date

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-layout-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from flask import g
from app import app, db, User, Event, Athlete, scorecard_level_flags


class LayoutTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()
        self.admin = User(username="adm", role="admin", password_hash="!UNSET!")
        db.session.add(self.admin)
        self.e = Event(name="ชู้ตติ้ง ประถม ชาย", event_group="ประถม", category="ชาย", competition_date=date(2026, 10, 25),
                       location="", lane_count=8, direct_qualifiers=4, has_round_two=True, round_two_cutoff_rank=16,
                       next_round_label="รอบ 8 คน", round_two_advancers=4)
        db.session.add(self.e); db.session.flush()
        self.a = Athlete(event_id=self.e.id, bib_no="7", name="ด.ช.ทดสอบ", affiliation="โรงเรียนสาธิต", start_order=1, lane_no=1, lane_order=1)
        db.session.add(self.a); db.session.commit()

    def tearDown(self):
        db.session.remove(); self.ctx.pop()

    def client(self):
        g.pop("_login_user", None)
        c = app.test_client()
        with c.session_transaction() as s:
            s["_user_id"] = str(self.admin.id); s["_fresh"] = True
        return c

    def test_standard_default(self):
        html = self.client().get(f"/athletes/{self.a.id}/scorecard-print?round=1").get_data(as_text=True)
        self.assertIn("Atelier 1", html); self.assertNotIn("s49-sheet", html)

    def test_school49_layout(self):
        self.e.scorecard_layout = "school49"; db.session.commit()
        html = self.client().get(f"/athletes/{self.a.id}/scorecard-print?round=1").get_data(as_text=True)
        self.assertIn("s49-sheet", html)
        self.assertIn("คำมอกหลวงเกมส์", html)
        self.assertIn("ด.ช.ทดสอบ", html); self.assertIn("โรงเรียนสาธิต", html)
        self.assertIn("25 ต.ค. 2569", html)
        self.assertEqual(html.count('class="s49-redbox"'), 25)
        f = scorecard_level_flags(self.e)
        self.assertTrue(f["primary"] and f["male"] and not f["female"] and not f["lower"])

    def test_layout_override_param(self):
        html = self.client().get(f"/athletes/{self.a.id}/scorecard-print?round=1&layout=school49").get_data(as_text=True)
        self.assertIn("s49-sheet", html)


if __name__ == "__main__":
    unittest.main()
