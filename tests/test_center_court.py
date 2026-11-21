"""Center Court: เลือกสนามที่แสดงชื่อเป็น Center Court. Run: python -m unittest tests.test_center_court -v"""
import os, re, tempfile, unittest
from datetime import date
from types import SimpleNamespace

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-cc-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from app import app, db, User, Event, Athlete, court_name


class CenterCourtTests(unittest.TestCase):
    def test_court_name(self):
        e = SimpleNamespace(center_court=1, court_letters=True)
        self.assertEqual(court_name(e, 1), "Center Court")
        self.assertEqual(court_name(e, 1, short=True), "CC")
        self.assertEqual(court_name(e, 2), "B")
        self.assertEqual(court_name(SimpleNamespace(center_court=None, court_letters=False), 1), "1")

    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()
        user = User(username="boss", role="superadmin", password_hash="!UNSET!")
        db.session.add(user); db.session.commit()
        self.client = app.test_client()
        with self.client.session_transaction() as s:
            s["_user_id"] = str(user.id); s["_fresh"] = True
        page = self.client.get("/events/new").get_data(as_text=True)
        self.client.environ_base["HTTP_X_CSRF_TOKEN"] = re.search(r'name="csrf-token" content="([^"]+)"', page).group(1)

    def tearDown(self):
        db.session.remove(); self.ctx.pop()

    def test_edit_form_saves_and_overview_shows_cc(self):
        e = Event(name="E", event_group="ทั่วไป", category="ชาย", competition_date=date(2026, 11, 21), location="",
                  lane_count=4, direct_qualifiers=4, has_round_two=False, next_round_label="รอบ 8 คน", court_letters=True)
        db.session.add(e); db.session.flush()
        db.session.add(Athlete(event_id=e.id, bib_no="1", name="A1", affiliation="T", start_order=1, lane_no=1, lane_order=1))
        db.session.commit()
        form = {"name": "E", "event_group": "ทั่วไป", "category": "ชาย", "competition_date": "2026-11-21", "lane_count": "4",
                "direct_qualifiers": "4", "has_round_two": "no", "next_round_label": "รอบ 8 คน", "court_letters": "yes",
                "center_court": "1", "lane_start": "1", "round_two_advancers": "4"}
        self.client.post(f"/events/{e.id}/edit", data=form)
        db.session.refresh(e)
        self.assertEqual(e.center_court, 1)
        self.assertIn('<option value="1" selected>สนาม 1 (A)</option>', self.client.get(f"/events/{e.id}/edit").get_data(as_text=True))
        self.assertIn('<td class="lane-no">CC</td>', self.client.get(f"/events/{e.id}/overview?round=1").get_data(as_text=True))
        form["center_court"] = ""
        self.client.post(f"/events/{e.id}/edit", data=form)
        db.session.refresh(e)
        self.assertIsNone(e.center_court)


if __name__ == "__main__":
    unittest.main()
