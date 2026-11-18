import os, re, tempfile, unittest
from datetime import date

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-athlete-edit-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from flask import g
from app import app, db, User, Event, Athlete, ScoreEntry


class AthleteEditTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()
        self.admin = User(username="adm", role="admin", password_hash="!UNSET!")
        self.event = Event(name="World", event_group="World", category="ชาย", competition_date=date(2026, 11, 20), location="", lane_count=2)
        db.session.add_all([self.admin, self.event]); db.session.flush()
        self.a = Athlete(event_id=self.event.id, bib_no="1", name="Algeria", affiliation="Algeria (ALG)", start_order=1, lane_no=1, lane_order=1)
        self.b = Athlete(event_id=self.event.id, bib_no="2", name="Andorra", affiliation="Andorra (AND)", start_order=2, lane_no=2, lane_order=1)
        db.session.add_all([self.a, self.b]); db.session.commit()
        g.pop("_login_user", None)
        self.client = app.test_client()
        with self.client.session_transaction() as s:
            s["_user_id"] = str(self.admin.id); s["_fresh"] = True
        page = self.client.get("/login").get_data(as_text=True)
        self.token = re.search(r'name="csrf-token" content="([^"]+)"', page).group(1)

    def tearDown(self):
        db.session.remove(); self.ctx.pop()

    def test_update_name_country_and_swap_start_order(self):
        response = self.client.post(f"/athletes/{self.a.id}/edit", data={
            "csrf_token": self.token, "name": "Player A", "affiliation": "Algeria (ALG)", "start_order": "2",
        })
        self.assertEqual(response.status_code, 302)
        db.session.expire_all()
        a, b = db.session.get(Athlete, self.a.id), db.session.get(Athlete, self.b.id)
        self.assertEqual((a.name, a.start_order, a.lane_no), ("Player A", 2, 2))
        self.assertEqual((b.start_order, b.lane_no), (1, 1))

    def test_order_change_blocked_after_scoring_begins(self):
        db.session.add(ScoreEntry(athlete_id=self.a.id, round_no=1, station_no=1, distance_m=6, score=5, is_scored=True))
        db.session.commit()
        response = self.client.post(f"/athletes/{self.a.id}/edit", data={
            "csrf_token": self.token, "name": "Player A", "affiliation": "Algeria (ALG)", "start_order": "2",
        })
        self.assertEqual(response.status_code, 200)
        db.session.expire_all()
        self.assertEqual(db.session.get(Athlete, self.a.id).start_order, 1)


if __name__ == "__main__":
    unittest.main()
