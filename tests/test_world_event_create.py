import os, re, tempfile, unittest

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-world-event-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from flask import g
from app import app, db, User, Event, Athlete, nation_flag_file


class WorldEventCreateTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()
        self.boss = User(username="boss", role="superadmin", password_hash="!UNSET!")
        db.session.add(self.boss); db.session.commit()
        g.pop("_login_user", None)
        self.client = app.test_client()
        with self.client.session_transaction() as s:
            s["_user_id"] = str(self.boss.id); s["_fresh"] = True
        page = self.client.get("/login").get_data(as_text=True)
        self.token = re.search(r'name="csrf-token" content="([^"]+)"', page).group(1)

    def tearDown(self):
        db.session.remove(); self.ctx.pop()

    def test_create_48_country_event(self):
        response = self.client.post("/events/world-48/new", data={
            "csrf_token": self.token,
            "name": "WORLD MEN",
            "event_group": "World Championship",
            "category": "ชาย",
            "competition_date": "2026-11-20",
            "lane_count": "8",
            "location": "Khon Kaen",
        })
        self.assertEqual(response.status_code, 302)
        event = Event.query.filter_by(name="WORLD MEN").one()
        athletes = Athlete.query.filter_by(event_id=event.id).order_by(Athlete.start_order).all()
        self.assertEqual(len(athletes), 48)
        self.assertEqual((athletes[0].affiliation, athletes[-1].affiliation), ("Algeria (ALG)", "United States (USA)"))
        self.assertTrue(all(nation_flag_file(a.affiliation) for a in athletes))
        self.assertEqual(sorted(a.start_order for a in athletes), list(range(1, 49)))


if __name__ == "__main__":
    unittest.main()
