"""ซ่อนอีเวนต์จากรายการหน้าแรก. Run: python -m unittest tests.test_event_hidden -v"""
import os, re, tempfile, unittest
from datetime import date

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-hidden-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from flask import g
from app import app, db, User, Event


class EventHiddenTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()
        self.admin = User(username="adm", role="admin", password_hash="!UNSET!"); db.session.add(self.admin)
        self.events = []
        for i in range(3):
            e = Event(name=f"EVENT-{i}", event_group="g", category="ชาย", competition_date=date(2026, 11, 20), location="", lane_count=4)
            db.session.add(e); self.events.append(e)
        db.session.commit()

    def tearDown(self):
        db.session.remove(); self.ctx.pop()

    def client(self, login=True):
        g.pop("_login_user", None)
        c = app.test_client()
        if login:
            with c.session_transaction() as s:
                s["_user_id"] = str(self.admin.id); s["_fresh"] = True
        page = c.get("/login").get_data(as_text=True)
        token = re.search(r'name="csrf-token" content="([^"]+)"', page).group(1)
        return c, token

    def test_hide_show_only(self):
        c, token = self.client()
        a, b, d = self.events
        r = c.post("/events/visibility", data={"csrf_token": token, "ids": str(a.id), "action": "hide", "next": "/"})
        self.assertEqual(r.status_code, 302)
        g.pop("_login_user", None)
        html = c.get("/").get_data(as_text=True)
        self.assertNotIn("EVENT-0", html); self.assertIn("EVENT-1", html); self.assertIn("ซ่อนไว้ 1 อีเวนต์", html)
        self.assertIn("EVENT-0", c.get("/?show_hidden=1").get_data(as_text=True))
        # แสดงเฉพาะที่เลือก
        c.post("/events/visibility", data={"csrf_token": token, "ids": f"{a.id},{d.id}", "action": "only"})
        self.assertEqual({e.name for e in Event.query.all() if not e.is_hidden}, {"EVENT-0", "EVENT-2"})
        # ผู้ชม (ไม่ล็อกอิน) ไม่เห็นที่ซ่อน และ show_hidden ใช้ไม่ได้
        anon, _ = self.client(login=False)
        html = anon.get("/?show_hidden=1").get_data(as_text=True)
        self.assertNotIn("EVENT-1", html); self.assertIn("EVENT-2", html)
        self.assertEqual(anon.get(f"/events/{b.id}/overview?round=1").status_code, 200)  # ลิงก์ตรงยังเปิดได้
        api = anon.get("/api/public/shooting/events").get_json()
        self.assertEqual({e["name"] for e in api["events"]}, {"EVENT-0", "EVENT-2"})
        # ผู้ชมกดซ่อนไม่ได้
        anon.post("/events/visibility", data={"ids": str(d.id), "action": "hide"})
        db.session.expire_all()
        self.assertFalse(db.session.get(Event, d.id).is_hidden)


if __name__ == "__main__":
    unittest.main()
