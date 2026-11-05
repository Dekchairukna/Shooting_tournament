"""จัดการผู้ใช้: แก้ไข / ลบ. Run: python -m unittest tests.test_users_manage -v"""
import os, re, tempfile, unittest
_test_dir = tempfile.TemporaryDirectory(prefix="shooting-users-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")
from flask import g
from app import app, db, User, ScoreEditLog


class UsersManageTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()
        self.boss = User(username="boss", role="superadmin"); self.boss.set_password("Boss-pass-123")
        self.a = User(username="adm1", role="admin"); self.a.set_password("Adm1-pass-123")
        db.session.add_all([self.boss, self.a]); db.session.commit()

    def tearDown(self):
        db.session.remove(); self.ctx.pop()

    def client(self):
        g.pop("_login_user", None)
        c = app.test_client()
        with c.session_transaction() as s:
            s["_user_id"] = str(self.boss.id); s["_fresh"] = True
        page = c.get("/admin/users").get_data(as_text=True)
        self.token = re.search(r'name="csrf-token" content="([^"]+)"', page).group(1)
        return c

    def post(self, c, **data):
        data["csrf_token"] = self.token
        return c.post("/admin/users", data=data)

    def test_edit_and_delete(self):
        c = self.client()
        self.post(c, action="edit_user", user_id=self.a.id, username="referee1", role="user")
        db.session.expire_all()
        u = db.session.get(User, self.a.id)
        self.assertEqual((u.username, u.role), ("referee1", "user"))
        # ชื่อซ้ำไม่ได้
        self.post(c, action="edit_user", user_id=self.a.id, username="boss", role="user")
        db.session.expire_all(); self.assertEqual(db.session.get(User, self.a.id).username, "referee1")
        # ลดสิทธิ์ superadmin คนสุดท้าย / ลบตัวเองไม่ได้
        self.post(c, action="edit_user", user_id=self.boss.id, username="boss", role="admin")
        self.post(c, action="delete_user", user_id=self.boss.id)
        db.session.expire_all(); self.assertEqual(db.session.get(User, self.boss.id).role, "superadmin")
        # ลบได้ แม้มีประวัติแก้คะแนนอ้างถึง
        db.session.add(ScoreEditLog(athlete_id=1, round_no=1, station_no=1, distance_m=6, edited_by=self.a.id)); db.session.commit()
        self.post(c, action="delete_user", user_id=self.a.id)
        db.session.expire_all()
        self.assertIsNone(db.session.get(User, self.a.id))
        self.assertEqual(ScoreEditLog.query.first().editor_username, "referee1")


if __name__ == "__main__":
    unittest.main()
