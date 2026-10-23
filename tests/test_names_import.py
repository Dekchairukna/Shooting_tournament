"""นำเข้าชื่อนักกีฬา. Run: python -m unittest tests.test_names_import -v"""
import os, re, tempfile, unittest
from datetime import date
from io import BytesIO

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-names-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from flask import g
from openpyxl import Workbook
from app import app, db, User, Event, Athlete


class NamesImportTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()
        self.admin = User(username="adm", role="admin", password_hash="!UNSET!"); db.session.add(self.admin)
        e = Event(name="E", event_group="ม.ต้น", category="ชาย", competition_date=date(2026, 10, 24), location="", lane_count=6,
                  direct_qualifiers=4, has_round_two=True, round_two_cutoff_rank=16, next_round_label="รอบ 8 คน", round_two_advancers=4)
        db.session.add(e); db.session.flush()
        self.a = Athlete(event_id=e.id, bib_no="1", name="สาธิตเกษตร", affiliation="สาธิตเกษตร", start_order=3, lane_no=3, lane_order=1)
        self.b = Athlete(event_id=e.id, bib_no="2", name="สาธิตพะเยา", affiliation="สาธิตพะเยา", start_order=4, lane_no=4, lane_order=1)
        db.session.add_all([self.a, self.b]); db.session.commit()

    def tearDown(self):
        db.session.remove(); self.ctx.pop()

    def test_import(self):
        g.pop("_login_user", None)
        c = app.test_client()
        with c.session_transaction() as s:
            s["_user_id"] = str(self.admin.id); s["_fresh"] = True
        page = c.get("/athletes/names-import").get_data(as_text=True)
        token = re.search(r'name="csrf-token" content="([^"]+)"', page).group(1)
        wb = Workbook(); ws = wb.active
        ws.append(["หัวเรื่อง"]); ws.append([])
        ws.append(["athlete_id", "อีเวนต์", "ลำดับตี", "สนาม", "สังกัด", "ชื่อนักกีฬา (ใส่ระบบ)"])
        ws.append([self.a.id, "E", 3, "3/1", "สาธิตเกษตร", "ด.ช.ไทกล้า  สายสุนทร"])
        ws.append([self.b.id, "E", 4, "4/1", "สาธิตพะเยา", ""])
        ws.append([99999, "E", 5, "", "", "ใครก็ไม่รู้"])
        buf = BytesIO(); wb.save(buf); buf.seek(0)
        r = c.post("/athletes/names-import", data={"file": (buf, "n.xlsx"), "csrf_token": token}, content_type="multipart/form-data")
        self.assertEqual(r.status_code, 200)
        db.session.expire_all()
        a = db.session.get(Athlete, self.a.id); b = db.session.get(Athlete, self.b.id)
        self.assertEqual(a.name, "ด.ช.ไทกล้า สายสุนทร"); self.assertEqual(a.affiliation, "สาธิตเกษตร"); self.assertEqual(a.start_order, 3)
        self.assertEqual(b.name, "สาธิตพะเยา")
        self.assertIn("99999", r.get_data(as_text=True))


if __name__ == "__main__":
    unittest.main()
