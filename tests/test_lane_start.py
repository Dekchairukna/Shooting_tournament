"""สนามแรกของอีเวนต์ + ชื่อสนามตัวอักษร. Run: python -m unittest tests.test_lane_start -v"""
import os, re, tempfile, unittest
from datetime import date

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-lanestart-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from flask import g
from app import app, db, User, Event, Athlete, court_name, round2_lane_slot

SCHOOLS = ["ปทุมวัน", "ขอนแก่น", "องครักษ์", "นเรศวร", "พะเยา", "รามมัธยม", "สงขลามัธยม", "ประสานมิตรมัธยม",
           "มหาสารคามมัธยม", "จุฬาฯมัธยม", "ศิลปากร", "มข.หนองคาย"]


class LaneStartTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()
        self.boss = User(username="boss", role="superadmin", password_hash="!UNSET!"); db.session.add(self.boss)
        self.e = Event(name="ม.ปลาย หญิง", event_group="ม.ปลาย", category="หญิง", competition_date=date(2026, 10, 24), location="",
                       lane_count=6, direct_qualifiers=4, has_round_two=True, round_two_cutoff_rank=16,
                       next_round_label="รอบ 8 คน", round_two_advancers=4)
        db.session.add(self.e); db.session.flush()
        for i, sch in enumerate(SCHOOLS, start=1):
            db.session.add(Athlete(event_id=self.e.id, bib_no=str(i), name=sch, affiliation=sch, start_order=i,
                                   lane_no=(i - 1) % 6 + 1, lane_order=(i - 1) // 6 + 1))
        db.session.commit()

    def tearDown(self):
        db.session.remove(); self.ctx.pop()

    def client(self):
        g.pop("_login_user", None)
        c = app.test_client()
        with c.session_transaction() as s:
            s["_user_id"] = str(self.boss.id); s["_fresh"] = True
        page = c.get("/login").get_data(as_text=True)
        c.environ_base["HTTP_X_CSRF_TOKEN"] = re.search(r'name="csrf-token" content="([^"]+)"', page).group(1)
        return c

    def test_bulk_settings_lanes_match_slide(self):
        c = self.client()
        r = c.post("/events/bulk-settings", data={"ids": str(self.e.id), "apply_lanes": "1", "lane_count": "3", "lane_start": "5",
                                                  "apply_court_letters": "1", "court_letters": "yes"})
        self.assertEqual(r.status_code, 302)
        db.session.expire_all()
        got = {a.name: (court_name(a.event, a.lane_no), a.lane_order, a.start_order)
               for a in Athlete.query.filter_by(event_id=self.e.id)}
        # ตามสไลด์: E = ปทุมวัน นเรศวร สงขลา จุฬาฯ · F = ขอนแก่น พะเยา ประสานมิตร ศิลปากร · G = องครักษ์ ราม มหาสารคาม มข.หนองคาย
        self.assertEqual([n for n, v in sorted(got.items(), key=lambda kv: kv[1][1]) if v[0] == "E"], ["ปทุมวัน", "นเรศวร", "สงขลามัธยม", "จุฬาฯมัธยม"])
        self.assertEqual([n for n, v in sorted(got.items(), key=lambda kv: kv[1][1]) if v[0] == "F"], ["ขอนแก่น", "พะเยา", "ประสานมิตรมัธยม", "ศิลปากร"])
        self.assertEqual([n for n, v in sorted(got.items(), key=lambda kv: kv[1][1]) if v[0] == "G"], ["องครักษ์", "รามมัธยม", "มหาสารคามมัธยม", "มข.หนองคาย"])
        self.assertEqual(got["ปทุมวัน"][2], 1)  # ลำดับตีไม่เปลี่ยน
        courts = sorted(u.court_no for u in User.query.filter_by(role="court", court_event_id=self.e.id))
        self.assertEqual(courts, [5, 6, 7])
        html = c.get(f"/events/{self.e.id}/court?round=1&court=6").get_data(as_text=True)
        self.assertIn("สนาม F", html); self.assertIn("ขอนแก่น", html); self.assertNotIn("ปทุมวัน</", html)
        ov = c.get(f"/events/{self.e.id}/overview?round=1").get_data(as_text=True)
        self.assertIn('<td class="lane-no">E</td>', ov)
        e = db.session.get(Event, self.e.id)
        self.assertEqual(round2_lane_slot(e, 1, 6), (5, 1)); self.assertEqual(round2_lane_slot(e, 4, 6), (5, 2))

    def test_edit_event_form(self):
        c = self.client()
        data = {"name": "ม.ปลาย หญิง", "event_group": "ม.ปลาย", "category": "หญิง", "competition_date": "2026-10-24", "lane_count": "3",
                "direct_qualifiers": "4", "has_round_two": "yes", "round_two_cutoff_rank": "16", "round_two_mode": "cutoff",
                "next_round_label": "รอบ 8 คน", "round_two_advancers": "4", "lane_start": "5", "court_letters": "yes"}
        c.post(f"/events/{self.e.id}/edit", data=data)
        db.session.expire_all()
        a = Athlete.query.filter_by(event_id=self.e.id, start_order=2).one()
        self.assertEqual((a.lane_no, a.lane_order), (6, 1))


if __name__ == "__main__":
    unittest.main()
