"""Court chain (ต่อคิวสนามข้ามรุ่น). Run: python -m unittest tests.test_court_chain -v"""
import os, re, tempfile, unittest
from datetime import date

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-chain-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from app import app, db, User, Event, Athlete, ScoreEntry, CourtChain, sync_event_court_users


class ChainTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()
        self.boss = User(username="boss", role="superadmin", password_hash="!UNSET!")
        db.session.add(self.boss); db.session.commit()
        self.client = self._login(self.boss)
        self.m12 = self._event("ชู้ตติ้งชาย 12", 45)
        self.f12 = self._event("ชู้ตติ้งหญิง 12", 34)
        self.m14 = self._event("ชู้ตติ้งชาย 14", 33)

    def tearDown(self):
        db.session.remove(); self.ctx.pop()

    def _login(self, user):
        from flask import g
        g.pop("_login_user", None)  # เทสต์เปิด app context ค้างไว้ ล้างผู้ใช้ที่ Flask-Login แคชไว้
        c = app.test_client()
        with c.session_transaction() as s:
            s["_user_id"] = str(user.id); s["_fresh"] = True
        page = c.get("/login").get_data(as_text=True)
        c.environ_base["HTTP_X_CSRF_TOKEN"] = re.search(r'name="csrf-token" content="([^"]+)"', page).group(1)
        return c

    def _event(self, name, n, lanes=4):
        e = Event(name=name, event_group="รุ่นอายุ 12 ปี", category="ชาย", competition_date=date(2026, 10, 8), location="",
                  lane_count=lanes, direct_qualifiers=4, has_round_two=True, round_two_cutoff_rank=16,
                  next_round_label="รอบ 8 คน", round_two_advancers=4)
        db.session.add(e); db.session.flush()
        for i in range(1, n + 1):
            db.session.add(Athlete(event_id=e.id, bib_no=str(i), name=f"{name}-{i}", affiliation="T", start_order=i,
                                   lane_no=(i - 1) % lanes + 1, lane_order=(i - 1) // lanes + 1))
        sync_event_court_users(e)
        db.session.commit()
        return e

    def _chain(self, ids, lanes=8, **extra):
        data = {"name": "ชู้ตติ้ง รอบแรก", "lane_count": str(lanes), "event_ids": [str(i) for i in ids]}
        for pos, i in enumerate(ids, 1):
            data[f"pos_{i}"] = str(pos)
        data.update(extra)
        return self.client.post("/chains/new", data=data)

    def lanes(self, e):
        return [(a.lane_no, a.lane_order) for a in Athlete.query.filter_by(event_id=e.id).order_by(Athlete.start_order)]

    def test_fills_empty_courts_with_next_category(self):
        self.assertEqual(self._chain([self.m12.id, self.f12.id, self.m14.id]).status_code, 302)
        m, f, m14 = self.lanes(self.m12), self.lanes(self.f12), self.lanes(self.m14)
        self.assertEqual(m[-1], (5, 6))                       # ชาย 12 คนที่ 45: สนาม 5 คิว 6
        self.assertEqual(f[:3], [(6, 6), (7, 6), (8, 6)])     # หญิง 12 เริ่มสนาม 6–8 คิว 6
        self.assertEqual(f[-1], (7, 10))                      # คนที่ 79 รวม → คิว 10 สนาม 7
        self.assertEqual(m14[0], (8, 10))                     # ชาย 14 เริ่มสนาม 8 คิว 10
        for e in (self.m12, self.f12, self.m14):
            db.session.refresh(e)
            self.assertEqual(e.lane_count, 8)
            self.assertEqual(User.query.filter_by(role="court", court_event_id=e.id).count(), 8)
        html = self.client.get("/chains").get_data(as_text=True)
        self.assertIn("14 รอบคิว", html)        # 112 คน / 8 = 14
        self.assertIn("แยกอีเวนต์ 16 รอบ", html)  # 6 + 5 + 5

    def test_draw_keeps_chain_and_separate_events_untouched(self):
        other = self._event("แยกเดี่ยว", 10, lanes=4)
        before_other = self.lanes(other)
        self._chain([self.m12.id, self.f12.id])
        r = self.client.post("/events/draw", data={"event_ids": [str(self.m12.id), str(self.f12.id)]})
        self.assertEqual(r.status_code, 302)
        self.assertEqual(self.lanes(self.f12)[:3], [(6, 6), (7, 6), (8, 6)])
        self.assertEqual(self.lanes(other), before_other)
        self.assertEqual(sorted(a.start_order for a in self.f12.athletes), list(range(1, 35)))

    def test_delete_athlete_repacks_next_category(self):
        self._chain([self.m12.id, self.f12.id])
        victim = Athlete.query.filter_by(event_id=self.m12.id, start_order=1).one()
        self.client.post(f"/athletes/{victim.id}/delete")
        self.assertEqual(self.lanes(self.f12)[:4], [(5, 6), (6, 6), (7, 6), (8, 6)])

    def test_unchain_restores_separate_lanes(self):
        self._chain([self.m12.id, self.f12.id])
        chain = CourtChain.query.one()
        self.client.post(f"/chains/{chain.id}/delete")
        self.assertEqual(self.lanes(self.f12)[:2], [(1, 1), (2, 1)])
        self.assertIsNone(db.session.get(Event, self.f12.id).chain_id)

    def test_started_event_needs_confirmation(self):
        a = Athlete.query.filter_by(event_id=self.f12.id).first()
        db.session.add(ScoreEntry(athlete_id=a.id, round_no=1, station_no=1, distance_m=6, score=5, is_scored=True)); db.session.commit()
        self._chain([self.m12.id, self.f12.id])
        self.assertEqual(CourtChain.query.count(), 0)
        self._chain([self.m12.id, self.f12.id], confirm_started="yes")
        self.assertEqual(CourtChain.query.count(), 1)

    def test_court_user_sees_whole_chain_on_own_court(self):
        self._chain([self.m12.id, self.f12.id])
        court6 = User.query.filter_by(role="court", court_event_id=self.m12.id, court_no=6).one()
        c = self._login(court6)
        from flask import g
        g.pop("_login_user", None)
        html = c.get(f"/events/{self.m12.id}/court?round=1").get_data(as_text=True)
        self.assertIn("สายคิว: ชู้ตติ้ง รอบแรก", html)
        self.assertIn("ชู้ตติ้งชาย 12-6<", html)
        self.assertIn("ชู้ตติ้งหญิง 12-1<", html)     # หญิงคนแรกอยู่สนาม 6
        self.assertNotIn("ชู้ตติ้งหญิง 12-2<", html)  # คนที่ 2 อยู่สนาม 7
        girl = Athlete.query.filter_by(event_id=self.f12.id, start_order=1).one()
        g.pop("_login_user", None)
        self.assertEqual(c.get(f"/athletes/{girl.id}/scorecard?round=1").status_code, 200)
        other_court = Athlete.query.filter_by(event_id=self.f12.id, start_order=2).one()
        g.pop("_login_user", None)
        self.assertNotEqual(c.get(f"/athletes/{other_court.id}/scorecard?round=1").status_code, 200)

    def test_cannot_put_event_in_two_chains(self):
        self._chain([self.m12.id, self.f12.id])
        self._chain([self.f12.id, self.m14.id])
        self.assertEqual(CourtChain.query.count(), 1)

    def test_pages_render(self):
        self._chain([self.m12.id, self.f12.id])
        chain = CourtChain.query.one()
        for url in ["/chains", "/chains/new", f"/chains/{chain.id}/edit", "/events/draw", "/",
                    f"/events/{self.f12.id}/overview?round=1", f"/events/{self.f12.id}/athletes", f"/events/{self.f12.id}/edit"]:
            self.assertEqual(self.client.get(url).status_code, 200, url)


if __name__ == "__main__":
    unittest.main()
