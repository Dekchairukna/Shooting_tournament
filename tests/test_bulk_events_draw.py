"""Bulk event creation + simultaneous draw.

Run: python -m unittest tests.test_bulk_events_draw -v
"""
import os
import tempfile
import unittest
from datetime import date
from io import BytesIO

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-bulk-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from openpyxl import Workbook

from app import app, db, User, Event, Athlete, ScoreEntry, BracketMatch, parse_entry_list_workbook

SHARED = dict(competition_date="2026-10-08", location="สกลนคร", next_round_label="รอบ 8 คน",
              direct_qualifiers="4", has_round_two="yes", round_two_cutoff_rank="16",
              round_two_advancers="4", default_lane_count="3")


def entry_workbook() -> BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.append(["การแข่งขันกีฬาเปตอง"])
    ws.append(["ประเภทบุคคลชาย รุ่นอายุ 12 ปี"])
    ws.append(["ที่", "ชื่ออปท.", "อำเภอ", "จังหวัด"])
    ws.append([1, "เทศบาล A", "เมือง", "ขอนแก่น"])
    ws.append([])
    ws.append(["ประเภทชู้ตติ้งชาย รุ่นอายุ 12 ปี"])
    ws.append(["ที่", "ชื่ออปท.", "อำเภอ", "จังหวัด"])
    for i in range(1, 14):
        ws.append([i, f"เทศบาลชาย {i}", "เมือง", f"จังหวัด {i % 3}"])
    ws.append(["ประเภทชู้ตติ้งหญิง รุ่นอายุ 14 ปี"])
    ws.append(["ที่", "ชื่ออปท.", "อำเภอ", "จังหวัด"])
    for i in range(1, 8):
        ws.append([i, f"เทศบาลหญิง {i}", "เมือง", "อุดรธานี"])
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out


class _File:
    def __init__(self, path):
        self.filename = path
        self._f = open(path, "rb")

    def read(self):
        return self._f.read()


class BulkEventsDrawTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context()
        self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn)
            db.metadata.create_all(bind=conn)
            conn.commit()
        user = User(username="boss", role="superadmin", password_hash="!UNSET!")
        db.session.add(user)
        db.session.commit()
        self.client = app.test_client()
        with self.client.session_transaction() as s:
            s["_user_id"] = str(user.id)
            s["_fresh"] = True
        import re
        page = self.client.get("/events/draw").get_data(as_text=True)
        token = re.search(r'name="csrf-token" content="([^"]+)"', page).group(1)
        self.client.environ_base["HTTP_X_CSRF_TOKEN"] = token

    def tearDown(self):
        db.session.remove()
        self.ctx.pop()

    def test_pages_render(self):
        for url in ["/events/bulk-new", "/events/draw", "/"]:
            self.assertEqual(self.client.get(url).status_code, 200, url)

    def test_bulk_create_from_table(self):
        data = dict(SHARED, row_name=["ชู้ตติ้งชาย 12", "ชู้ตติ้งหญิง 12", ""],
                    row_group=["รุ่นอายุ 12 ปี", "รุ่นอายุ 12 ปี", "x"],
                    row_category=["ชาย", "หญิง", "ชาย"], row_lanes=["5", "", "1"])
        r = self.client.post("/events/bulk-new", data=data)
        self.assertEqual(r.status_code, 302)
        self.assertIn("/events/draw", r.headers["Location"])
        events = Event.query.order_by(Event.id).all()
        self.assertEqual([e.name for e in events], ["ชู้ตติ้งชาย 12", "ชู้ตติ้งหญิง 12"])
        self.assertEqual([e.lane_count for e in events], [5, 3])
        self.assertEqual(events[0].event_group, "รุ่นอายุ 12 ปี")
        self.assertEqual(events[1].category, "หญิง")
        self.assertEqual(events[0].competition_date, date(2026, 10, 8))
        # court IDs pre-created per event
        self.assertEqual(User.query.filter_by(role="court", court_event_id=events[0].id).count(), 5)
        # edit form keeps the custom age group
        page = self.client.get(f"/events/{events[0].id}/edit").get_data(as_text=True)
        self.assertIn('value="รุ่นอายุ 12 ปี" selected', page)

    def test_parse_entry_list_keyword(self):
        f = entry_workbook()
        f.filename = "entries.xlsx"
        events = parse_entry_list_workbook(f, "ชู้ตติ้ง")
        self.assertEqual([e["title"] for e in events], ["ประเภทชู้ตติ้งชาย รุ่นอายุ 12 ปี", "ประเภทชู้ตติ้งหญิง รุ่นอายุ 14 ปี"])
        self.assertEqual(len(events[0]["entries"]), 13)
        self.assertEqual(events[1]["category"], "หญิง")
        self.assertEqual(events[1]["event_group"], "รุ่นอายุ 14 ปี")
        self.assertEqual(events[1]["entries"][0], {"name": "เทศบาลหญิง 1", "affiliation": "เทศบาลหญิง 1"})

    def test_parse_person_and_team_columns(self):
        wb = Workbook(); ws = wb.active
        ws.append(["ประเภทชู้ตติ้งชาย รุ่นอายุ 14 ปี"])
        ws.append(["ที่", "ชื่อ-สกุล", "สังกัด", "จังหวัด"])
        ws.append([1, "ด.ช.สมชาย ใจดี", "ทต.หนองบัวแดง", "ชัยภูมิ"])
        out = BytesIO(); wb.save(out); out.seek(0); out.filename = "x.xlsx"
        e = parse_entry_list_workbook(out, "")[0]
        self.assertEqual(e["entries"][0], {"name": "ด.ช.สมชาย ใจดี", "affiliation": "ทต.หนองบัวแดง"})

    def test_import_preview_then_confirm_and_draw(self):
        r = self.client.post("/events/bulk-import", data={
            "entry_file": (entry_workbook(), "entries.xlsx"), "keyword": "ชู้ตติ้ง", "per_lane": "6",
            "default_lane_count": "4", "name_prefix": "อปท.41"}, content_type="multipart/form-data")
        self.assertEqual(r.status_code, 200)
        html = r.get_data(as_text=True)
        self.assertIn("อปท.41 ประเภทชู้ตติ้งชาย รุ่นอายุ 12 ปี", html)
        import re, html as h
        events_json = h.unescape(re.search(r'<textarea name="events_json" hidden>(.*?)</textarea>', html, re.S).group(1))
        r = self.client.post("/events/bulk-import/confirm", data=dict(
            SHARED, events_json=events_json, include=["0", "1"], name_0="ชู้ตติ้งชาย 12", lanes_0="3",
            draw_now="yes"))
        self.assertEqual(r.status_code, 302)
        events = Event.query.order_by(Event.id).all()
        self.assertEqual(len(events), 2)
        self.assertEqual(events[0].name, "ชู้ตติ้งชาย 12")
        self.assertEqual(events[0].lane_count, 3)
        self.assertEqual(events[1].lane_count, 2)  # ceil(7/6)
        a = Athlete.query.filter_by(event_id=events[0].id).all()
        self.assertEqual(len(a), 13)
        self.assertEqual(sorted(x.start_order for x in a), list(range(1, 14)))
        self.assertEqual({x.lane_no for x in a}, {1, 2, 3})
        page = self.client.get(r.headers["Location"]).get_data(as_text=True)
        self.assertIn("ผลการจับสลาก", page)

    def _event_with_athletes(self, name, n, lanes=3):
        e = Event(name=name, event_group="ทั่วไป", category="ชาย", competition_date=date(2026, 10, 8),
                  location="", lane_count=lanes, direct_qualifiers=4, has_round_two=True,
                  round_two_cutoff_rank=16, next_round_label="รอบ 8 คน", round_two_advancers=4)
        db.session.add(e)
        db.session.flush()
        for i in range(1, n + 1):
            db.session.add(Athlete(event_id=e.id, bib_no=str(i), name=f"{name}-{i}", affiliation="T",
                                   start_order=i, lane_no=((i - 1) % lanes) + 1, lane_order=((i - 1) // lanes) + 1))
        db.session.commit()
        return e

    def test_draw_many_at_once_and_lock_started(self):
        e1 = self._event_with_athletes("E1", 30)
        e2 = self._event_with_athletes("E2", 30)
        e3 = self._event_with_athletes("E3", 9)
        started = Athlete.query.filter_by(event_id=e3.id).first()
        db.session.add(ScoreEntry(athlete_id=started.id, round_no=1, station_no=1, distance_m=6, score=5, is_scored=True))
        db.session.add(BracketMatch(event_id=e3.id, round_name="QF", match_no=1))
        db.session.commit()
        before = {e.id: [a.name for a in Athlete.query.filter_by(event_id=e.id).order_by(Athlete.start_order)] for e in (e1, e2, e3)}

        r = self.client.post("/events/draw", data={"event_ids": [str(e1.id), str(e2.id), str(e3.id)]})
        self.assertEqual(r.status_code, 302)
        after = {e.id: [a.name for a in Athlete.query.filter_by(event_id=e.id).order_by(Athlete.start_order)] for e in (e1, e2, e3)}
        self.assertNotEqual(before[e1.id], after[e1.id])
        self.assertNotEqual(before[e2.id], after[e2.id])
        self.assertEqual(before[e3.id], after[e3.id])  # locked: scores exist
        for eid in (e1.id, e2.id):
            rows = Athlete.query.filter_by(event_id=eid).order_by(Athlete.start_order).all()
            self.assertEqual([a.start_order for a in rows], list(range(1, 31)))
            self.assertEqual([a.lane_no for a in rows], [((i) % 3) + 1 for i in range(30)])

        # superadmin force redraw resets bracket
        self.client.post("/events/draw", data={"event_ids": [str(e3.id)], "force": "yes"})
        self.assertEqual(BracketMatch.query.filter_by(event_id=e3.id).count(), 0)

    def test_draw_center_lists_status(self):
        self._event_with_athletes("Ready", 5)
        Event.query.count()
        page = self.client.get("/events/draw").get_data(as_text=True)
        self.assertIn("พร้อมจับ", page)

    def test_real_attachment_if_present(self):
        path = os.environ.get("ENTRY_XLS")
        if not path or not os.path.exists(path):
            self.skipTest("set ENTRY_XLS to a real entry list")
        events = parse_entry_list_workbook(_File(path), "ชู้ตติ้ง")
        for e in events:
            print(e["title"], e["category"], e["event_group"], len(e["entries"]), e["entries"][0])
        self.assertGreaterEqual(len(events), 8)


if __name__ == "__main__":
    unittest.main()
