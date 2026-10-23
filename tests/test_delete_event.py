"""ลบอีเวนต์ที่มีข้อมูลผูกอยู่ (บังคับ foreign key เหมือน Postgres). Run: python -m unittest tests.test_delete_event -v"""
import os, re, tempfile, unittest
from datetime import date, datetime

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-delete-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from flask import g
from sqlalchemy import event as sa_event
from app import (app, db, User, Event, Athlete, ScoreEntry, ScoreSignature, ScoreEditLog, TieBreakEntry, ScorecardLock,
                 AthletePenalty, BracketMatch, ResultsApprovedSetting, ensure_round_entries)


class DeleteEventTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()
        db.session.remove()
        self._fk = lambda dbapi_con, rec: dbapi_con.execute("PRAGMA foreign_keys=ON")
        sa_event.listen(db.engine, "connect", self._fk)
        db.engine.dispose()
        self.boss = User(username="boss", role="superadmin", password_hash="!UNSET!")
        db.session.add(self.boss); db.session.flush()
        self.e = Event(name="E", event_group="ทั่วไป", category="ชาย", competition_date=date(2026, 10, 23), location="",
                       lane_count=2, direct_qualifiers=1, has_round_two=True, round_two_cutoff_rank=2,
                       next_round_label="รอบ 8 คน", round_two_advancers=1, created_by=self.boss.id)
        db.session.add(self.e); db.session.flush()
        court = User(username=f"event{self.e.id}_court01", role="court", court_no=1, court_event_id=self.e.id, password_hash="!UNSET!")
        db.session.add(court); db.session.flush()
        a = Athlete(event_id=self.e.id, bib_no="1", name="A", affiliation="T", start_order=1, lane_no=1, lane_order=1,
                    round_two_disabled_by=court.id)
        db.session.add(a); db.session.flush()
        ensure_round_entries(a.id, 1)
        db.session.add_all([
            ScoreSignature(athlete_id=a.id, round_no=1, started_at=datetime.utcnow()),
            TieBreakEntry(athlete_id=a.id, round_no=1, station_no=1, score=3),
            ScoreEditLog(athlete_id=a.id, round_no=1, station_no=1, distance_m=6, edited_by=court.id, editor_username="c1"),
            ScorecardLock(athlete_id=a.id, round_no=1, user_id=court.id, username="c1", page_token="x"),
            AthletePenalty(athlete_id=a.id, round_no=1, level=1),
            BracketMatch(event_id=self.e.id, round_name="QF", match_no=1, athlete_a_id=a.id, winner_id=a.id),
            ResultsApprovedSetting(event_id=self.e.id),
        ])
        db.session.commit()
        self.court_id = court.id

    def tearDown(self):
        db.session.remove()
        sa_event.remove(db.engine, "connect", self._fk)
        db.engine.dispose()
        self.ctx.pop()

    def client(self):
        g.pop("_login_user", None)
        c = app.test_client()
        with c.session_transaction() as s:
            s["_user_id"] = str(self.boss.id); s["_fresh"] = True
        page = c.get("/login").get_data(as_text=True)
        c.environ_base["HTTP_X_CSRF_TOKEN"] = re.search(r'name="csrf-token" content="([^"]+)"', page).group(1)
        return c

    def test_delete_event_with_everything(self):
        r = self.client().post(f"/events/{self.e.id}/delete")
        self.assertEqual(r.status_code, 302)
        db.session.expire_all()
        self.assertEqual(Event.query.count(), 0)
        for m in (Athlete, ScoreEntry, ScoreSignature, ScoreEditLog, TieBreakEntry, ScorecardLock, AthletePenalty,
                  BracketMatch, ResultsApprovedSetting):
            self.assertEqual(m.query.count(), 0, m.__name__)
        self.assertEqual(User.query.filter_by(role="court").count(), 0)
        self.assertEqual(User.query.count(), 1)

    def test_delete_court_user_keeps_log(self):
        r = self.client().post(f"/events/{self.e.id}/courts/{self.court_id}/delete")
        self.assertEqual(r.status_code, 302)
        db.session.expire_all()
        log = ScoreEditLog.query.one()
        self.assertIsNone(log.edited_by); self.assertEqual(log.editor_username, "c1")


if __name__ == "__main__":
    unittest.main()
