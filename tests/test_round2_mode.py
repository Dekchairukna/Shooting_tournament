"""Round-2 selection mode + 5/3 highlight markup. Run: python -m unittest tests.test_round2_mode -v"""
import os, re, tempfile, unittest
from datetime import date, datetime

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-r2-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from app import app, db, User, Event, Athlete, ScoreEntry, ScoreSignature, round2_cutoff_rank, round_two_candidate_ids


class Round2ModeTests(unittest.TestCase):
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

    def _event(self, mode, n, athletes=20):
        e = Event(name="E", event_group="ทั่วไป", category="ชาย", competition_date=date(2026, 10, 8), location="",
                  lane_count=4, direct_qualifiers=4, has_round_two=True, round_two_cutoff_rank=n,
                  round_two_mode=mode, next_round_label="รอบ 8 คน", round_two_advancers=4)
        db.session.add(e); db.session.flush()
        for i in range(1, athletes + 1):
            a = Athlete(event_id=e.id, bib_no=str(i), name=f"A{i}", affiliation="T", start_order=i,
                        lane_no=(i - 1) % 4 + 1, lane_order=(i - 1) // 4 + 1, status="finished")
            db.session.add(a); db.session.flush()
            # คะแนนรวมไม่ซ้ำกัน: คนที่ i ได้ 5 คะแนน (21 - i) ช่อง ที่เหลือ 0
            fives = max(21 - i, 0)
            for k, (st, d) in enumerate([(s, d) for s in range(1, 6) for d in (6, 7, 8, 9)]):
                db.session.add(ScoreEntry(athlete_id=a.id, round_no=1, station_no=st, distance_m=d,
                                          score=5 if k < fives else 0, is_scored=True))
            db.session.add(ScoreSignature(athlete_id=a.id, round_no=1, bypass_signed=True, finished_at=datetime.utcnow()))
        db.session.commit()
        return e

    def test_cutoff_mode_unchanged(self):
        e = self._event("cutoff", 16)
        self.assertEqual(round2_cutoff_rank(e), 16)

    def test_next_mode_counts_after_direct(self):
        e = self._event("next", 16)
        self.assertEqual(round2_cutoff_rank(e), 20)
        self.assertEqual(len(round_two_candidate_ids(e)), 16)
        e2 = self._event("cutoff", 16)
        self.assertEqual(len(round_two_candidate_ids(e2)), 12)

    def test_unknown_mode_falls_back(self):
        e = self._event("weird", 16, athletes=1)
        self.assertEqual(round2_cutoff_rank(e), 16)

    def test_forms_save_mode(self):
        form = dict(name="New", event_group="ทั่วไป", category="ชาย", competition_date="2026-10-08",
                    lane_count="4", direct_qualifiers="4", has_round_two="yes", round_two_cutoff_rank="16",
                    round_two_mode="next", next_round_label="รอบ 8 คน", round_two_advancers="4")
        self.client.post("/events/new", data=form)
        e = Event.query.filter_by(name="New").one()
        self.assertEqual(e.round_two_mode, "next")
        page = self.client.get(f"/events/{e.id}/edit").get_data(as_text=True)
        self.assertIn('<option value="next" selected>', page)
        form.update(round_two_mode="cutoff")
        self.client.post(f"/events/{e.id}/edit", data=form)
        db.session.refresh(e)
        self.assertEqual(e.round_two_mode, "cutoff")
        bulk = dict(form, row_name=["B1"], row_group=["ทั่วไป"], row_category=["ชาย"], row_lanes=["4"], round_two_mode="next")
        self.client.post("/events/bulk-new", data=bulk)
        self.assertEqual(Event.query.filter_by(name="B1").one().round_two_mode, "next")

    def test_overview_cells_carry_score(self):
        e = self._event("next", 4, athletes=6)
        html = self.client.get(f"/events/{e.id}/overview?round=1").get_data(as_text=True)
        self.assertIn('data-score="5"', html)
        self.assertIn('tr.visual-finished > td.score-col[data-score="5"]', html)


if __name__ == "__main__":
    unittest.main()


class DatabaseUrlTests(unittest.TestCase):
    def test_postgres_urls_pin_psycopg2(self):
        from app import normalize_database_url as n
        self.assertEqual(n("postgres://u:p@h:5432/db"), "postgresql+psycopg2://u:p@h:5432/db")
        self.assertEqual(n("postgresql://u:p@h/db"), "postgresql+psycopg2://u:p@h/db")
        self.assertEqual(n("postgresql+psycopg://u:p@h/db"), "postgresql+psycopg://u:p@h/db")
        self.assertEqual(n("sqlite:////tmp/x.db"), "sqlite:////tmp/x.db")
        from sqlalchemy.engine import make_url
        self.assertEqual(make_url(n("postgresql://u:p@h/db")).get_dialect().driver, "psycopg2")
