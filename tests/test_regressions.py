"""Run with .venv/bin/python -m unittest discover -s tests -v.

Always select an isolated database before importing the application.
"""
import os
import tempfile
import unittest
from datetime import date

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-tests-")
os.environ["DATABASE_URL"] = "sqlite:///" + _test_dir.name + "/test.db"
os.environ["SECRET_KEY"] = "test-only-secret"

from app import app, db, User, Event, Athlete, BracketMatch
from app import ScoreEntry, ScoreSignature, ScoreEditLog, ResultsApprovedSetting, seed_defaults
from sqlalchemy.exc import IntegrityError
from concurrent.futures import ThreadPoolExecutor
from threading import Barrier
from io import BytesIO


class EventRegressionTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.context = app.app_context()
        self.context.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn)
            db.metadata.create_all(bind=conn)
            conn.commit()
            conn.exec_driver_sql("PRAGMA foreign_keys=ON")
        user = User(username="tester", role="superadmin", password_hash="!UNSET!")
        event = Event(name="Original", event_group="ทั่วไป", category="ชาย",
                      competition_date=date(2026, 9, 17), location="Original location",
                      lane_count=2, direct_qualifiers=4, has_round_two=True,
                      round_two_cutoff_rank=16, next_round_label="รอบ 8 คน",
                      round_two_advancers=4)
        db.session.add_all([user, event])
        db.session.flush()
        athlete = Athlete(event_id=event.id, bib_no="1", name="Example", affiliation="Team",
                          start_order=1, lane_no=2, lane_order=3)
        db.session.add(athlete)
        db.session.flush()
        db.session.add(BracketMatch(event_id=event.id, round_name="QF", match_no=1,
                                   athlete_a_id=athlete.id, winner_id=athlete.id))
        db.session.commit()
        self.event_id, self.athlete_id = event.id, athlete.id
        self.client = app.test_client()
        with self.client.session_transaction() as session:
            session["_user_id"] = str(user.id)
            session["_fresh"] = True
            session["csrf_token"] = "test-csrf-token"
        self.client.environ_base["HTTP_X_CSRF_TOKEN"] = "test-csrf-token"
        self.form = dict(name="Renamed", event_group="ทั่วไป", category="ชาย",
                         competition_date="2026-09-17", location="New location",
                         lane_count="2", direct_qualifiers="4", has_round_two="yes",
                         round_two_cutoff_rank="16", next_round_label="รอบ 8 คน",
                         round_two_advancers="4")

    def tearDown(self):
        db.session.remove()
        self.context.pop()

    def test_metadata_edit_preserves_winner_and_lane_assignment(self):
        response = self.client.post(f"/events/{self.event_id}/edit", data=self.form)
        self.assertEqual(response.status_code, 302)
        db.session.expire_all()
        self.assertEqual(db.session.get(Event, self.event_id).name, "Renamed")
        self.assertEqual(BracketMatch.query.filter_by(round_name="QF", match_no=1).one().winner_id, self.athlete_id)
        athlete = db.session.get(Athlete, self.athlete_id)
        self.assertEqual((athlete.lane_no, athlete.lane_order), (2, 3))

    def test_competition_settings_change_still_resets_bracket(self):
        response = self.client.post(f"/events/{self.event_id}/edit",
                                    data={**self.form, "next_round_label": "รอบ 16 คน"})
        self.assertEqual(response.status_code, 302)
        self.assertEqual(BracketMatch.query.filter(BracketMatch.winner_id.isnot(None)).count(), 0)
        self.assertEqual(BracketMatch.query.filter_by(round_name="R16").count(), 8)

    def test_invalid_event_inputs_do_not_mutate_database(self):
        for field, value in [("lane_count", "0"), ("lane_count", "bad"),
                             ("competition_date", "invalid"), ("name", " "),
                             ("direct_qualifiers", "-1"), ("round_two_advancers", "0")]:
            with self.subTest(field=field, value=value):
                form = {**self.form, field: value}
                self.assertEqual(self.client.post(f"/events/{self.event_id}/edit", data=form).status_code, 400)
                self.assertEqual(self.client.post("/events/new", data=form).status_code, 400)
                db.session.expire_all()
                self.assertEqual(Event.query.count(), 1)
                self.assertEqual(db.session.get(Event, self.event_id).name, "Original")
                self.assertEqual(BracketMatch.query.count(), 1)

    def test_main_pages_render(self):
        for path in ("/", f"/events/{self.event_id}/overview",
                     f"/events/{self.event_id}/bracket",
                     f"/athletes/{self.athlete_id}/scorecard",
                     f"/events/{self.event_id}/edit"):
            with self.subTest(path=path):
                self.assertEqual(self.client.get(path).status_code, 200)

    def score(self, **changes):
        payload = dict(round_no=1, station_no=1, distance_m=6, score=5,
                       red=False, played=True, version=0)
        payload.update(changes)
        return self.client.post(f"/api/scorecard/{self.athlete_id}/autosave", json=payload)

    def test_csrf_required_for_forms_and_json(self):
        self.client.environ_base.pop("HTTP_X_CSRF_TOKEN")
        self.assertEqual(self.client.post("/events/new", data=self.form).status_code, 400)
        self.assertEqual(self.score().status_code, 400)
        self.assertEqual(ScoreEntry.query.count(), 0)
        response = self.client.post("/events/new", data={**self.form, "csrf_token": "test-csrf-token"})
        self.assertEqual(response.status_code, 302)

    def test_stale_score_rejected_without_overwrite(self):
        first = self.score()
        self.assertEqual(first.status_code, 200, first.get_data(as_text=True))
        self.assertEqual(first.json["version"], 1)
        self.assertEqual(self.score(score=1).status_code, 409)
        self.assertEqual(self.score(score=3, version=1).status_code, 200)
        self.assertEqual(ScoreEntry.query.filter_by(station_no=1, distance_m=6).one().score, 3)
        self.assertEqual(ScoreEntry.query.count(), 20)
        self.assertEqual(ScoreSignature.query.count(), 1)

    def test_concurrent_writers_only_one_succeeds(self):
        barrier = Barrier(2)
        athlete_id = self.athlete_id
        user_id = User.query.filter_by(username="tester").one().id
        db.session.remove()
        def write(value):
            client = app.test_client()
            with client.session_transaction() as session:
                session["_user_id"] = str(user_id)
                session["csrf_token"] = "concurrent-test"
            barrier.wait(timeout=5)
            return client.post(f"/api/scorecard/{athlete_id}/autosave",
                headers={"X-CSRF-Token": "concurrent-test"},
                json=dict(round_no=1, station_no=1, distance_m=6, score=value,
                          played=True, red=False, version=0)).status_code
        with ThreadPoolExecutor(max_workers=2) as executor:
            statuses = list(executor.map(write, [1, 5]))
        self.assertEqual(sorted(statuses), [200, 409])
        self.assertEqual(ScoreEntry.query.count(), 20)
        self.assertEqual(ScoreSignature.query.count(), 1)

    def test_invalid_score_coordinates_leave_no_rows(self):
        for values in [dict(station_no=9), dict(distance_m=100), dict(round_no=99),
                       dict(score="bad"), dict(score=4), dict(red="false")]:
            with self.subTest(values=values):
                self.assertEqual(self.score(**values).status_code, 400)
        self.assertEqual(ScoreEntry.query.count(), 0)
        self.assertEqual(ScoreSignature.query.count(), 0)

    def test_finish_rejects_stale_page(self):
        self.assertEqual(self.score().status_code, 200)
        response = self.client.post(f"/athletes/{self.athlete_id}/scorecard", data={"score_version": "0"})
        self.assertEqual(response.status_code, 409)
        self.assertIsNone(ScoreSignature.query.one().finished_at)

    def test_score_uniqueness_is_enforced_by_database(self):
        self.assertEqual(self.score().status_code, 200)
        db.session.add(ScoreEntry(athlete_id=self.athlete_id, round_no=1, station_no=1, distance_m=6))
        with self.assertRaises(IntegrityError):
            db.session.commit()
        db.session.rollback()

    def test_read_pages_do_not_write_competition_records(self):
        tables = [Event, Athlete, ScoreEntry, ScoreSignature, BracketMatch, ResultsApprovedSetting, User]
        def snapshot():
            return [[tuple(getattr(row, c.name) for c in model.__table__.columns)
                     for row in model.query.order_by(model.id)] for model in tables]
        before = snapshot()
        for path in [f"/events/{self.event_id}/bracket", f"/events/{self.event_id}/bracket_data",
                     f"/events/{self.event_id}/overview?round=2",
                     f"/events/{self.event_id}/overview-data?round=2",
                     f"/events/{self.event_id}/athletes",
                     f"/events/{self.event_id}/results-approved/settings"]:
            self.assertEqual(self.client.get(path).status_code, 200)
        self.assertEqual(snapshot(), before)

    def add_related_records(self):
        court = User(username="event_court", role="court", password_hash="!UNSET!",
                     court_event_id=self.event_id, court_no=2)
        db.session.add(court)
        db.session.flush()
        db.session.add(ScoreEditLog(athlete_id=self.athlete_id, round_no=1, station_no=1,
                                   distance_m=6, edited_by=court.id, editor_username=court.username))
        db.session.add(ResultsApprovedSetting(event_id=self.event_id))
        db.session.commit()
        return court.id

    def test_delete_event_with_audit_and_report_foreign_keys(self):
        self.add_related_records()
        response = self.client.post(f"/events/{self.event_id}/delete")
        self.assertEqual(response.status_code, 302)
        self.assertEqual(Event.query.count(), 0)
        self.assertEqual(ScoreEditLog.query.count(), 0)
        self.assertEqual(ResultsApprovedSetting.query.count(), 0)

    def test_delete_court_preserves_audit_identity(self):
        court_id = self.add_related_records()
        self.assertEqual(self.client.post(f"/events/{self.event_id}/courts/{court_id}/delete").status_code, 302)
        db.session.expire_all()
        log = ScoreEditLog.query.one()
        self.assertIsNone(log.edited_by)
        self.assertEqual(log.editor_username, "event_court")

    def test_fake_image_is_not_saved(self):
        response = self.client.post(f"/events/{self.event_id}/results-approved/settings", data={
            "cover_main_logo": (BytesIO(b"<html>not an image</html>"), "fake.png")})
        self.assertEqual(response.status_code, 302)
        self.assertIsNone(ResultsApprovedSetting.query.one().cover_main_logo_path)

    def test_no_default_accounts_created(self):
        seed_defaults()
        self.assertIsNone(User.query.filter_by(username="superadmin").first())

    def test_logout_requires_post(self):
        self.assertEqual(self.client.get("/logout").status_code, 405)

    def test_legacy_password_requires_change(self):
        user = User.query.filter_by(username="tester").one()
        user.set_password("admin1234")
        db.session.commit()
        response = self.client.post("/login", data={"username": "tester", "password": "admin1234"})
        self.assertTrue(response.location.endswith("/account/password"))
        self.assertTrue(self.client.get("/").location.endswith("/account/password"))


if __name__ == "__main__":
    unittest.main()
