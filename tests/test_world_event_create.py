import os
import tempfile
import unittest
from datetime import date

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-world-event-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from app import Athlete, Event, ScoreEntry, User, WORLD_2026_EVENT_NAME, WORLD_2026_EVENT_NAMES, app, db, nation_flag_file


class WorldEventCreateTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()

    def tearDown(self):
        db.session.remove(); self.ctx.pop()

    def test_create_48_country_event(self):
        command = ["seed-world-2026", "--competition-date", "2026-11-21", "--lanes", "8"]
        runner = app.test_cli_runner()
        result = runner.invoke(args=command)
        self.assertEqual(result.exit_code, 0, result.output)
        event = Event.query.filter_by(name=WORLD_2026_EVENT_NAME).one()
        self.assertEqual(event.competition_date, date(2026, 11, 21))
        self.assertFalse(event.has_round_two)
        athletes = Athlete.query.filter_by(event_id=event.id).order_by(Athlete.start_order).all()
        self.assertEqual(len(athletes), 48)
        self.assertEqual((athletes[0].affiliation, athletes[-1].affiliation), ("Algeria (ALG)", "United States (USA)"))
        self.assertEqual(athletes[0].name, "Algeria")
        self.assertTrue(all(nation_flag_file(a.affiliation) for a in athletes))
        self.assertEqual(sorted(a.start_order for a in athletes), list(range(1, 49)))
        self.assertEqual(User.query.filter_by(court_event_id=event.id, role="court").count(), 8)
        self.assertEqual(ScoreEntry.query.count(), 0)

        athletes[0].name = "Actual Athlete"
        event.name = WORLD_2026_EVENT_NAMES[1]
        db.session.commit()
        retry = runner.invoke(args=command)
        self.assertEqual(retry.exit_code, 0, retry.output)
        self.assertIn("already exists", retry.output)
        self.assertEqual(Event.query.count(), 1)
        self.assertEqual(Athlete.query.filter_by(event_id=event.id).count(), 48)
        self.assertEqual(athletes[0].name, "Actual Athlete")


if __name__ == "__main__":
    unittest.main()
