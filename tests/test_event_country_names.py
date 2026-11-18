"""Regression tests for updating an existing event's athlete names by country."""

import os
import re
import unittest
from datetime import date
from io import BytesIO
from unittest.mock import patch

from openpyxl import Workbook, load_workbook
from sqlalchemy import create_engine


class EventCountryNamesTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # Importing the application initializes its configured database. Never
        # let a standalone run initialize the regular local or hosted database.
        with patch.dict(os.environ, {
            "DATABASE_URL": "sqlite://",
            "SECRET_KEY": "country-name-import-test-secret",
            "SUPERADMIN_PASSWORD": "country-name-import-test-password",
        }):
            import app as application

        cls.application = application
        cls.app = application.app
        cls.db = application.db

    def setUp(self):
        application = self.application
        db = self.db
        engine = create_engine("sqlite://")
        self.addCleanup(engine.dispose)
        self.enterContext(patch.dict(self.app.config, {"TESTING": True}))
        with self.app.app_context():
            self.enterContext(patch.dict(db.engines, {None: engine}))
            db.create_all()
            admin = application.User(username="admin", role="admin", password_hash="!UNSET!")
            supervisor = application.User(username="boss", role="superadmin", password_hash="!UNSET!")
            viewer = application.User(username="viewer", role="user", password_hash="!UNSET!")
            event = self.make_event("World championship")
            other = self.make_event("Other event")
            db.session.add_all([admin, supervisor, viewer, event, other])
            db.session.flush()
            singapore = self.make_athlete(event.id, "Singapore", "Singapore (SGP)", 1, 1)
            france = self.make_athlete(event.id, "Existing France player", "France (FRA)", 2, 2)
            other_singapore = self.make_athlete(other.id, "Other event player", "Singapore (SGP)", 1, 1)
            db.session.add_all([singapore, france, other_singapore])
            db.session.flush()
            score = application.ScoreEntry(
                athlete_id=singapore.id, round_no=1, station_no=1, distance_m=6,
                score=5, is_scored=True,
            )
            db.session.add(score)
            db.session.commit()
            self.event_id = event.id
            self.singapore_id = singapore.id
            self.france_id = france.id
            self.other_singapore_id = other_singapore.id
            self.score_id = score.id
            self.admin_id = admin.id
            self.supervisor_id = supervisor.id
            self.viewer_id = viewer.id

        self.athletes_url = f"/events/{self.event_id}/athletes"
        self.import_url = f"{self.athletes_url}/names-by-country"
        self.template_url = f"{self.import_url}.xlsx"

    def make_event(self, name):
        return self.application.Event(
            name=name, event_group="World", category="ชาย", competition_date=date(2026, 11, 20),
            location="Khon Kaen", lane_count=2,
        )

    def make_athlete(self, event_id, name, country, order, lane):
        return self.application.Athlete(
            event_id=event_id, bib_no=str(order), name=name, affiliation=country,
            start_order=order, lane_no=lane, lane_order=1,
        )

    def signed_in_client(self, user_id):
        client = self.app.test_client()
        with client.session_transaction() as session:
            session["_user_id"] = str(user_id)
            session["_fresh"] = True
        page = client.get("/login").get_data(as_text=True)
        client.environ_base["HTTP_X_CSRF_TOKEN"] = re.search(
            r'name="csrf-token" content="([^"]+)"', page,
        ).group(1)
        return client

    def upload(self, client, rows):
        workbook = Workbook()
        for row in rows:
            workbook.active.append(row)
        content = BytesIO()
        workbook.save(content)
        workbook.close()
        content.seek(0)
        return client.post(self.import_url, data={"excel_file": (content, "roster.xlsx")})

    def assert_flash(self, client, category, text):
        with client.session_transaction() as session:
            self.assertTrue(any(
                level == category and text in message for level, message in session.get("_flashes", [])
            ))

    def athlete_names(self):
        with self.app.app_context():
            return [self.db.session.get(self.application.Athlete, athlete_id).name for athlete_id in (
                self.singapore_id, self.france_id, self.other_singapore_id,
            )]

    def test_file_matches_country_instead_of_row_position_and_only_updates_name(self):
        client = self.signed_in_client(self.admin_id)
        response = self.upload(client, [
            ("ประเทศ", "ชื่อนักกีฬา"), ("ฝรั่งเศส", ""), ("SGP", "  ทดสอบ นักกีฬา  "),
        ])
        self.assertEqual(response.status_code, 302)
        self.assertEqual(response.location, self.athletes_url)
        self.assert_flash(client, "success", "อัปเดตชื่อนักกีฬา 1 คน")
        self.assertEqual(self.athlete_names(), ["ทดสอบ นักกีฬา", "Existing France player", "Other event player"])
        with self.app.app_context():
            athlete = self.db.session.get(self.application.Athlete, self.singapore_id)
            score = self.db.session.get(self.application.ScoreEntry, self.score_id)
            self.assertEqual((athlete.affiliation, athlete.bib_no, athlete.start_order), ("Singapore (SGP)", "1", 1))
            self.assertEqual((athlete.lane_no, athlete.lane_order), (1, 1))
            self.assertEqual((score.athlete_id, score.score, score.is_scored), (athlete.id, 5, True))

    def test_invalid_file_never_updates_even_rows_already_matched(self):
        client = self.signed_in_client(self.admin_id)
        invalid_rows = [
            (("ZZZ", "Unknown country"), "ไม่พบประเทศ ZZZ"),
            (("Singapore (SGP)", "Duplicate country"), "ซ้ำในไฟล์"),
        ]
        for invalid_row, error in invalid_rows:
            with self.subTest(invalid_row=invalid_row):
                response = self.upload(client, [("SGP", "Should not save"), invalid_row])
                self.assertEqual(response.status_code, 302)
                self.assert_flash(client, "danger", error)
                self.assertEqual(self.athlete_names(), ["Singapore", "Existing France player", "Other event player"])

    def test_ambiguous_country_in_same_event_does_not_update_either_athlete(self):
        with self.app.app_context():
            duplicate = self.make_athlete(self.event_id, "Second Singapore", "Singapore", 3, 1)
            self.db.session.add(duplicate)
            self.db.session.commit()
            duplicate_id = duplicate.id

        client = self.signed_in_client(self.admin_id)
        response = self.upload(client, [("SGP", "Should not save")])
        self.assertEqual(response.status_code, 302)
        self.assert_flash(client, "danger", "มากกว่าหนึ่งทีม")
        self.assertEqual(self.athlete_names()[0], "Singapore")
        with self.app.app_context():
            duplicate = self.db.session.get(self.application.Athlete, duplicate_id)
            self.assertEqual(duplicate.name, "Second Singapore")

    def test_download_contains_event_countries_and_empty_name_cells(self):
        client = self.signed_in_client(self.supervisor_id)
        page = client.get(self.athletes_url).get_data(as_text=True)
        self.assertIn(self.template_url, page)
        self.assertIn(self.import_url, page)
        response = client.get(self.template_url)
        self.assertEqual(response.status_code, 200)
        workbook = load_workbook(BytesIO(response.data), read_only=True)
        self.addCleanup(workbook.close)
        self.assertEqual(list(workbook.active.values), [
            ("ประเทศ", "ชื่อนักกีฬา"), ("Singapore (SGP)", None), ("France (FRA)", None),
        ])

    def test_regular_user_cannot_download_or_update_roster(self):
        client = self.signed_in_client(self.viewer_id)
        self.assertEqual(client.get(self.template_url).status_code, 302)
        self.assertEqual(self.upload(client, [("SGP", "Should not save")]).status_code, 302)
        self.assertEqual(self.athlete_names()[0], "Singapore")
