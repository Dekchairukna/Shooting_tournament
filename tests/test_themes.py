"""Site-wide themes. Run: python -m unittest tests.test_themes -v"""
import os, re, tempfile, unittest
from io import BytesIO

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-theme-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from app import app, db, User, SiteTheme, SiteThemeAsset, ensure_default_themes

PNG = (b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01\x08\x06\x00\x00\x00\x1f\x15\xc4\x89"
       b"\x00\x00\x00\rIDATx\x9cc\xf8\x0f\x00\x00\x01\x01\x00\x05\x18\xd8N\x00\x00\x00\x00IEND\xaeB`\x82")


class ThemeTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.ctx = app.app_context(); self.ctx.push()
        with db.engine.connect() as conn:
            conn.exec_driver_sql("PRAGMA foreign_keys=OFF")
            db.metadata.drop_all(bind=conn); db.metadata.create_all(bind=conn); conn.commit()
        user = User(username="boss", role="superadmin", password_hash="!UNSET!")
        db.session.add(user); ensure_default_themes(); db.session.commit()
        self.client = app.test_client()
        with self.client.session_transaction() as s:
            s["_user_id"] = str(user.id); s["_fresh"] = True
        page = self.client.get("/admin/themes").get_data(as_text=True)
        self.client.environ_base["HTTP_X_CSRF_TOKEN"] = re.search(r'name="csrf-token" content="([^"]+)"', page).group(1)

    def tearDown(self):
        db.session.remove(); self.ctx.pop()

    def test_default_kku_theme_renders_like_before(self):
        self.assertEqual(SiteTheme.query.count(), 3)
        html = self.client.get("/").get_data(as_text=True)
        self.assertIn("--th-primary:#ef4b12", html)
        self.assertIn("52nd PÉTANQUE WORLD CHAMPIONSHIP 2026", html)
        self.assertIn("kku2026_event_logo.jpg", html)

    def test_switch_theme_sitewide(self):
        plain = SiteTheme.query.filter_by(name="มาตรฐาน (ทั่วไป)").first()
        self.assertEqual(self.client.post(f"/admin/themes/{plain.id}/activate").status_code, 302)
        html = self.client.get("/").get_data(as_text=True)
        self.assertIn("--th-primary:#2563eb", html)
        self.assertNotIn("kku2026_event_logo.jpg", html)
        self.assertEqual(SiteTheme.query.filter_by(is_active=True).count(), 1)

    def test_create_theme_with_image_and_activate(self):
        r = self.client.post("/admin/themes/new", data={
            "name": "อปท.41", "short_name": "อปท 41", "title": "กีฬาอปท. ครั้งที่ 41", "subtitle": "สกลนคร",
            "color_primary": "#0e7a3a", "color_primary_dark": "#0a5a2b", "color_soft": "#e8f5ec",
            "color_cream": "#f6fbf7", "color_accent": "#f2b705", "color_ink": "#0f1f15", "color_line": "#b9dcc5",
            "show_hero": "yes", "activate": "yes",
            "image_logo": (BytesIO(PNG), "logo.png"),
            "image_poster": (BytesIO(b"<svg onload=alert(1)>"), "evil.png"),
        }, content_type="multipart/form-data")
        self.assertEqual(r.status_code, 302)
        t = SiteTheme.query.filter_by(name="อปท.41").one()
        self.assertTrue(t.is_active)
        self.assertEqual({a.kind for a in t.assets}, {"logo"})  # fake image rejected
        html = self.client.get("/").get_data(as_text=True)
        self.assertIn("กีฬาอปท. ครั้งที่ 41", html)
        self.assertIn("--th-primary:#0e7a3a", html)
        logo = re.search(r'src="(/theme-asset/\d+/logo\?v=\d+)"', html).group(1)
        img = self.client.get(logo)
        self.assertEqual(img.status_code, 200)
        self.assertEqual(img.mimetype, "image/png")
        self.assertEqual(img.data, PNG)

    def test_bad_color_and_hero_off(self):
        t = SiteTheme.query.filter_by(is_active=True).first()
        self.client.post(f"/admin/themes/{t.id}/edit", data={"name": t.name, "color_primary": "red;}</style><script>",
                                                             "title": "X"})
        db.session.refresh(t)
        self.assertEqual(t.color_primary, "#ef4b12")
        self.assertFalse(t.show_hero)
        html = self.client.get("/").get_data(as_text=True)
        self.assertNotIn('<header class="championship-hero', html)

    def test_cannot_delete_active_and_duplicate(self):
        active = SiteTheme.query.filter_by(is_active=True).first()
        self.client.post(f"/admin/themes/{active.id}/delete")
        self.assertIsNotNone(db.session.get(SiteTheme, active.id))
        self.client.post(f"/admin/themes/{active.id}/duplicate")
        self.assertEqual(SiteTheme.query.count(), 4)
        builtin = SiteTheme.query.filter_by(is_active=False, is_builtin=True).first()
        self.client.post(f"/admin/themes/{builtin.id}/delete")
        self.assertIsNotNone(db.session.get(SiteTheme, builtin.id))
        copy = SiteTheme.query.filter_by(is_builtin=False).one()
        self.client.post(f"/admin/themes/{copy.id}/delete")
        self.assertIsNone(db.session.get(SiteTheme, copy.id))

    def test_phuphan_builtin(self):
        t = SiteTheme.query.filter_by(short_name="ภูพานเกมส์").one()
        self.client.post(f"/admin/themes/{t.id}/activate")
        html = self.client.get("/").get_data(as_text=True)
        self.assertIn("themes/phuphan_logo.jpg", html)
        self.assertIn("themes/phuphan_poster.jpg", html)
        self.assertIn("--th-primary:#aa241a", html)
        # running seeding again never duplicates
        ensure_default_themes(); db.session.commit()
        self.assertEqual(SiteTheme.query.filter_by(short_name="ภูพานเกมส์").count(), 1)

    def test_pages_render(self):
        t = SiteTheme.query.first()
        for url in ["/admin/themes", "/admin/themes/new", f"/admin/themes/{t.id}/edit", "/events/draw", "/events/bulk-new"]:
            self.assertEqual(self.client.get(url).status_code, 200, url)


if __name__ == "__main__":
    unittest.main()
