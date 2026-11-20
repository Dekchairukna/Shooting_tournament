"""รอบ 2: แถวอยู่ตามคิวตีรอบ 2 · สลับตามอันดับเฉพาะคนที่ตีจบ · ผู้ผ่านตรงอยู่บนสุดเมื่อปิดเรียง.
Run: python -m unittest tests.test_round2_view_order -v"""
import os, tempfile, unittest
from datetime import datetime

_test_dir = tempfile.TemporaryDirectory(prefix="shooting-r2view-tests-")
os.environ.setdefault("DATABASE_URL", "sqlite:///" + _test_dir.name + "/test.db")
os.environ.setdefault("SECRET_KEY", "test-only-secret")

from flask import g
from app import db, ScoreEntry, ScoreSignature, invalidate_poll_cache
from test_round2_mode import Round2ModeTests


class Round2ViewOrderTests(Round2ModeTests):
    def _r2(self, e):
        g.pop("shooting_request_cache", None)
        invalidate_poll_cache(e.id)
        data = self.client.get(f"/events/{e.id}/overview-data?round=2").get_json()
        return [x["name"] for x in sorted(data, key=lambda x: x["view_order"])]

    def _shoot(self, e, name, entries, finish=False):
        a = next(x for x in e.athletes if x.name == name)
        self.client.get(f"/events/{e.id}/overview?round=2")  # sync รายชื่อรอบ 2
        for en in ScoreEntry.query.filter_by(athlete_id=a.id, round_no=2).order_by(ScoreEntry.id).limit(entries):
            en.score, en.is_scored = 5, True
        sig = ScoreSignature.query.filter_by(athlete_id=a.id, round_no=2).first()
        sig.started_at = datetime.utcnow()
        a.status = "active"
        if finish:
            sig.finished_at = datetime.utcnow(); a.status = "finished"
        db.session.commit()

    def setUp(self):
        super().setUp()
        # ผ่านตรง 4 (A1–A4) · รอบ 2 = Class 5–10 (A5–A10) · คิวตีรอบ 2: คะแนนน้อยสุดก่อน = A10 … A5
        self.e = self._event("cutoff", 10, athletes=12)
        self.queue = ["A10", "A9", "A8", "A7", "A6", "A5"]

    def test_opens_in_round2_shooting_order(self):
        self.assertEqual(self._r2(self.e), ["A1", "A2", "A3", "A4"] + self.queue)

    def test_shooting_player_stays_in_queue(self):
        self._shoot(self.e, "A7", 10)  # กำลังตี ได้ 5 ไปแล้ว 10 ลูก แต่ยังไม่จบ
        self.assertEqual(self._r2(self.e), ["A1", "A2", "A3", "A4"] + self.queue)

    def test_finished_player_moves_up_by_rank(self):
        self._shoot(self.e, "A7", 20, finish=True)
        self.assertEqual(self._r2(self.e), ["A1", "A2", "A3", "A4", "A7", "A10", "A9", "A8", "A6", "A5"])

    def test_sort_off_keeps_queue_and_direct_on_top(self):
        self._shoot(self.e, "A7", 20, finish=True)
        self.client.post(f"/events/{self.e.id}/overview-sort", data={"round": "2"})
        self.assertEqual(self._r2(self.e), ["A1", "A2", "A3", "A4"] + self.queue)


if __name__ == "__main__":
    unittest.main()
