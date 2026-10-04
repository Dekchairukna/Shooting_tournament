"""Disposable browser fixture; never opens the real tournament database."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from test_regressions import EventRegressionTests
from app import app, db, User, Athlete, ensure_bracket, Event

case = EventRegressionTests()
case.setUp()
user = User.query.filter_by(username="tester").one()
user.set_password("Browser-test-password-2026")
athlete = db.session.get(Athlete, case.athlete_id)
athlete.name = '<img src=x onerror="window.injected=true">'
db.session.commit()
ensure_bracket(db.session.get(Event, case.event_id))
case.context.pop()
app.config["TESTING"] = False
app.run(host="127.0.0.1", port=5017, debug=False)
