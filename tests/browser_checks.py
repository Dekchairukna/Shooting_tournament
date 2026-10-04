"""Run against tests/browser_server.py with a fresh, disposable Chrome profile."""
from pathlib import Path
from playwright.sync_api import sync_playwright, expect

BASE = "http://127.0.0.1:5017"

with sync_playwright() as p:
    browser = p.chromium.launch(
        executable_path="/Applications/Google Chrome.app/Contents/MacOS/Google Chrome", headless=True)
    context = browser.new_context(viewport={"width": 1440, "height": 1000})
    page = context.new_page()
    errors = []
    page.on("pageerror", lambda error: errors.append(str(error)))
    page.goto(BASE + "/login")
    page.locator('[name="username"]').fill("tester")
    page.locator('[name="password"]').fill("Browser-test-password-2026")
    page.locator('button[type="submit"]').click()
    expect(page).to_have_url(BASE + "/")
    page.goto(BASE + "/events/1/overview")
    page.get_by_role("button", name="📊 ดูสถิติ 5/3").click()
    expect(page.locator('#statsModalBody .name')).to_contain_text('<img src=x')
    assert page.locator('#statsModalBody img').count() == 0
    assert page.evaluate("window.injected === undefined")
    print("PASS: login with CSRF; untrusted athlete name remains plain text")

    page.goto(BASE + "/athletes/1/scorecard")
    other = context.new_page()
    other.on("pageerror", lambda error: errors.append(str(error)))
    other.goto(BASE + "/athletes/1/scorecard")
    cell = '.auto-score[data-round="1"][data-station="1"][data-distance="6"]'
    with page.expect_response('**/api/scorecard/1/autosave') as saved:
        page.locator(cell).select_option("5")
    assert saved.value.status == 200, saved.value.text()
    expect(page.locator('#autosaveStatus')).to_have_text("บันทึกแล้ว")
    with other.expect_response('**/api/scorecard/1/autosave') as stale:
        other.locator(cell).select_option("1")
    assert stale.value.status == 409
    expect(other.locator('#autosaveStatus')).to_contain_text("อีกหน้าจอ")
    assert other.locator('[data-save-error="true"]').count() == 1
    other.get_by_role('button', name='จบการตี', exact=True).click()
    expect(other).to_have_url(BASE + '/athletes/1/scorecard')
    print("PASS: autosave succeeds; stale second screen rejected; finish blocked")
    other.on("dialog", lambda dialog: dialog.accept())
    other.close()

    page.goto(BASE + '/events/1/bracket')
    expect(page.locator('#liveConnectionStatus')).to_contain_text('เชื่อมต่อแล้ว')
    page.route('**/bracket_data*', lambda route: route.abort())
    expect(page.locator('#liveConnectionStatus')).to_contain_text('เชื่อมต่อไม่ได้', timeout=15000)
    page.unroute('**/bracket_data*')
    expect(page.locator('#liveConnectionStatus')).to_contain_text('เชื่อมต่อแล้ว', timeout=20000)
    print('PASS: live scoreboard indicates disconnection and recovers')

    page.set_viewport_size({'width': 390, 'height': 844})
    page.goto(BASE + '/athletes/1/scorecard')
    expect(page.locator('#autosaveStatus')).to_be_visible()
    page.screenshot(path='/tmp/shooting-scorecard-mobile.png', full_page=True)
    page.goto(BASE + '/events/1/athletes')
    assert page.locator('form[method="post"]:not(:has(input[name="csrf_token"]))').count() == 0
    page.goto(BASE + '/events/1/edit')
    page.locator('[name="name"]').fill('Browser reviewed event')
    page.get_by_role('button', name='บันทึกการแก้ไข').click()
    expect(page).to_have_url(BASE + '/events/1/overview?round=1')
    page.get_by_role('button', name='Toggle navigation').click()
    page.locator('.logout-link').click()
    expect(page).to_have_url(BASE + '/login')
    assert not errors, errors
    print('PASS: mobile scorecard, protected forms, edit event, POST logout; no JavaScript errors')
    browser.close()
