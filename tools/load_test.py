"""ทดสอบโหลดจำลองวันแข่ง: ผู้ชม V คนเปิดหน้า Overview ของอีเวนต์ที่กำลังแข่ง + 8 สนามบันทึกคะแนนทุก 2 วินาที

ใช้กับเซิร์ฟเวอร์ทดสอบ/preview เท่านั้น (สคริปต์นี้เขียนคะแนนจริงลงฐานข้อมูล) ห้ามยิงใส่ระบบวันแข่งจริง
    pip install aiohttp
    LOAD_USER=superadmin LOAD_PASSWORD=... python tools/load_test.py https://preview.example.app new 150 30 1.5 --event 5 --athletes 112-119
ผลที่ดี: save p95 ต่ำกว่า ~300 ms และ err = 0
"""
import asyncio, aiohttp, os, re, sys, time, random
args, opts, it = [], {}, iter(sys.argv[1:])
for a in it:
    if a.startswith('--'): opts[a] = next(it)
    else: args.append(a)
BASE, MODE, VIEWERS, DURATION, POLL = args[0].rstrip('/'), args[1], int(args[2]), float(args[3]), float(args[4])
EVENT = int(opts.get('--event', 5))
_a, _b = (int(x) for x in opts.get('--athletes', '112-119').split('-'))
ATHLETES = list(range(_a, _b + 1))[:8]
lat = {"poll": [], "save": []}; errors = {"poll": 0, "save": 0}; not_modified = [0]; bytes_in = [0]; last_err = [None]

def pct(xs, p):
    xs = sorted(xs); return xs[min(len(xs) - 1, int(len(xs) * p))] * 1000 if xs else float('nan')

async def login(session):
    html = await (await session.get(BASE + '/login')).text()
    m = re.search(r'name="csrf-token" content="([^"]+)"', html)
    data = {'username': os.environ.get('LOAD_USER', 'superadmin'), 'password': os.environ['LOAD_PASSWORD']}
    if m: data['csrf_token'] = m.group(1)
    await session.post(BASE + '/login', data=data)
    html = await (await session.get(BASE + '/')).text()
    m = re.search(r'name="csrf-token" content="([^"]+)"', html)
    return m.group(1) if m else None

async def viewer(stop):
    etag = None
    async with aiohttp.ClientSession() as s:
        await asyncio.sleep(random.random() * POLL)
        while time.time() < stop:
            t = time.time(); h = {'X-Requested-With': 'XMLHttpRequest'}
            if etag and MODE == 'new': h['If-None-Match'] = etag
            try:
                async with s.get(f'{BASE}/events/{EVENT}/overview-data?round=1', headers=h, timeout=aiohttp.ClientTimeout(total=30)) as r:
                    body = await r.read(); bytes_in[0] += int(r.headers.get('Content-Length') or len(body))
                    if r.status == 304: not_modified[0] += 1
                    elif r.status == 200: etag = r.headers.get('ETag')
                    else: errors['poll'] += 1
                lat['poll'].append(time.time() - t)
            except Exception: errors['poll'] += 1
            await asyncio.sleep(max(0, POLL - (time.time() - t)))

async def court(i, stop):
    async with aiohttp.ClientSession(cookie_jar=aiohttp.CookieJar(unsafe=True)) as s:
        token = await login(s); aid = ATHLETES[i]; n = 0
        await asyncio.sleep(i * 0.25)
        while time.time() < stop:
            station = 1 + (n // 4) % 5; dist = 6 + n % 4; n += 1
            body = {'round_no': 1, 'station_no': station, 'distance_m': dist, 'score': random.choice(['5', '3', '0']), 'played': True, 'page_token': f'dev{i}'}
            h = {'Content-Type': 'application/json'}
            if token: h['X-CSRFToken'] = token
            t = time.time()
            try:
                async with s.post(f'{BASE}/api/scorecard/{aid}/autosave', json=body, headers=h, allow_redirects=False, timeout=aiohttp.ClientTimeout(total=30)) as r:
                    txt = await r.text()
                    if r.status != 200 or '"ok":true' not in txt.replace(' ', ''): errors['save'] += 1; last_err[0] = (r.status, txt[:120])
                lat['save'].append(time.time() - t)
            except Exception: errors['save'] += 1
            await asyncio.sleep(max(0, 2.0 - (time.time() - t)))

async def main():
    stop = time.time() + DURATION
    await asyncio.gather(*[viewer(stop) for _ in range(VIEWERS)], *[court(i, stop) for i in range(len(ATHLETES))])
    print(f"{MODE:4} viewers={VIEWERS:3} | save p50 {pct(lat['save'],.5):6.0f} ms p95 {pct(lat['save'],.95):6.0f} ms err {errors['save']:3} "
          f"| poll p50 {pct(lat['poll'],.5):6.0f} ms p95 {pct(lat['poll'],.95):6.0f} ms n={len(lat['poll'])} 304={not_modified[0]} err {errors['poll']} "
          f"| {bytes_in[0]/DURATION/1024:7.0f} KB/s out" + (f" | last save error {last_err[0]}" if last_err[0] else ""))
asyncio.run(main())
