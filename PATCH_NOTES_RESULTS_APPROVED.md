# PATCH: Results Approved Report

เพิ่มรายงาน Results Approved ตามโครงเอกสารตัวอย่าง `Results_Approved.docx`

## เพิ่มในระบบ

- ปุ่ม `Results Approved` ในหน้า Overview ของแต่ละ Event
- หน้า preview/print: `/events/<event_id>/results-approved`
- ดาวน์โหลด DOCX: `/events/<event_id>/results-approved.docx`

## รายงานที่สร้างให้

1. Cover: PÉTANQUE RESULTS APPROVED
2. NAME LISTS
3. QUALIFICATION ROUND 1
4. Qualification Shooting รายสถานีรอบ 1
5. QUALIFICATION ROUND 2 ถ้ามีรอบ 2
6. Qualification Shooting Round 2 รายสถานี ถ้ามีรอบ 2
7. KNOCKOUT RESULTS จาก BracketMatch
8. RANKING RESULT เหรียญ Gold / Silver / Bronze / Bronze

## กติกาข้อมูลที่ใช้

- รอบ 1 ใช้อันดับจาก `build_round_ranking(event, 1)`
- รอบ 2 ใช้คิวจาก `build_round_two_overview_rows(event)` เรียงตามคิวตีรอบ 2
- RANK(QF2) ใช้ ranking ปัจจุบันใน Overview รอบ 2
- Knockout ดึงคะแนนตาม round_no ของ bracket จริง
- Medal result ใช้ winner/loser จาก Final และ Semifinal ถ้ามีผลแล้ว
- ถ้ายังไม่มีผล bracket จะ fallback เป็น seed/อันดับรวม เพื่อให้ทำฉบับร่างได้

## Dependency ใหม่

เพิ่ม `python-docx` ใน `requirements.txt`

หลังแตก zip ให้ติดตั้ง dependency ใหม่ด้วย:

```bash
source venv/bin/activate
python -m pip install -r requirements.txt
python app.py
```
