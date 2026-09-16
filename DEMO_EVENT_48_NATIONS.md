# Demo Event: 48 Nations — Petanque Shooting 2026

มี Event ตัวอย่างชื่อ **DEMO • 48 NATIONS • PETANQUE SHOOTING 2026** อยู่ใน `instance/shooting.db` แล้ว

## โครงการแข่งขัน
- นักกีฬา/ประเทศ: 48 ประเทศ
- สนาม: 8 สนาม (สนามละ 6 ประเทศ)
- Qualification Round 1: ครบ 48 ประเทศ
- ผ่านตรง: Class 1–4 → Quarter Final
- Round 2: Class 5–16 → 12 ประเทศ
- ผ่านจาก Round 2: 4 ประเทศ
- Knockout: Quarter Final 8 → Semi Final 4 → Final 2
- Champion: BELGIUM
- Runner-up: SPAIN

## ผู้ผ่านเข้า Quarter Final
Seed 1 BELGIUM
Seed 2 SPAIN
Seed 3 ITALY
Seed 4 MADAGASCAR
Seed 5 SWITZERLAND
Seed 6 MALAYSIA
Seed 7 JAPAN
Seed 8 GERMANY

## ผล Knockout จำลอง
Quarter Final
- BELGIUM ชนะ GERMANY
- MADAGASCAR ชนะ SWITZERLAND
- ITALY ชนะ MALAYSIA
- SPAIN ชนะ JAPAN

Semi Final
- BELGIUM ชนะ MADAGASCAR
- SPAIN ชนะ ITALY

Final
- BELGIUM ชนะ SPAIN

> เวอร์ชันประชุมใช้ทีมกลาง ๆ ใน Knockout/Final เพื่อให้เห็นการทำงานของระบบโดยไม่ให้ THAILAND หรือ FRANCE เป็นคู่เด่น

## ใช้สำหรับพรีเซ็น
เปิด Event นี้แล้วสามารถกดดูได้ตามลำดับ:
1. ROUND 1 — คะแนนครบ 48 ประเทศ / Ranking / Cut line
2. ROUND 2 — 12 ประเทศจาก Class 5–16 / รวมคะแนน R1+R2
3. Quarter Final — 8 ประเทศ
4. Semi Final — 4 ประเทศ
5. Final — BELGIUM vs SPAIN
6. Results Approved — ผลสรุปอย่างเป็นทางการ

## สร้างข้อมูล Demo ใหม่
หากต้องการรีเซ็ต Demo ให้รัน:

```bash
python3 seed_demo_48_countries_sqlite.py
```

สคริปต์จะลบเฉพาะ Event Demo ชื่อนี้แล้วสร้างใหม่ ไม่แตะ Event อื่น
