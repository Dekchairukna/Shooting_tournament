# Round 2 Finalization Only After Round 1 Complete

## พฤติกรรมใหม่
- Ranking/Class รอบ 1 ยังแสดงสดระหว่างการแข่งขันตามเดิม
- ระหว่างที่ผู้เล่นรอบ 1 ยังยิงไม่ครบทั้งอีเวนต์ ระบบจะยังไม่ฟันธง:
  - ผู้ผ่านตรง
  - รายชื่อผู้มีสิทธิ์รอบ 2
  - Shoot-off สำหรับ Top 4
  - Cut line
  - Start list / Overview รอบ 2
- เมื่อรอบ 1 จบครบทุกคนแล้ว ระบบจึงคำนวณ:
  1. ผู้ผ่านตรงตามโควตา
  2. ผู้เล่นถัดมาอย่างน้อย N คนสำหรับรอบ 2
  3. รับเพิ่มทุกคนที่ Total เท่ากับคนลำดับ N ของกลุ่มรอบ 2
  4. Manual Round 2 Override ยังมีผลตามเดิม

## จุดที่แก้
- `is_round_one_complete(event)` เป็นจุดตรวจสถานะกลาง
- `round_two_candidate_ids()` รอรอบ 1 จบก่อนสร้าง candidate
- `get_progression_groups()` ไม่ประกาศ Direct/R2 ระหว่างแข่ง
- `overview_shootoff_ids()` รอบ 1 รอผลครบก่อน
- `build_combined_qualifiers()`, `build_round_two_start_list()` และ `build_round_two_overview_rows()` ไม่สร้างผล/คิวรอบ 2 ก่อนเวลา

## สิ่งที่ไม่เปลี่ยน
- Ranking/Class สดระหว่างรอบ 1
- สี รอตี / กำลังตี / ตีแล้ว
- Court ID / สิทธิ์สนาม
- Audit log / ลายเซ็นการแก้คะแนน
- Round 2 manual override
