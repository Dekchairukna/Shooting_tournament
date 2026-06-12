# Live Report Board API ที่เพิ่มให้ Shooting

เพิ่ม endpoint public สำหรับเว็บรายงานผล:

- `/api/public/shooting/events` รายการอีเวนต์
- `/api/public/shooting/event/<event_id>/report` อันดับ/คะแนนนักกีฬา
- `/public/shooting/<event_id>/live` หน้า public สำหรับ iframe

เพิ่ม CORS และอนุญาต iframe แล้วใน `after_request`
