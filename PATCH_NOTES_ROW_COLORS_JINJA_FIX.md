# Fix: Overview row colors inside Jinja content block

แก้ปัญหา CSS สีแถวถูกวางหลัง `{% endblock %}` ทำให้ template inheritance ไม่ render CSS ชุดนั้นในหน้าเว็บจริง

กฎสีแถวหลังแก้:
- รอตี / กำลังตี: ใช้พื้นตารางเดิม
- ตีครบ/ส่งคะแนนแล้ว แต่ยังไม่ APPROVED: สีเขียวทั้งแถว
- APPROVED: สีฟ้าทั้งแถว
- cursor/ช่องคะแนนกำลังตี: ใช้กฎเดิม ไม่เปลี่ยน
- Qualification Cut Line: ใช้กฎเดิม ไม่เปลี่ยน

แก้โดยย้าย CSS ของ `visual-finished` และ `visual-approved` เข้า `<style>` หลักภายใน `{% block content %}` และลบ style orphan ที่อยู่นอก block
