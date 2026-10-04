"""ตั้งรหัสผ่านใหม่ให้บัญชีใดก็ได้ โดยเขียนลงฐานข้อมูลตรง ๆ (ใช้ตอนเข้าระบบไม่ได้เลย)

เครื่องตัวเอง (SQLite ใน instance/shooting.db):
    python tools/set_password.py superadmin

บน Railway (ใช้ฐานข้อมูลเดียวกับเว็บ):
    railway run python tools/set_password.py superadmin
    หรือเปิด Railway > service > Shell แล้วรันคำสั่งเดียวกัน

ระบบจะถามรหัสใหม่ 2 ครั้ง (ไม่แสดงบนจอ) รหัสต้องยาวอย่างน้อย 8 ตัว และห้ามเป็นรหัสตั้งต้นเดิม
"""
import getpass
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
os.environ.setdefault("SECRET_KEY", "set-password-tool")  # ไม่ใช้ session ในสคริปต์นี้

from app import app, db, User, password_problem  # noqa: E402


def main() -> int:
    if len(sys.argv) != 2:
        print("วิธีใช้: python tools/set_password.py <ชื่อผู้ใช้>")
        return 2
    username = sys.argv[1].strip()
    with app.app_context():
        print("ฐานข้อมูล:", app.config["SQLALCHEMY_DATABASE_URI"].split("@")[-1])
        user = User.query.filter_by(username=username).first()
        if not user:
            names = ", ".join(u.username for u in User.query.filter(User.role != "court").order_by(User.id))
            print(f"ไม่พบผู้ใช้ '{username}'  บัญชีที่มี: {names}")
            return 1
        password = os.environ.get("NEW_PASSWORD") or getpass.getpass(f"รหัสใหม่ของ {username}: ")
        if not os.environ.get("NEW_PASSWORD") and getpass.getpass("พิมพ์ซ้ำอีกครั้ง: ") != password:
            print("รหัสสองครั้งไม่ตรงกัน ไม่ได้เปลี่ยนอะไร")
            return 1
        problem = password_problem(password)
        if problem:
            print(problem, "- ไม่ได้เปลี่ยนอะไร")
            return 1
        user.set_password(password)
        db.session.commit()
        print(f"ตั้งรหัสใหม่ให้ {username} ({user.role}) เรียบร้อย เข้าระบบได้เลย")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
