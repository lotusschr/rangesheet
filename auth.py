"""
RangeSheet — Authentication ผ่านบัญชีพนักงาน Microsoft 365

วิธีทำงาน:
  1. พนักงานกรอก "รหัสพนักงาน" + "รหัสผ่าน M365 ส่วนตัว"
  2. ระบบลอง login เข้า SharePoint site ของบริษัทจริง
  3. สำเร็จ = พนักงานจริง → เข้าแอพได้
  4. Role:
     - รหัสพนักงานอยู่ใน ADMIN_EMPLOYEE_IDS (2 คน) → ADMIN
     - พนักงานคนอื่นทั้งหมด                        → VIEWER

ข้อดี: ไม่ต้องสร้างรหัสใหม่ ไม่ต้องจำรหัสเพิ่ม ใช้รหัสบริษัทที่มีอยู่แล้ว
       คนนอกบริษัท login ไม่ได้เพราะไม่มีบัญชี M365
"""
import streamlit as st
from datetime import datetime
from config import ADMIN_EMPLOYEE_IDS, STORAGE_MODE

# import SharePoint เฉพาะตอนใช้โหมด sharepoint — โหมด local ไม่ต้องติดตั้ง lib เลย
if STORAGE_MODE == "sharepoint":
    from sharepoint_client import verify_employee_login, SHAREPOINT_AVAILABLE
else:
    verify_employee_login = None
    SHAREPOINT_AVAILABLE = False

# ── โหมด local: ใช้รายชื่อพนักงานในไฟล์นี้ (ไม่ต้องต่อ M365) ────────────────────
# โหมด sharepoint: เช็ครหัสกับ M365 จริง — รายชื่อนี้ไม่ถูกใช้
LOCAL_MODE = (STORAGE_MODE == "local")

# รายชื่อพนักงานสำหรับโหมด local — เพิ่มได้เรื่อยๆ ทีละบรรทัด
# รูปแบบ: "รหัสพนักงาน": ("รหัสผ่าน", "ชื่อที่แสดง")
LOCAL_USERS = {
    "TH000001": ("Range@2026",  "Manager — RangeSheet"),
    "TH000002": ("Merc@2026",   "Manager — Merchandiser"),
    "TH111111": ("Lotus@2026",  "Range Team Member"),
    # เพิ่มพนักงานคนอื่น:
    # "TH111112": ("รหัสของเขา", "Somchai W."),
}

# ── Permission matrix ─────────────────────────────────────────────────────────
PERMISSIONS = {
    "admin": {
        "view_data": True, "edit_display_columns": True, "create_canvas": True,
        "export": True,
        "upload_files": True, "delete_files": True,
        "edit_raw_data": True, "sharepoint_sync": True, "submit_range": True,
    },
    "viewer": {
        "view_data": True, "edit_display_columns": True, "create_canvas": True,
        "export": True,
        "upload_files": False, "delete_files": False,
        "edit_raw_data": False, "sharepoint_sync": False, "submit_range": False,
    },
}


def _role_for(employee_id: str) -> str:
    """รหัสพนักงานอยู่ใน admin list → admin, นอกนั้น viewer"""
    return "admin" if employee_id.strip().upper() in [a.upper() for a in ADMIN_EMPLOYEE_IDS] else "viewer"


def login(employee_id: str, password: str) -> tuple[bool, str]:
    """
    Login ด้วยบัญชีพนักงาน M365
    Returns: (success, error_message)
    """
    emp = employee_id.strip().upper()

    # ── โหมด LOCAL: เช็ครายชื่อในไฟล์นี้ ──
    if LOCAL_MODE:
        rec = LOCAL_USERS.get(emp)
        if rec and rec[0] == password:
            st.session_state.auth_user = {
                "employee_id": emp,
                "name": rec[1],
                "email": f"{emp.lower()}@local",
                "role": _role_for(emp),
                "login_time": datetime.now().strftime("%H:%M"),
            }
            st.session_state.sp_client = None
            return True, ""
        return False, "รหัสพนักงานหรือรหัสผ่านไม่ถูกต้อง"

    # ── โหมด SHAREPOINT: เช็คกับ M365 จริง ──
    if not SHAREPOINT_AVAILABLE:
        return False, "Server ยังไม่ได้ติดตั้ง Office365-REST-Python-Client"

    result = verify_employee_login(emp, password)
    if result["success"]:
        st.session_state.auth_user = {
            "employee_id": emp,
            "name": result["display_name"],     # ชื่อจริงจาก M365
            "email": result["email"],
            "role": _role_for(emp),
            "login_time": datetime.now().strftime("%H:%M"),
        }
        # เก็บ SharePoint client ไว้ใช้ upload/download ต่อด้วยสิทธิ์คนนี้
        st.session_state.sp_client = result["client"]
        return True, ""
    return False, result["error"] or "เข้าสู่ระบบไม่สำเร็จ"


def logout():
    st.session_state.auth_user = None
    st.session_state.sp_client = None


def current_user() -> dict | None:
    return st.session_state.get("auth_user")


def get_sp_client():
    """คืน SharePoint client ของ user ที่ login อยู่ (None ถ้า dev mode)"""
    return st.session_state.get("sp_client")


def has_permission(action: str) -> bool:
    user = current_user()
    if user is None:
        return False
    return PERMISSIONS.get(user["role"], {}).get(action, False)


def is_admin() -> bool:
    user = current_user()
    return user is not None and user["role"] == "admin"


def require_login() -> bool:
    """แสดงหน้า login ถ้ายังไม่ login — คืน True เมื่อ login แล้ว"""
    if current_user() is not None:
        return True

    st.markdown("""
    <div style="max-width:380px;margin:60px auto 0;text-align:center;">
      <div style="width:56px;height:56px;background:#2BBFA4;border-radius:14px;
           display:inline-flex;align-items:center;justify-content:center;font-size:28px;">🗂️</div>
      <div style="font-size:22px;font-weight:800;color:#1A1A1A;margin-top:12px;">RangeSheet</div>
      <div style="font-size:13px;color:#888;">เข้าสู่ระบบด้วยบัญชีพนักงาน Microsoft 365</div>
    </div>""", unsafe_allow_html=True)

    _, mid, _ = st.columns([1, 1.2, 1])
    with mid:
        with st.form("login_form"):
            emp_id = st.text_input("รหัสพนักงาน", placeholder="เช่น TH111111")
            password = st.text_input("รหัสผ่าน (M365 ส่วนตัว)", type="password")
            submitted = st.form_submit_button("เข้าสู่ระบบ", use_container_width=True)
            if submitted:
                if not emp_id or not password:
                    st.error("กรอกรหัสพนักงานและรหัสผ่านให้ครบ")
                else:
                    with st.spinner("กำลังตรวจสอบกับ Microsoft 365..."):
                        ok, err = login(emp_id, password)
                    if ok:
                        st.rerun()
                    else:
                        st.error(err)
        if LOCAL_MODE:
            st.caption("โหมด Local — ใช้รหัสที่ Manager กำหนดให้ · เมื่อย้ายขึ้น SharePoint จะเปลี่ยนเป็นรหัส M365 ของบริษัท")
        else:
            st.caption("ใช้รหัสผ่านเดียวกับที่ login Outlook / Teams ของบริษัท · ระบบไม่เก็บรหัสผ่านของคุณ")
    return False
