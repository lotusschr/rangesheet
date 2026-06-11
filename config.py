# RangeSheet - Configuration

# เลือกที่เก็บไฟล์: "local" = ลงเครื่อง / "sharepoint" = M365 (เมื่อ IT พร้อม)

STORAGE_MODE = "local"

# SharePoint (เตรียมไว้ ยังไม่ใช้ในโหมด local)

SHAREPOINT_CONFIG = {

    "site_url": "https://thlotuss.sharepoint.com/teams/RangeSheet261",

    "doc_library": "Shared Documents/RangeSheet",

    "email_domain": "thlotuss.com",

}

# รหัสพนักงาน 2 คนที่เป็น ADMIN (แก้เป็นรหัสจริง)

ADMIN_EMPLOYEE_IDS = [

    "",

    "TH000002",

]

# App settings

APP_CONFIG = {

    "allowed_extensions": ["csv", "xlsx", "xls"],

    "max_file_mb": 1024,

    "merge_similarity_threshold": 0.5,

}
 