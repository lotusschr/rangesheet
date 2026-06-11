"""
RangeSheet — Local Storage (ใช้แทน SharePoint ชั่วคราว)
บันทึกไฟล์ลงเครื่อง/server ที่รันแอพ ระหว่างรอ IT อนุมัติ SharePoint

โครงสร้างโฟลเดอร์ที่สร้างอัตโนมัติ:
  rangesheet_data/
  ├── uploads/      ← ไฟล์ดิบที่อัปโหลด (.csv .xlsx)
  ├── merged/       ← ตารางรวมหลัง auto-merge (เก็บเป็น snapshot)
  └── audit/        ← audit log (.csv)

ภายหลังเมื่อ SharePoint พร้อม: สลับกลับโดยแก้ STORAGE_MODE ใน config.py
โค้ดส่วนอื่นไม่ต้องแก้เลย เพราะ function ชื่อเดียวกัน
"""
import os
import io
import shutil
import pandas as pd
from datetime import datetime

# ── โฟลเดอร์หลัก ───────────────────────────────────────────────────────────────
BASE_DIR    = "rangesheet_data"
UPLOAD_DIR  = os.path.join(BASE_DIR, "uploads")
MERGED_DIR  = os.path.join(BASE_DIR, "merged")
AUDIT_DIR   = os.path.join(BASE_DIR, "audit")

def _ensure_dirs():
    """สร้างโฟลเดอร์ถ้ายังไม่มี — เรียกอัตโนมัติทุกครั้งก่อนบันทึก"""
    for d in (UPLOAD_DIR, MERGED_DIR, AUDIT_DIR):
        os.makedirs(d, exist_ok=True)


# ══════════════════════════════════════════════════════════════════════════════
# ฟังก์ชันหลัก — ชื่อ/รูปแบบเหมือน SharePointClient เป๊ะ เพื่อให้สลับกลับได้ง่าย
# ══════════════════════════════════════════════════════════════════════════════

def upload_file(file_bytes: bytes, filename: str, subfolder: str = "") -> dict:
    """
    บันทึกไฟล์ลงเครื่อง (แทน upload ขึ้น SharePoint)
    Returns: {"success": bool, "url": str(path ในเครื่อง), "error": str}
    """
    try:
        _ensure_dirs()
        folder = UPLOAD_DIR if not subfolder else os.path.join(UPLOAD_DIR, subfolder)
        os.makedirs(folder, exist_ok=True)

        # กันไฟล์ชื่อซ้ำ — เติม timestamp ถ้ามีอยู่แล้ว
        path = os.path.join(folder, filename)
        if os.path.exists(path):
            name, ext = os.path.splitext(filename)
            path = os.path.join(folder, f"{name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{ext}")

        with open(path, "wb") as f:
            f.write(file_bytes)
        return {"success": True, "url": path, "error": ""}
    except Exception as e:
        return {"success": False, "url": "", "error": str(e)}


def download_file(path: str) -> bytes | None:
    """อ่านไฟล์จากเครื่องกลับมาเป็น bytes"""
    try:
        with open(path, "rb") as f:
            return f.read()
    except Exception:
        return None


def list_files(subfolder: str = "") -> list[dict]:
    """รายชื่อไฟล์ทั้งหมดใน uploads/"""
    _ensure_dirs()
    folder = UPLOAD_DIR if not subfolder else os.path.join(UPLOAD_DIR, subfolder)
    if not os.path.isdir(folder):
        return []
    out = []
    for fname in sorted(os.listdir(folder)):
        fpath = os.path.join(folder, fname)
        if os.path.isfile(fpath):
            stat = os.stat(fpath)
            out.append({
                "name": fname,
                "url": fpath,
                "size": stat.st_size,
                "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%d/%m/%Y %H:%M"),
            })
    return out


# ══════════════════════════════════════════════════════════════════════════════
# ฟังก์ชันเสริม — บันทึกตารางรวม + audit log ให้อยู่ถาวร (ปิดแอพแล้วไม่หาย)
# ══════════════════════════════════════════════════════════════════════════════

def save_merged_snapshot(df: pd.DataFrame) -> str:
    """
    บันทึกตารางรวมเป็น snapshot
    ใช้ .parquet ถ้ามี pyarrow (เล็ก+เร็ว) — ไม่มีก็ fallback เป็น .csv อัตโนมัติ
    คืน path ที่บันทึก
    """
    _ensure_dirs()
    try:
        # ── พยายามใช้ parquet ก่อน (ต้องมี pyarrow) ──
        path = os.path.join(MERGED_DIR, "merged_latest.parquet")
        df.to_parquet(path, index=False)
        daily = os.path.join(MERGED_DIR, f"merged_{datetime.now().strftime('%Y%m%d')}.parquet")
        df.to_parquet(daily, index=False)
        return path
    except ImportError:
        # ── ไม่มี pyarrow → ใช้ .csv แทน (ติดมากับ pandas อยู่แล้ว) ──
        path = os.path.join(MERGED_DIR, "merged_latest.csv")
        df.to_csv(path, index=False, encoding="utf-8-sig")
        daily = os.path.join(MERGED_DIR, f"merged_{datetime.now().strftime('%Y%m%d')}.csv")
        df.to_csv(daily, index=False, encoding="utf-8-sig")
        return path


def load_merged_snapshot() -> pd.DataFrame | None:
    """โหลดตารางรวมล่าสุดกลับมา (ใช้ตอนเปิดแอพใหม่ — ข้อมูลไม่หาย)"""
    pq = os.path.join(MERGED_DIR, "merged_latest.parquet")
    cs = os.path.join(MERGED_DIR, "merged_latest.csv")
    # เลือกไฟล์ที่ใหม่กว่า ถ้ามีทั้งคู่
    candidates = [(p, os.path.getmtime(p)) for p in (pq, cs) if os.path.exists(p)]
    if not candidates:
        return None
    path = max(candidates, key=lambda x: x[1])[0]
    try:
        if path.endswith(".parquet"):
            return pd.read_parquet(path)
        return pd.read_csv(path, encoding="utf-8-sig")
    except Exception:
        return None


def append_audit(record: dict):
    """บันทึก audit log ต่อท้ายไฟล์ .csv (เก็บถาวร)"""
    _ensure_dirs()
    path = os.path.join(AUDIT_DIR, "audit_log.csv")
    df = pd.DataFrame([record])
    df.to_csv(path, mode="a", header=not os.path.exists(path),
              index=False, encoding="utf-8-sig")


def load_audit() -> pd.DataFrame | None:
    """โหลด audit log ทั้งหมด"""
    path = os.path.join(AUDIT_DIR, "audit_log.csv")
    if os.path.exists(path):
        try:
            return pd.read_csv(path, encoding="utf-8-sig")
        except Exception:
            return None
    return None


def storage_info() -> dict:
    """สรุปสถานะ storage — โชว์ใน sidebar"""
    _ensure_dirs()
    n_files = len(list_files())
    total_mb = sum(f["size"] for f in list_files()) / 1024 / 1024
    has_snapshot = os.path.exists(os.path.join(MERGED_DIR, "merged_latest.parquet"))
    return {"files": n_files, "total_mb": round(total_mb, 1), "snapshot": has_snapshot}
