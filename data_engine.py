"""
RangeSheet — Data Engine
อ่านไฟล์ .csv/.xlsx + Auto-merge ตาม header similarity
ตามแผน W3: "Develop ingestion scripts specifically for .csv files,
populate initial data frames using Pandas"
"""
import io
import os
import pandas as pd
from datetime import datetime
from config import APP_CONFIG


# ── File reading ──────────────────────────────────────────────────────────────
def read_uploaded_file(uploaded_file) -> pd.DataFrame | None:
    """
    อ่านไฟล์ที่อัปโหลด (.csv primary ตามแผน, รองรับ .xlsx/.xls ด้วย)
    คืน DataFrame หรือ None ถ้าอ่านไม่ได้
    """
    name = uploaded_file.name
    ext = os.path.splitext(name)[-1].lower()
    try:
        if ext == ".csv":
            # ลองหลาย encoding — ไฟล์ไทยมักเป็น cp874 หรือ utf-8-sig
            raw = uploaded_file.read()
            uploaded_file.seek(0)
            for enc in ["utf-8-sig", "utf-8", "cp874", "tis-620", "latin1"]:
                try:
                    df = pd.read_csv(io.BytesIO(raw), encoding=enc)
                    break
                except (UnicodeDecodeError, pd.errors.ParserError):
                    continue
            else:
                return None
        elif ext in (".xlsx", ".xls"):
            df = pd.read_excel(uploaded_file)
        else:
            return None

        # Clean headers: strip whitespace
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return None


# ── Header similarity ─────────────────────────────────────────────────────────
def header_similarity(cols_a, cols_b) -> float:
    """
    วัดความคล้ายของ header สองชุด (0.0–1.0)
    ใช้ set intersection / max set size
    """
    set_a = {str(c).lower().strip() for c in cols_a}
    set_b = {str(c).lower().strip() for c in cols_b}
    if not set_a or not set_b:
        return 0.0
    return len(set_a & set_b) / max(len(set_a), len(set_b))


# ── Auto-merge ────────────────────────────────────────────────────────────────
def auto_merge(files_info: list[dict]) -> tuple[pd.DataFrame | None, list[str]]:
    """
    Auto-merge ไฟล์ทั้งหมด:
    - header คล้าย ≥ threshold (default 50%) และข้อมูล column ตรงกัน → concat ต่อแถว (ตารางยาว)
    - header เหมือนบางส่วน → outer merge บน common columns
    - ไม่คล้ายเลย → stack ต่อท้าย

    Returns: (merged_df, log_messages)
    """
    threshold = APP_CONFIG["merge_similarity_threshold"]
    dfs = [(f["name"], f["df"]) for f in files_info if f.get("df") is not None]

    if not dfs:
        return None, []
    if len(dfs) == 1:
        return dfs[0][1].copy(), [f"📄 ไฟล์เดียว — ไม่ต้อง merge ({dfs[0][0]})"]

    log = []
    base_name, base = dfs[0]
    base = base.copy()
    log.append(f"📌 ใช้ **{base_name}** เป็นฐาน ({len(base):,} แถว)")

    for fname, df in dfs[1:]:
        sim = header_similarity(base.columns, df.columns)
        common = [c for c in base.columns if c in df.columns]

        if sim >= threshold and len(common) == len(base.columns) == len(df.columns):
            # Header เหมือนกันทุก column → ต่อแถวเป็นตารางยาว (ตาม spec)
            df_aligned = df[list(base.columns)]
            base = pd.concat([base, df_aligned], ignore_index=True)
            log.append(f"✅ **{fname}** — header ตรงกัน 100% → ต่อแถว ({len(df):,} แถว, รวม {len(base):,})")

        elif sim >= threshold and common:
            # คล้ายมาก → ต่อแถวเฉพาะ common columns + เก็บ columns ที่ต่างไว้
            base = pd.concat([base, df], ignore_index=True, sort=False)
            log.append(f"✅ **{fname}** — คล้าย {sim:.0%} ({len(common)} cols ร่วม) → ต่อแถว")

        else:
            base = pd.concat([base, df], ignore_index=True, sort=False)
            log.append(f"📋 **{fname}** — คล้ายแค่ {sim:.0%} → stack ต่อท้าย (ตรวจสอบ column)")

    # Dedupe exact duplicate rows ที่อาจเกิดจากไฟล์ซ้ำ
    before = len(base)
    base = base.drop_duplicates(ignore_index=True)
    if len(base) < before:
        log.append(f"🧹 ลบแถวซ้ำ {before - len(base):,} แถว")

    return base, log


# ── Export helpers ────────────────────────────────────────────────────────────
def df_to_xlsx_bytes(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()

def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")  # utf-8-sig เปิดใน Excel ภาษาไทยได้
