"""
AI Leaders PINFL Checker — To'liq avtomatik
============================================
requirements.txt:
    streamlit
    openpyxl
    pandas
    requests
"""

import re
import time
import requests
import pandas as pd
import streamlit as st
from io import BytesIO

from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ──────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────
API_URL     = "https://aileaders.uz/api/v1/check/certificates"
DELAY_SEC   = 0.7
SKIP_SHEETS = {"ЖАМИ СЕРТИФИКАТ ОЛГАНЛАР", "Лист1"}

REPORT_COLS = ["Holat", "Ism (sayt)", "Email", "Kurslar", "Yakunlangan", "Xato"]

# ──────────────────────────────────────────
# cURL DAN COOKIE AJRATISH
# ──────────────────────────────────────────
def parse_curl(curl_text: str) -> dict:
    result = {"cookie": "", "user_agent": ""}

    cookie_match = re.search(r"-b\s+'([^']+)'", curl_text)
    if not cookie_match:
        cookie_match = re.search(r'-b\s+"([^"]+)"', curl_text)
    if not cookie_match:
        cookie_match = re.search(r"-H\s+'cookie:\s*([^']+)'", curl_text, re.IGNORECASE)
    if not cookie_match:
        cookie_match = re.search(r'-H\s+"cookie:\s*([^"]+)"', curl_text, re.IGNORECASE)
    if cookie_match:
        result["cookie"] = cookie_match.group(1).strip()

    ua_match = re.search(r"-H\s+'user-agent:\s*([^']+)'", curl_text, re.IGNORECASE)
    if not ua_match:
        ua_match = re.search(r"-H\s+\"user-agent:\s*([^\"]+)\"", curl_text, re.IGNORECASE)
    if ua_match:
        result["user_agent"] = ua_match.group(1).strip()

    return result

# ──────────────────────────────────────────
# EXCEL O'QISH
# ──────────────────────────────────────────
def read_excel(file) -> pd.DataFrame:
    xls = pd.ExcelFile(file, engine="openpyxl")
    frames = []
    rid = 1

    for sheet in xls.sheet_names:
        if sheet in SKIP_SHEETS:
            continue

        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=1, engine="openpyxl")
        except Exception:
            continue

        df.columns = df.columns.map(str).str.strip()

        pinfl_col = next(
            (c for c in df.columns if "ПИНФЛ" in c.upper() or "PINFL" in c.upper()),
            None
        )

        if pinfl_col is None:
            continue

        df["_PINFL_COL_"] = pinfl_col
        df["_RID_"] = range(rid, rid + len(df))
        rid += len(df)

        df["PINFL"] = (
            df[pinfl_col]
            .astype(str)
            .str.strip()
            .str.replace(r"\.0$", "", regex=True)
        )

        df["Maktab"] = sheet

        df = df[df["PINFL"].str.len() >= 10].copy()
        frames.append(df)

    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

# ──────────────────────────────────────────
# BITTA PINFL TEKSHIRISH
# ──────────────────────────────────────────
def check_pinfl(pinfl: str, session: requests.Session) -> dict:
    empty = {"holat": "", "ism": "", "email": "", "kurslar": 0, "yakunlangan": 0, "xato": ""}
    try:
        r = session.get(API_URL, params={"pinfl": pinfl}, timeout=15)

        if r.status_code in (200, 304):
            try:
                data = r.json()
            except Exception:
                return {**empty, "holat": "⚠️ JSON xato", "xato": r.text[:80]}

            courses     = data.get("courses", [])
            completed   = sum(1 for c in courses if c.get("isCompleted"))
            has_courses = data.get("hasCourses", False)

            return {
                "holat":       "✅ Sertifikat bor" if has_courses else "⚠️ Ro'yxatda bor, kurs yo'q",
                "ism":         data.get("fullName", ""),
                "email":       data.get("email", ""),
                "kurslar":     len(courses),
                "yakunlangan": completed,
                "xato":        "",
            }

        elif r.status_code == 404:
            return {**empty, "holat": "❌ Topilmadi"}
        elif r.status_code == 401:
            return {**empty, "holat": "🔐 Cookie eskirgan", "xato": "Cookie yangilang"}
        elif r.status_code == 429:
            time.sleep(15)
            return {**empty, "holat": "⏳ Rate limit", "xato": "Keyingi PINFL dan davom eting"}
        else:
            return {**empty, "holat": f"🔴 {r.status_code}", "xato": r.text[:80]}

    except Exception as e:
        return {**empty, "holat": "🔴 Xato", "xato": str(e)[:80]}

# ──────────────────────────────────────────
# EXCEL FORMATLASH
# ──────────────────────────────────────────
def style_excel(writer):
    wb = writer.book

    for ws in wb.worksheets:
        ws.freeze_panes = "A2"

        header_fill = PatternFill("solid", fgColor="1F4E78")
        header_font = Font(color="FFFFFF", bold=True)
        thin = Side(style="thin", color="B7B7B7")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(vertical="center", wrap_text=True)

        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)

            for cell in col:
                value = "" if cell.value is None else str(cell.value)
                max_len = max(max_len, len(value))

            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

# ──────────────────────────────────────────
# HISOBOT EXCEL YARATISH
# ──────────────────────────────────────────
def build_report_excel(df_all: pd.DataFrame, result_df: pd.DataFrame, summary: pd.DataFrame) -> BytesIO:
    export_df = df_all.copy()

    result_map = result_df.set_index("_RID_")[REPORT_COLS].to_dict("index")

    for col in REPORT_COLS:
        export_df[col] = export_df["_RID_"].map(lambda x: result_map.get(x, {}).get(col, ""))

    hidden_cols = ["_RID_", "_PINFL_COL_"]

    for col in hidden_cols:
        if col in export_df.columns:
            export_df = export_df.drop(columns=[col])

    # PINFL ustunidan keyin hisobot ustunlarini joylash
    cols = list(export_df.columns)

    for col in REPORT_COLS:
        if col in cols:
            cols.remove(col)

    pinfl_index = None
    for i, col in enumerate(cols):
        if "ПИНФЛ" in str(col).upper() or "PINFL" in str(col).upper():
            pinfl_index = i
            break

    if pinfl_index is None:
        pinfl_index = cols.index("PINFL") if "PINFL" in cols else 0

    final_cols = cols[:pinfl_index + 1] + REPORT_COLS + cols[pinfl_index + 1:]
    export_df = export_df[final_cols]

    out = BytesIO()

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        export_df.to_excel(writer, sheet_name="Natijalar", index=False)
        summary.to_excel(writer, sheet_name="Maktab xulosasi", index=False)
        style_excel(writer)

    out.seek(0)
    return out

# ──────────────────────────────────────────
# STREAMLIT UI
# ──────────────────────────────────────────
st.set_page_config(page_title="AI Leaders PINFL Checker", page_icon="🎓", layout="wide")

st.markdown("""
<style>
.stApp {
    background:
        radial-gradient(circle at top left, rgba(0,255,255,0.18), transparent 35%),
        radial-gradient(circle at top right, rgba(255,0,150,0.16), transparent 35%),
        linear-gradient(135deg, #07111f 0%, #111827 45%, #020617 100%);
    color: #E5E7EB;
}

.block-container {
    padding-top: 2rem;
    padding-bottom: 3rem;
}

h1, h2, h3 {
    color: #F8FAFC !important;
}

[data-testid="stMetric"] {
    background: rgba(15, 23, 42, 0.78);
    border: 1px solid rgba(56, 189, 248, 0.35);
    padding: 18px;
    border-radius: 20px;
    box-shadow: 0 0 25px rgba(56, 189, 248, 0.12);
}

[data-testid="stFileUploader"] {
    background: rgba(15, 23, 42, 0.72);
    border: 1px dashed rgba(34, 211, 238, 0.55);
    border-radius: 20px;
    padding: 18px;
}

.stTextArea textarea {
    background: rgba(2, 6, 23, 0.88) !important;
    color: #E0F2FE !important;
    border: 1px solid rgba(56, 189, 248, 0.55) !important;
    border-radius: 16px !important;
}

.stButton > button {
    background: linear-gradient(90deg, #06B6D4, #8B5CF6, #EC4899);
    color: white;
    border: none;
    border-radius: 16px;
    padding: 0.75rem 1.4rem;
    font-weight: 800;
    box-shadow: 0 0 25px rgba(139, 92, 246, 0.35);
}

.stDownloadButton > button {
    background: linear-gradient(90deg, #22C55E, #06B6D4);
    color: white;
    border: none;
    border-radius: 16px;
    padding: 0.75rem 1.4rem;
    font-weight: 800;
}

div[data-testid="stExpander"] {
    background: rgba(15, 23, 42, 0.68);
    border: 1px solid rgba(148, 163, 184, 0.22);
    border-radius: 18px;
}

[data-testid="stDataFrame"] {
    border-radius: 18px;
    overflow: hidden;
    box-shadow: 0 0 30px rgba(15, 23, 42, 0.35);
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
# 🎓 AI Leaders — PINFL Sertifikat Tekshiruvi
Excel fayldagi barcha PINFLlarni `aileaders.uz` orqali avtomatik tekshiradi.
""")

# ── 1-QADAM: cURL ──
st.subheader("🔐 1-qadam: cURL joylashtiring")

with st.expander("📋 cURL qanday olish kerak? (bosing)", expanded=True):
    st.markdown("""
    **Bir marta bajaring:**

    1. Chrome da **`https://aileaders.uz/auth/login/check`** oching  
    2. Istalgan PINFL kiriting → **Tekshirish** bosing  
    3. **F12** → **Network** tab  
    4. Pastdagi `certificates?pinfl=...` qatoriga **o'ng klik**  
    5. **Copy → Copy as cURL (bash)** tanlang  
    6. Quyidagi maydonga **Ctrl+V** bilan joylashtiring  
    """)

curl_input = st.text_area(
    "cURL matni:",
    placeholder="curl 'https://aileaders.uz/api/v1/check/certificates?pinfl=...' \\\n  -H 'cookie: HWWAFSESID=...' \\\n  ...",
    height=120,
)

parsed   = {}
curl_ok  = False

if curl_input.strip():
    parsed = parse_curl(curl_input)
    if parsed["cookie"] and "HWWAFSESID" in parsed["cookie"]:
        st.success("✅ cURL muvaffaqiyatli o'qildi! Cookie topildi.")
        curl_ok = True
    else:
        st.error("❌ Cookie topilmadi. cURL to'liq ko'chirilganmi?")
        st.code("Kerakli format:\ncurl '...' -H 'cookie: HWWAFSESID=...; HWWAFSESTIME=...'")

st.divider()

# ── 2-QADAM: EXCEL ──
st.subheader("📂 2-qadam: Excel yuklang")
uploaded = st.file_uploader("Excel fayl", type=["xlsx"], label_visibility="collapsed")

if uploaded:
    with st.spinner("Excel o'qilmoqda..."):
        df_all = read_excel(uploaded)

    if df_all.empty:
        st.error("❌ Excel dan PINFL topilmadi!")
        st.stop()

    st.success(f"✅ **{len(df_all)}** ta o'quvchi | **{df_all['Maktab'].nunique()}** ta maktab")

    with st.expander("🏫 Maktab filtri (ixtiyoriy)"):
        maktablar = sorted(df_all["Maktab"].unique().tolist())
        tanlangan = st.multiselect("Tekshiriladigan maktablar (bo'sh = hammasi)", maktablar)
        if tanlandan := tanlangan:
            df_all = df_all[df_all["Maktab"].isin(tanlandan)].copy()
            st.info(f"Tanlangan: {len(df_all)} ta o'quvchi")

    daqiqa = round(len(df_all) * DELAY_SEC / 60, 1)
    st.info(f"⏱️ Taxminiy vaqt: **{daqiqa} daqiqa** ({len(df_all)} ta × {DELAY_SEC}s)")

    st.divider()

    # ── 3-QADAM: TEKSHIRISH ──
    st.subheader("🚀 3-qadam: Tekshirishni boshlash")

    if not curl_ok:
        st.warning("⚠️ Avval 1-qadamda cURL ni joylashtiring!")

    if curl_ok and st.button("🚀 Tekshirishni boshlash", type="primary"):

        session = requests.Session()
        session.headers.update({
            "User-Agent": parsed.get("user_agent") or "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36",
            "Accept":     "*/*",
            "Referer":    "https://aileaders.uz/auth/login/check",
            "Cookie":     parsed["cookie"],
        })

        progress_bar = st.progress(0)
        status_text  = st.empty()
        results      = []
        total        = len(df_all)
        cookie_dead  = False

        for i, (_, row) in enumerate(df_all.iterrows(), 1):
            pinfl = str(row["PINFL"])
            maktab = row["Maktab"]

            status_text.text(f"🔄 {i}/{total} — {pinfl} | {maktab}")
            progress_bar.progress(i / total)

            res = check_pinfl(pinfl, session)

            if res["holat"] == "🔐 Cookie eskirgan":
                st.error("🔐 Cookie eskirdi! Yangi cURL olib, qayta boshlang.")
                cookie_dead = True
                break

            name_col = next(
                (c for c in df_all.columns if "наименование" in str(c).lower()),
                None
            )

            fio = row.get(name_col, "") if name_col else ""

            results.append({
                "_RID_":       row["_RID_"],
                "Maktab":      maktab,
                "F.I.Sh.":     fio,
                "PINFL":       pinfl,
                "Holat":       res["holat"],
                "Ism (sayt)":  res["ism"],
                "Email":       res["email"],
                "Kurslar":     res["kurslar"],
                "Yakunlangan": res["yakunlangan"],
                "Xato":        res["xato"],
            })

            time.sleep(DELAY_SEC)

        if not cookie_dead:
            status_text.success(f"✅ Yakunlandi! {len(results)}/{total} ta tekshirildi")

        progress_bar.progress(1.0)

        if not results:
            st.stop()

        result_df = pd.DataFrame(results)

        # Statistika
        st.subheader("📊 Natijalar")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Jami",          len(result_df))
        c2.metric("✅ Sertifikat", (result_df["Holat"] == "✅ Sertifikat bor").sum())
        c3.metric("⚠️ Kurs yo'q",  (result_df["Holat"].str.contains("Ro'yxatda", na=False)).sum())
        c4.metric("❌ Topilmadi",  (result_df["Holat"] == "❌ Topilmadi").sum())

        st.dataframe(result_df.drop(columns=["_RID_"]), use_container_width=True)

        # Maktab xulosasi
        st.subheader("🏫 Maktab bo'yicha hisobot")
        summary = (
            result_df.groupby("Maktab")
            .agg(
                Jami=("PINFL", "count"),
                Sertifikat=("Holat", lambda x: (x == "✅ Sertifikat bor").sum()),
                Kurs_yoq=("Holat", lambda x: (x.str.contains("Ro'yxatda", na=False)).sum()),
                Topilmadi=("Holat", lambda x: (x == "❌ Topilmadi").sum()),
            )
            .assign(Foiz=lambda d: (d["Sertifikat"] / d["Jami"] * 100).round(1))
            .sort_values("Foiz", ascending=False)
            .reset_index()
        )

        st.dataframe(summary, use_container_width=True)

        # Excel yuklab olish
        out = build_report_excel(df_all, result_df, summary)

        st.download_button(
            "📥 Hisobotni yuklab olish (Excel)",
            data=out,
            file_name="aileaders_hisobot.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
