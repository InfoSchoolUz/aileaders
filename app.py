"""
AI Leaders PINFL Checker — Premium Edition
Developer: Azamat Madrimov
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
DELAY_SEC   = 3.5   # Rate limit: 20 req/min → xavfsiz 3.5s
SKIP_SHEETS = {"ЖАМИ СЕРТИФИКАТ ОЛГАНЛАР", "Лист1"}
REPORT_COLS = [
    "Tekshiruv holati", "Saytdagi F.I.Sh.", "Email",
    "Kurslar soni", "Yakunlangan kurslar", "Kurslar tafsiloti", "Izoh"
]

# ──────────────────────────────────────────
# CURL PARSER
# ──────────────────────────────────────────
def parse_curl(curl_text: str) -> dict:
    result = {"cookie": "", "user_agent": ""}
    for pattern in [r"-b\s+'([^']+)'", r'-b\s+"([^"]+)"',
                    r"-H\s+'cookie:\s*([^']+)'", r'-H\s+"cookie:\s*([^"]+)"']:
        m = re.search(pattern, curl_text, re.IGNORECASE)
        if m:
            result["cookie"] = m.group(1).strip()
            break
    for pattern in [r"-H\s+'user-agent:\s*([^']+)'", r'-H\s+"user-agent:\s*([^"]+)"']:
        m = re.search(pattern, curl_text, re.IGNORECASE)
        if m:
            result["user_agent"] = m.group(1).strip()
            break
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
            (c for c in df.columns if "ПИНФЛ" in str(c).upper() or "PINFL" in str(c).upper()), None
        )
        if pinfl_col is None:
            continue
        df["_PINFL_COL_"] = pinfl_col
        df["_RID_"] = range(rid, rid + len(df))
        rid += len(df)
        df["PINFL"] = (
            df[pinfl_col].astype(str).str.strip()
            .str.replace(r"\.0$", "", regex=True)
        )
        df["Maktab"] = sheet
        df = df[df["PINFL"].str.len() >= 10].copy()
        frames.append(df)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()

# ──────────────────────────────────────────
# PINFL TEKSHIRISH — rate limit qayta urinish bilan
# ──────────────────────────────────────────
def check_pinfl(pinfl: str, session: requests.Session) -> dict:
    empty = {"holat": "", "ism": "", "email": "", "kurslar": 0,
             "yakunlangan": 0, "kurs_tafsiloti": "", "xato": ""}
    try:
        r = session.get(API_URL, params={"pinfl": pinfl}, timeout=15)

        if r.status_code == 200:
            try:
                data = r.json()
            except Exception:
                return {**empty, "holat": "⚠️ JSON xato", "xato": r.text[:80]}

            courses   = data.get("courses", [])
            completed = sum(1 for c in courses if c.get("isCompleted"))

            if completed > 0:
                holat = "✅ Sertifikat olgan"
            elif len(courses) > 0:
                holat = "⚠️ Kurs bor, sertifikat olinmagan"
            else:
                holat = "⚠️ PINFL bor, kurs topilmadi"

            kurslar = []
            for c in courses:
                nomi     = (c.get("courseName") or c.get("name") or c.get("title") or "Noma'lum kurs")
                hamkor   = (c.get("partner") or c.get("partnerName") or c.get("provider") or "")
                progress = (c.get("progress") or c.get("progressPercent") or "")
                davom    = (c.get("duration") or c.get("approxTotalCourseHrs") or "")
                yozildi  = (c.get("enrolledAt") or c.get("createdAt") or "")
                tugaldi  = (c.get("completedAt") or "")
                ochirildi= (c.get("deletedAt") or "")
                tugatdimi= "Ha" if c.get("isCompleted") else "Yo'q"
                qator = f"Kurs: {nomi}"
                if hamkor:   qator += f" | Hamkor: {hamkor}"
                if progress != "": qator += f" | Progress: {progress}%"
                if davom:    qator += f" | Davomiylik: {davom}h"
                if yozildi:  qator += f" | Yozilgan: {yozildi}"
                if tugaldi:  qator += f" | Tugallangan: {tugaldi}"
                if ochirildi: qator += f" | O'chirilgan: {ochirildi}"
                qator += f" | Sertifikat: {tugatdimi}"
                kurslar.append(qator)

            return {
                "holat": holat,
                "ism":   data.get("fullName", ""),
                "email": data.get("email", ""),
                "kurslar":     len(courses),
                "yakunlangan": completed,
                "kurs_tafsiloti": "\n".join(kurslar),
                "xato": "",
            }

        elif r.status_code == 404:
            return {**empty, "holat": "❌ Ro'yxatdan o'tmagan"}
        elif r.status_code == 401:
            return {**empty, "holat": "🔐 Cookie eskirgan", "xato": "Cookie yangilang"}
        elif r.status_code == 429:
            return {**empty, "holat": "⏳ Rate limit", "xato": "limit"}
        else:
            return {**empty, "holat": f"🔴 Server xatosi: {r.status_code}", "xato": r.text[:80]}
    except Exception as e:
        return {**empty, "holat": "🔴 Xato", "xato": str(e)[:80]}


def check_pinfl_safe(pinfl: str, session: requests.Session, status_el, i: int, total: int) -> dict:
    """Rate limit bo'lsa 3 marta qayta urinadi."""
    for urinish in range(3):
        res = check_pinfl(pinfl, session)
        if res["holat"] != "⏳ Rate limit":
            return res
        wait = 20 + urinish * 10
        for s in range(wait, 0, -1):
            status_el.markdown(
                f'<div class="status-bar">⏳ {i}/{total} — Rate limit! {s}s kutilmoqda... (urinish {urinish+1}/3)</div>',
                unsafe_allow_html=True
            )
            time.sleep(1)
    return {**{"holat": "⏳ Rate limit (3 marta urinildi)", "ism": "", "email": "",
               "kurslar": 0, "yakunlangan": 0, "kurs_tafsiloti": "", "xato": "limit"}}

# ──────────────────────────────────────────
# EXCEL STYLE
# ──────────────────────────────────────────
def style_excel(writer):
    wb = writer.book
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        hfill  = PatternFill("solid", fgColor="1F4E78")
        hfont  = Font(color="FFFFFF", bold=True)
        thin   = Side(style="thin", color="B7B7B7")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(vertical="center", wrap_text=True)
        for cell in ws[1]:
            cell.fill = hfill
            cell.font = hfont
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for col in ws.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = max((len(str(c.value or "")) for c in col), default=10)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 60)

def build_report_excel(df_all, result_df, summary) -> BytesIO:
    export_df  = df_all.copy()
    result_map = result_df.set_index("_RID_")[REPORT_COLS].to_dict("index")
    for col in REPORT_COLS:
        export_df[col] = export_df["_RID_"].map(lambda x, c=col: result_map.get(x, {}).get(c, ""))
    for col in ["_RID_", "_PINFL_COL_"]:
        if col in export_df.columns:
            export_df = export_df.drop(columns=[col])
    cols = list(export_df.columns)
    for col in REPORT_COLS:
        if col in cols:
            cols.remove(col)
    pinfl_index = next(
        (i for i, c in enumerate(cols) if "ПИНФЛ" in str(c).upper() or c == "PINFL"), 0
    )
    final_cols = cols[:pinfl_index + 1] + REPORT_COLS + cols[pinfl_index + 1:]
    export_df  = export_df[[c for c in final_cols if c in export_df.columns]]
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        export_df.to_excel(writer, sheet_name="Natijalar",       index=False)
        summary.to_excel(  writer, sheet_name="Maktab xulosasi", index=False)
        style_excel(writer)
    out.seek(0)
    return out

# ──────────────────────────────────────────
# PREMIUM CSS
# ──────────────────────────────────────────
PREMIUM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap');

*, *::before, *::after { box-sizing: border-box; }

html, body, [data-testid="stAppViewContainer"] {
    background: #020817 !important;
    color: #CBD5E1;
    font-family: 'Syne', sans-serif;
}

/* Mesh gradient background */
[data-testid="stAppViewContainer"]::before {
    content: '';
    position: fixed;
    inset: 0;
    background:
        radial-gradient(ellipse 80% 50% at 10% 0%, rgba(56,189,248,0.08) 0%, transparent 60%),
        radial-gradient(ellipse 60% 40% at 90% 10%, rgba(168,85,247,0.07) 0%, transparent 55%),
        radial-gradient(ellipse 50% 60% at 50% 100%, rgba(20,184,166,0.05) 0%, transparent 60%);
    pointer-events: none;
    z-index: 0;
}

[data-testid="stVerticalBlock"] { position: relative; z-index: 1; }

.block-container {
    max-width: 1100px !important;
    padding: 2.5rem 2rem 4rem !important;
}

/* ── HEADER ── */
.app-header {
    text-align: center;
    padding: 3rem 0 2rem;
    border-bottom: 1px solid rgba(56,189,248,0.12);
    margin-bottom: 2.5rem;
}
.app-logo {
    display: inline-flex;
    align-items: center;
    gap: 10px;
    background: linear-gradient(135deg, rgba(56,189,248,0.12), rgba(168,85,247,0.10));
    border: 1px solid rgba(56,189,248,0.25);
    border-radius: 16px;
    padding: 8px 20px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 12px;
    color: #38BDF8;
    letter-spacing: 3px;
    text-transform: uppercase;
    margin-bottom: 1.2rem;
}
.app-title {
    font-size: clamp(2rem, 5vw, 3.2rem);
    font-weight: 800;
    background: linear-gradient(135deg, #F8FAFC 0%, #94A3B8 60%, #38BDF8 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    line-height: 1.1;
    margin: 0 0 0.6rem;
}
.app-subtitle {
    color: #64748B;
    font-size: 1rem;
    font-family: 'JetBrains Mono', monospace;
    letter-spacing: 0.5px;
}

/* ── STEP CARDS ── */
.step-card {
    background: rgba(15,23,42,0.70);
    border: 1px solid rgba(51,65,85,0.60);
    border-radius: 20px;
    padding: 1.8rem 2rem;
    margin-bottom: 1.5rem;
    backdrop-filter: blur(12px);
    transition: border-color 0.3s;
}
.step-card:hover { border-color: rgba(56,189,248,0.30); }
.step-header {
    display: flex;
    align-items: center;
    gap: 12px;
    margin-bottom: 1.2rem;
}
.step-num {
    width: 36px; height: 36px;
    border-radius: 50%;
    background: linear-gradient(135deg, #0EA5E9, #8B5CF6);
    display: flex; align-items: center; justify-content: center;
    font-weight: 800; font-size: 14px; color: white;
    flex-shrink: 0;
}
.step-title {
    font-size: 1.1rem;
    font-weight: 700;
    color: #F1F5F9;
}

/* ── METRICS ── */
[data-testid="metric-container"] {
    background: rgba(15,23,42,0.80) !important;
    border: 1px solid rgba(51,65,85,0.70) !important;
    border-radius: 18px !important;
    padding: 1.2rem 1.4rem !important;
    backdrop-filter: blur(10px);
}
[data-testid="stMetricValue"] {
    font-family: 'Syne', sans-serif !important;
    font-weight: 800 !important;
    font-size: 2rem !important;
    color: #F8FAFC !important;
}
[data-testid="stMetricLabel"] {
    color: #64748B !important;
    font-size: 0.8rem !important;
    font-family: 'JetBrains Mono', monospace !important;
    letter-spacing: 0.5px !important;
}

/* ── INPUTS ── */
.stTextArea textarea {
    background: rgba(2,6,23,0.90) !important;
    color: #94A3B8 !important;
    border: 1px solid rgba(51,65,85,0.80) !important;
    border-radius: 14px !important;
    font-family: 'JetBrains Mono', monospace !important;
    font-size: 12px !important;
    transition: border-color 0.3s !important;
}
.stTextArea textarea:focus {
    border-color: rgba(56,189,248,0.50) !important;
    box-shadow: 0 0 0 3px rgba(56,189,248,0.08) !important;
}

/* ── BUTTONS ── */
.stButton > button {
    background: linear-gradient(135deg, #0EA5E9 0%, #8B5CF6 50%, #EC4899 100%) !important;
    color: white !important;
    border: none !important;
    border-radius: 14px !important;
    padding: 0.7rem 1.8rem !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    font-size: 0.95rem !important;
    letter-spacing: 0.3px !important;
    box-shadow: 0 4px 24px rgba(139,92,246,0.25) !important;
    transition: all 0.2s !important;
}
.stButton > button:hover {
    transform: translateY(-1px) !important;
    box-shadow: 0 8px 32px rgba(139,92,246,0.35) !important;
}
.stDownloadButton > button {
    background: linear-gradient(135deg, #059669 0%, #0EA5E9 100%) !important;
    color: white !important;
    border: none !important;
    border-radius: 14px !important;
    padding: 0.7rem 1.8rem !important;
    font-family: 'Syne', sans-serif !important;
    font-weight: 700 !important;
    box-shadow: 0 4px 20px rgba(5,150,105,0.25) !important;
}

/* ── PROGRESS ── */
.status-bar {
    font-family: 'JetBrains Mono', monospace;
    font-size: 12px;
    color: #38BDF8;
    background: rgba(14,165,233,0.06);
    border: 1px solid rgba(14,165,233,0.15);
    border-radius: 10px;
    padding: 10px 16px;
    margin: 8px 0;
}
[data-testid="stProgress"] > div > div {
    background: linear-gradient(90deg, #0EA5E9, #8B5CF6, #EC4899) !important;
    border-radius: 99px !important;
}
[data-testid="stProgress"] > div {
    background: rgba(30,41,59,0.80) !important;
    border-radius: 99px !important;
    height: 8px !important;
}

/* ── EXPANDER ── */
[data-testid="stExpander"] {
    background: rgba(15,23,42,0.60) !important;
    border: 1px solid rgba(51,65,85,0.50) !important;
    border-radius: 16px !important;
    overflow: hidden;
}

/* ── ALERTS ── */
[data-testid="stSuccess"] {
    background: rgba(5,150,105,0.10) !important;
    border: 1px solid rgba(5,150,105,0.30) !important;
    border-radius: 12px !important;
    color: #6EE7B7 !important;
}
[data-testid="stWarning"] {
    background: rgba(217,119,6,0.10) !important;
    border: 1px solid rgba(217,119,6,0.30) !important;
    border-radius: 12px !important;
}
[data-testid="stError"] {
    background: rgba(220,38,38,0.10) !important;
    border: 1px solid rgba(220,38,38,0.30) !important;
    border-radius: 12px !important;
}
[data-testid="stInfo"] {
    background: rgba(14,165,233,0.08) !important;
    border: 1px solid rgba(14,165,233,0.25) !important;
    border-radius: 12px !important;
    color: #7DD3FC !important;
}

/* ── DATAFRAME ── */
[data-testid="stDataFrame"] {
    border-radius: 16px !important;
    overflow: hidden !important;
    border: 1px solid rgba(51,65,85,0.50) !important;
}

/* ── FILE UPLOADER ── */
[data-testid="stFileUploader"] {
    background: rgba(15,23,42,0.60) !important;
    border: 1.5px dashed rgba(56,189,248,0.30) !important;
    border-radius: 18px !important;
    padding: 1rem !important;
    transition: border-color 0.3s !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: rgba(56,189,248,0.55) !important;
}

/* ── MULTISELECT ── */
[data-testid="stMultiSelect"] span {
    background: rgba(14,165,233,0.15) !important;
    border: 1px solid rgba(14,165,233,0.30) !important;
    border-radius: 8px !important;
    color: #7DD3FC !important;
}

/* ── DIVIDER ── */
hr {
    border-color: rgba(51,65,85,0.40) !important;
    margin: 2rem 0 !important;
}

/* ── FOOTER ── */
.app-footer {
    text-align: center;
    padding: 2.5rem 0 1rem;
    border-top: 1px solid rgba(51,65,85,0.30);
    margin-top: 3rem;
}
.footer-badge {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    background: rgba(15,23,42,0.80);
    border: 1px solid rgba(51,65,85,0.60);
    border-radius: 99px;
    padding: 8px 20px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 12px;
    color: #475569;
}
.footer-badge span { color: #38BDF8; font-weight: 600; }

/* ── SCROLLBAR ── */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb {
    background: rgba(51,65,85,0.70);
    border-radius: 99px;
}
</style>
"""

# ──────────────────────────────────────────
# STREAMLIT APP
# ──────────────────────────────────────────
st.set_page_config(
    page_title="AI Leaders PINFL Checker",
    page_icon="🎓",
    layout="wide",
)

st.markdown(PREMIUM_CSS, unsafe_allow_html=True)

# ── HEADER ──
st.markdown("""
<div class="app-header">
    <div class="app-logo">⚡ InfoSchoolUz · Khorezm</div>
    <div class="app-title">PINFL Sertifikat Tekshiruvi</div>
    <div class="app-subtitle">aileaders.uz · avtomatik bulk tekshiruv tizimi</div>
</div>
""", unsafe_allow_html=True)

# ── 1-QADAM: cURL ──
st.markdown("""
<div class="step-card">
    <div class="step-header">
        <div class="step-num">1</div>
        <div class="step-title">cURL — Autentifikatsiya</div>
    </div>
</div>
""", unsafe_allow_html=True)

with st.expander("📋 cURL qanday olish kerak?", expanded=True):
    st.markdown("""
**Chrome da bir marta bajaring:**

1. `https://aileaders.uz/auth/login/check` sahifasini oching
2. Istalgan PINFL kiriting → **Tekshirish** bosing
3. **F12** → **Network** tab
4. `certificates?pinfl=...` qatoriga **o'ng klik**
5. **Copy → Copy as cURL (bash)** tanlang
6. Quyidagi maydonga **Ctrl+V** bilan joylashtiring
    """)

curl_input = st.text_area(
    "cURL matni:",
    placeholder="curl 'https://aileaders.uz/api/v1/check/certificates?pinfl=...' \\\n  -b 'HWWAFSESID=...; HWWAFSESTIME=...' \\\n  -H 'user-agent: Mozilla/5.0 ...'",
    height=110,
    label_visibility="collapsed",
)

parsed  = {}
curl_ok = False

if curl_input.strip():
    parsed = parse_curl(curl_input)
    if parsed["cookie"] and "HWWAFSESID" in parsed["cookie"]:
        st.success("✅ Cookie muvaffaqiyatli aniqlandi — tayyor!")
        curl_ok = True
    else:
        st.error("❌ Cookie topilmadi. cURL to'liq ko'chirilganmi?")

st.divider()

# ── 2-QADAM: EXCEL ──
st.markdown("""
<div class="step-card">
    <div class="step-header">
        <div class="step-num">2</div>
        <div class="step-title">Excel — O'quvchilar ro'yxati</div>
    </div>
</div>
""", unsafe_allow_html=True)

uploaded = st.file_uploader(
    "Excel fayl (.xlsx)",
    type=["xlsx"],
    label_visibility="collapsed",
)

if uploaded:
    with st.spinner("Fayl tahlil qilinmoqda..."):
        df_all = read_excel(uploaded)

    if df_all.empty:
        st.error("❌ Excel dan PINFL ustuni topilmadi!")
        st.stop()

    col_a, col_b = st.columns(2)
    col_a.success(f"✅ **{len(df_all):,}** ta o'quvchi yuklandi")
    col_b.info(f"🏫 **{df_all['Maktab'].nunique()}** ta maktab aniqlandi")

    with st.expander("🏫 Maktab filtri (ixtiyoriy)"):
        maktablar = sorted(df_all["Maktab"].unique().tolist())
        tanlangan = st.multiselect(
            "Faqat quyidagi maktablarni tekshirish (bo'sh = hammasi)",
            maktablar,
        )
        if tanlandan := tanlangan:
            df_all = df_all[df_all["Maktab"].isin(tanlandan)].copy()
            st.info(f"Tanlangan: **{len(df_all)}** ta o'quvchi")

    daqiqa = round(len(df_all) * DELAY_SEC / 60, 1)
    soat   = round(daqiqa / 60, 1)
    st.info(
        f"⏱️ Taxminiy vaqt: **{daqiqa} daqiqa** (~{soat} soat) · "
        f"{len(df_all):,} ta × {DELAY_SEC}s · Rate limit: 20 req/min"
    )

    st.divider()

    # ── 3-QADAM: TEKSHIRISH ──
    st.markdown("""
    <div class="step-card">
        <div class="step-header">
            <div class="step-num">3</div>
            <div class="step-title">Tekshirishni boshlash</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if not curl_ok:
        st.warning("⚠️ Avval 1-qadamda cURL ni joylashtiring!")

    if curl_ok and st.button("🚀 Tekshirishni boshlash", type="primary"):

        session = requests.Session()
        session.headers.update({
            "User-Agent": parsed.get("user_agent") or (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36"
            ),
            "Accept":   "*/*",
            "Referer":  "https://aileaders.uz/auth/login/check",
            "Cookie":   parsed["cookie"],
        })

        progress_bar = st.progress(0)
        status_el    = st.empty()
        results      = []
        total        = len(df_all)
        cookie_dead  = False

        name_col = next(
            (c for c in df_all.columns if "наименование" in str(c).lower()), None
        )

        for i, (_, row) in enumerate(df_all.iterrows(), 1):
            pinfl  = str(row["PINFL"])
            maktab = str(row["Maktab"])

            status_el.markdown(
                f'<div class="status-bar">🔄 {i}/{total} — <b>{pinfl}</b> · {maktab}</div>',
                unsafe_allow_html=True,
            )
            progress_bar.progress(i / total)

            # Rate limit bo'lsa qayta urinish
            res = check_pinfl_safe(pinfl, session, status_el, i, total)

            if res["holat"] == "🔐 Cookie eskirgan":
                st.error("🔐 Cookie eskirdi! Yangi cURL olib, 1-qadamdan qayta boshlang.")
                cookie_dead = True
                break

            fio = str(row.get(name_col, "") or "") if name_col else ""

            results.append({
                "_RID_":               row["_RID_"],
                "Maktab":              maktab,
                "F.I.Sh.":             fio,
                "PINFL":               pinfl,
                "Tekshiruv holati":    res["holat"],
                "Saytdagi F.I.Sh.":   res["ism"],
                "Email":               res["email"],
                "Kurslar soni":        res["kurslar"],
                "Yakunlangan kurslar": res["yakunlangan"],
                "Kurslar tafsiloti":   res["kurs_tafsiloti"],
                "Izoh":                res["xato"],
            })
            time.sleep(DELAY_SEC)

        if not cookie_dead:
            status_el.markdown(
                f'<div class="status-bar">✅ Yakunlandi — {len(results)}/{total} ta tekshirildi</div>',
                unsafe_allow_html=True,
            )
        progress_bar.progress(1.0)

        if not results:
            st.stop()

        result_df = pd.DataFrame(results)

        # ── STATISTIKA ──
        st.subheader("📊 Natijalar")
        c1, c2, c3, c4 = st.columns(4)
        n_total    = len(result_df)
        n_sert     = (result_df["Tekshiruv holati"] == "✅ Sertifikat olgan").sum()
        n_kurs     = result_df["Tekshiruv holati"].str.contains("Kurs bor|kurs topilmadi", case=False, na=False).sum()
        n_yoq      = (result_df["Tekshiruv holati"] == "❌ Ro'yxatdan o'tmagan").sum()

        c1.metric("Jami tekshirildi",    f"{n_total:,}")
        c2.metric("✅ Sertifikat olgan", f"{n_sert:,}")
        c3.metric("⚠️ Sertifikatsiz",   f"{n_kurs:,}")
        c4.metric("❌ Ro'yxatda yo'q",  f"{n_yoq:,}")

        st.dataframe(
            result_df.drop(columns=["_RID_"]),
            use_container_width=True,
            height=400,
        )

        # ── MAKTAB XULOSASI ──
        st.subheader("🏫 Maktab bo'yicha hisobot")
        summary = (
            result_df.groupby("Maktab")
            .agg(
                Jami=("PINFL", "count"),
                Sertifikat_olgan=("Tekshiruv holati", lambda x: (x == "✅ Sertifikat olgan").sum()),
                Sertifikatsiz=("Tekshiruv holati", lambda x: x.str.contains("Kurs bor|kurs topilmadi", case=False, na=False).sum()),
                Royxatdan_otmagan=("Tekshiruv holati", lambda x: (x == "❌ Ro'yxatdan o'tmagan").sum()),
            )
            .assign(Foiz=lambda d: (d["Sertifikat_olgan"] / d["Jami"] * 100).round(1))
            .sort_values("Foiz", ascending=False)
            .reset_index()
        )
        st.dataframe(summary, use_container_width=True)

        # ── YUKLAB OLISH ──
        with st.spinner("Excel hisobot tayyorlanmoqda..."):
            out = build_report_excel(df_all, result_df, summary)

        st.download_button(
            "📥 To'liq hisobotni yuklab olish (Excel)",
            data=out,
            file_name="aileaders_hisobot.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ── FOOTER ──
st.markdown("""
<div class="app-footer">
    <div class="footer-badge">
        Developed by <span>Azamat Madrimov</span> &nbsp;·&nbsp; InfoSchoolUz Khorezm &nbsp;·&nbsp; 2026
    </div>
</div>
""", unsafe_allow_html=True)
