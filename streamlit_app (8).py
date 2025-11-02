
# streamlit_app.py
# -*- coding: utf-8 -*-
import io
import pandas as pd
import numpy as np
from datetime import datetime
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="התאמות לקוחות – OV/RC", page_icon="✅", layout="centered")

# RTL UI
st.markdown('''
    <style>
    html, body, [class*="css"]  { direction: rtl; text-align: right; }
    .block-container { padding-top: 2rem; }
    </style>
''', unsafe_allow_html=True)

st.title("אפליקציה להתאמות לקוחות בבנק (175/120 + OV/RC)")

st.write("העלו קובץ אקסל מקורי. האפליקציה תחזיר את אותו קובץ **רק עם שינוי בעמודה _מס.התאמה_**, לפי הכללים:")
st.markdown("""
1. חיפוש שורות בנק עם **קוד פעולת בנק = 175 או 120** ו-**סכום בדף במינוס**.  
2. בשורה זו נבדקת עמודת **אסמכתא 1** שמתחילה ב-**OV** או **RC**.  
3. חיפוש בשורות הספרים באותו גיליון **סכום בספרים בפלוס** **באותו תאריך מאזן** ובאותו סכום האבסולוטי.  
4. כאשר נמצאת התאמה **מסמנים 1** בעמודה **מס.התאמה** (אפשר לבחור אם לסמן גם בשורת הספרים).
""")

# --- Helpers ---
def normalize_date(series):
    def to_date(x):
        if pd.isna(x): 
            return pd.NaT
        if isinstance(x, (pd.Timestamp, datetime)):
            return pd.Timestamp(x.date())
        try:
            return pd.to_datetime(x, dayfirst=True, errors="coerce").normalize()
        except Exception:
            return pd.NaT
    return series.apply(to_date)

def to_number(series):
    return pd.to_numeric(series, errors="coerce")

def ref_starts_with_ov_rc(text):
    t = (str(text) if text is not None else "").strip().upper()
    return t.startswith("OV") or t.startswith("RC")

def exact_or_contains(df, names):
    # prefer exact name, then contains
    for n in names:
        if n in df.columns:
            return n
    for n in names:
        for c in df.columns:
            if isinstance(c, str) and n in c:
                return c
    return None

MATCH_COL_CANDS = ["מס.התאמה","מס. התאמה","מס התאמה","מספר התאמה","התאמה"]
BANK_CODE_CANDS = ["קוד פעולת בנק","קוד פעולה","קוד פעולת ה"]
BANK_AMT_CANDS  = ["סכום בדף","סכום דף","סכום בבנק","סכום תנועת בנק"]
BOOKS_AMT_CANDS = ["סכום בספרים","סכום בספר","סכום ספרים"]
REF_CANDS       = ["אסמכתא 1","אסמכתא1","אסמכתא","אסמכתה"]
DATE_CANDS      = ["תאריך מאזן","תאריך ערך","תאריך"]

@st.cache_data(show_spinner=False)
def process_workbook(file_bytes, mark_both_rows: bool, require_ref_on_books: bool):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_names = xls.sheet_names

    writer_buf = io.BytesIO()
    with pd.ExcelWriter(writer_buf, engine="xlsxwriter") as writer:
        summary_rows = []
        for sh in sheet_names:
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sh, dtype=object)
            df_save = df.copy()

            col_match = exact_or_contains(df, MATCH_COL_CANDS) or df.columns[0]
            col_bank_code = exact_or_contains(df, BANK_CODE_CANDS)
            col_bank_amt  = exact_or_contains(df, BANK_AMT_CANDS)
            col_books_amt = exact_or_contains(df, BOOKS_AMT_CANDS)
            col_ref       = exact_or_contains(df, REF_CANDS)
            col_date      = exact_or_contains(df, DATE_CANDS)

            applied=False; pairs=0

            # Start with original values, modify only the match col
            if col_match in df_save.columns:
                col_values = df_save[col_match].copy()
            else:
                col_values = pd.Series([None]*len(df_save))

            if all([col_bank_code, col_bank_amt, col_books_amt, col_ref, col_date]):
                applied=True
                _date = normalize_date(pd.to_datetime(df[col_date], errors="coerce"))
                _bank_amt = to_number(pd.to_numeric(df[col_bank_amt], errors="coerce"))
                _books_amt = to_number(pd.to_numeric(df[col_books_amt], errors="coerce"))
                _bank_code = to_number(pd.to_numeric(df[col_bank_code], errors="coerce"))
                _ref = df[col_ref].astype(str).fillna("")

                # Build index for candidate "books" rows
                pos_index = {}
                for i in range(len(df)):
                    ok_ref = True
                    if require_ref_on_books:
                        ok_ref = ref_starts_with_ov_rc(_ref.iat[i])
                    if pd.notna(_books_amt.iat[i]) and _books_amt.iat[i] > 0 and pd.notna(_date.iat[i]) and ok_ref:
                        key = (_date.iat[i], round(float(_books_amt.iat[i]), 2))
                        pos_index.setdefault(key, []).append(i)

                for i in range(len(df)):
                    try:
                        if pd.notna(_bank_code.iat[i]) and int(_bank_code.iat[i]) in (175, 120)                            and pd.notna(_bank_amt.iat[i]) and _bank_amt.iat[i] < 0 and pd.notna(_date.iat[i]):
                            if ref_starts_with_ov_rc(_ref.iat[i]):
                                key = (_date.iat[i], round(abs(float(_bank_amt.iat[i])), 2))
                                cands = pos_index.get(key, [])
                                if len(cands) == 1:
                                    j = cands[0]
                                    col_values.iat[i] = 1
                                    if mark_both_rows:
                                        col_values.iat[j] = 1
                                    pairs += 1
                    except Exception:
                        continue

            df_to_write = df_save.copy()
            df_to_write[col_match] = col_values
            df_to_write.to_excel(writer, index=False, sheet_name=sh)

            summary_rows.append({
                "גיליון": sh,
                "בוצע": "כן" if applied else "לא – חסרות עמודות",
                "זוגות שסומנו": pairs,
                "עמודת התאמה": col_match,
                "קוד פעולה": col_bank_code,
                "סכום בדף": col_bank_amt,
                "סכום בספרים": col_books_amt,
                "אסמכתא 1": col_ref,
                "תאריך": col_date
            })

    # Add RTL to all sheets using openpyxl
    wb = load_workbook(io.BytesIO(writer_buf.getvalue()))
    for ws in wb.worksheets:
        ws.sheet_view.rightToLeft = True
    final_buf = io.BytesIO()
    wb.save(final_buf)
    final_buf.seek(0)

    return final_buf.getvalue(), pd.DataFrame(summary_rows)

uploaded = st.file_uploader("בחרי קובץ Excel (.xlsx)", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    mark_both_rows = st.checkbox("סמני גם את שורת הספרים (בנוסף לשורת הבנק)", value=True)
with col2:
    require_ref_on_books = st.checkbox("לדרוש OV/RC גם בשורת הספרים", value=True)

if uploaded:
    with st.spinner("מעבד…"):
        out_bytes, summary = process_workbook(uploaded.getvalue(), mark_both_rows, require_ref_on_books)
    st.success("מוכן!")
    st.dataframe(summary, use_container_width=True)

    st.download_button(
        "הורדת הקובץ המעודכן (RTL)",
        data=out_bytes,
        file_name="קובץ_מקורי_שינוי_רק_מס_התאמה_RTL.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.caption("© Rise Accounting – כלי עזר להתאמות")
