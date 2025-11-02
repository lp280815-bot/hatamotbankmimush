# streamlit_app.py
# -*- coding: utf-8 -*-
import io
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="התאמות לקוחות – OV/RC", page_icon="✅", layout="centered")

# ===== RTL UI =====
st.markdown('''
    <style>
      html, body, [class*="css"] { direction: rtl; text-align: right; }
      .block-container { padding-top: 2rem; }
    </style>
''', unsafe_allow_html=True)

st.title("אפליקציה להתאמות לקוחות בבנק (175/120 + OV/RC רק בספרים)")

st.write("העלו קובץ Excel מקורי. האפליקציה תחזיר קובץ זהה **עם שינוי רק בעמודה _מס.התאמה_** לפי הכללים הבאים:")

st.markdown("""
1. שורות **בנק**: `קוד פעולת בנק ∈ {175, 120}` ו־`סכום בדף` **במינוס**.
2. שורות **ספרים**: `סכום בספרים` **בפלוס**, `אסמכתא 1` **מתחילה** ב־**OV** או **RC**.
3. תואמים לפי **אותו תאריך מאזן** ו־**אותו סכום אבסולוטי** (פלוס בספרים מול מינוס בבנק).
4. מסמנים `1` בעמודה **מס.התאמה** בשורת הבנק, ואפשר לבחור לסמן גם בשורת הספרים.
""")

# ---------- helpers ----------
def normalize_date(series: pd.Series) -> pd.Series:
    def to_date(x):
        if pd.isna(x):
            return pd.NaT
        if isinstance(x, (pd.Timestamp, datetime)):
            return pd.Timestamp(x.date())
        return pd.to_datetime(x, dayfirst=True, errors="coerce").normalize()
    return series.apply(to_date)

def to_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce")

def ref_starts_with_ov_rc(val) -> bool:
    t = (str(val) if val is not None else "").strip().upper()
    return t.startswith("OV") or t.startswith("RC")

def exact_or_contains(df: pd.DataFrame, names: list[str]) -> str | None:
    # העדפה לשם מדויק; אם לא, מכיל
    for n in names:
        if n in df.columns:
            return n
    for n in names:
        for c in df.columns:
            if isinstance(c, str) and n in c:
                return c
    return None

MATCH_COL_CANDS = ["מס.התאמה", "מס. התאמה", "מס התאמה", "מספר התאמה", "התאמה"]
BANK_CODE_CANDS = ["קוד פעולת בנק", "קוד פעולה", "קוד פעולת ה"]
BANK_AMT_CANDS  = ["סכום בדף", "סכום דף", "סכום בבנק", "סכום תנועת בנק"]
BOOKS_AMT_CANDS = ["סכום בספרים", "סכום בספר", "סכום ספרים"]
REF_CANDS       = ["אסמכתא 1", "אסמכתא1", "אסמכתא", "אסמכתה"]
DATE_CANDS      = ["תאריך מאזן", "תאריך ערך", "תאריך"]

# ---------- core processing ----------
@st.cache_data(show_spinner=False)
def process_workbook(file_bytes: bytes,
                     mark_both_rows: bool,
                     tie_breaker: str  # "nearest" or "first"
                     ) -> tuple[bytes, pd.DataFrame]:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_names = xls.sheet_names

    tmp = io.BytesIO()
    summary_rows = []

    with pd.ExcelWriter(tmp, engine="xlsxwriter") as writer:
        for sh in sheet_names:
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sh, dtype=object)
            df_save = df.copy()

            col_match = exact_or_contains(df, MATCH_COL_CANDS) or df.columns[0]
            col_bank_code = exact_or_contains(df, BANK_CODE_CANDS)
            col_bank_amt  = exact_or_contains(df, BANK_AMT_CANDS)
            col_books_amt = exact_or_contains(df, BOOKS_AMT_CANDS)
            col_ref       = exact_or_contains(df, REF_CANDS)
            col_date      = exact_or_contains(df, DATE_CANDS)

            applied = False
            pairs   = 0

            # נתחיל מערכי המקור בעמודת ההתאמה ונעדכן רק אותה
            match_values = df_save[col_match].copy() if col_match in df_save.columns else pd.Series([None]*len(df_save))

            if all([col_bank_code, col_bank_amt, col_books_amt, col_ref, col_date]):
                applied = True

                _date       = normalize_date(pd.to_datetime(df[col_date], errors="coerce"))
                _bank_amt   = to_number(pd.to_numeric(df[col_bank_amt], errors="coerce"))
                _books_amt  = to_number(pd.to_numeric(df[col_books_amt], errors="coerce"))
                _bank_code  = to_number(pd.to_numeric(df[col_bank_code], errors="coerce"))
                _ref        = df[col_ref].astype(str).fillna("")

                # בונים רשימת מועמדי ספרים: פלוס + OV/RC + תאריך קיים
                books_candidates = []
                for i in range(len(df)):
                    if pd.notna(_books_amt.iat[i]) and _books_amt.iat[i] > 0 \
                       and pd.notna(_date.iat[i]) and ref_starts_with_ov_rc(_ref.iat[i]):
                        books_candidates.append(i)

                used_books = set()

                # עוברים על שורות בנק: קוד 175/120 + מינוס + תאריך קיים (ללא דרישת OV/RC)
                for i in range(len(df)):
                    if pd.notna(_bank_code.iat[i]) and int(_bank_code.iat[i]) in (175, 120) \
                       and pd.notna(_bank_amt.iat[i]) and _bank_amt.iat[i] < 0 \
                       and pd.notna(_date.iat[i]):

                        target_date = _date.iat[i]
                        target_amt  = round(abs(float(_bank_amt.iat[i])), 2)

                        # מציאת מועמדים תואמים בתאריך/סכום
                        cands = [j for j in books_candidates
                                 if j not in used_books
                                 and _date.iat[j] == target_date
                                 and round(float(_books_amt.iat[j]), 2) == target_amt]

                        chosen = None
                        if len(cands) == 1:
                            chosen = cands[0]
                        elif len(cands) > 1:
                            if tie_breaker == "nearest":
                                # הקרובה ביותר לפי אינדקס
                                chosen = min(cands, key=lambda j: abs(j - i))
                            else:  # "first"
                                chosen = cands[0]

                        if chosen is not None:
                            match_values.iat[i] = 1
                            if mark_both_rows:
                                match_values.iat[chosen] = 1
                            used_books.add(chosen)
                            pairs += 1

            # שמירה: רק עמודת ההתאמה משתנה
            df_out = df_save.copy()
            df_out[col_match] = match_values
            df_out.to_excel(writer, index=False, sheet_name=sh)

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

    # הופך את כל הגיליונות ל-RTL
    wb = load_workbook(io.BytesIO(tmp.getvalue()))
    for ws in wb.worksheets:
        ws.sheet_view.rightToLeft = True
    final = io.BytesIO()
    wb.save(final)
    final.seek(0)

    return final.getvalue(), pd.DataFrame(summary_rows)

# ---------- UI ----------
uploaded = st.file_uploader("בחרי קובץ Excel (.xlsx)", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    mark_both_rows = st.checkbox("לסמן גם את שורת הספרים", value=True)
with col2:
    tie_breaker = st.selectbox("אם יש כמה מועמדים זהים:", ["הקרובה ביותר", "הראשונה"], index=0)
    tie_breaker_key = "nearest" if tie_breaker == "הקרובה ביותר" else "first"

if uploaded:
    with st.spinner("מעבדת…"):
        out_bytes, summary = process_workbook(uploaded.getvalue(),
                                              mark_both_rows=mark_both_rows,
                                              tie_breaker=tie_breaker_key)
    st.success("מוכן! הקובץ זהה למקור עם שינוי רק בעמודה 'מס.התאמה'.")
    st.dataframe(summary, use_container_width=True)

    st.download_button(
        "הורדת הקובץ המעודכן (RTL)",
        data=out_bytes,
        file_name="קובץ_מקורי_שינוי_רק_מס_התאמה_RTL.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.caption("© Rise Accounting – כלי עזר להתאמות")
