# -*- coding: utf-8 -*-
"""
התאמות בנק – מימוש
כללים 1–4: OV/RC + הוראות קבע VLOOKUP + העברות + שיקים ספקים
כללים 5–10: עמלות, פאיימי, שיקים ממשמרת, הפק' שיק-שידור,
            הפק.שיק במכונה, קודים — ללא דריסה של 1–4
"""

import io, re, os, json
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# -------------------------------------------------------
# UI RTL
# -------------------------------------------------------
st.set_page_config(page_title="התאמות בנק 1–10", page_icon="✅", layout="centered")

st.markdown("""
<style>
html, body, [class*="css"] { direction: rtl; text-align: right; }
.block-container { padding-top: 1rem; }
</style>
""", unsafe_allow_html=True)

st.title("התאמות בנק – 1 עד 10")

# -------------------------------------------------------
# קבועים – לוגיקות
# -------------------------------------------------------
STANDING_CODES = {469, 515}
OVRC_CODES = {120, 175}
TRANSFER_CODE = 485
TRANSFER_PHRASE = "העב' במקבץ-נט"
RULE4_CODE = 493
RULE4_EPS = 0.50

# כלל 5–10
RULE5_CODES = {453, 472, 473, 124}
RULE6_COMPANY = 'פאיימי בע"מ'
RULE7_CODE = 143; RULE7_PHRASE = "שיקים ממשמרת"
RULE8_CODE = 191; RULE8_PHRASE = "הפק' שיק-שידור"
RULE9_CODE = 205; RULE9_PHRASE = "הפק.שיק במכונה"
RULE10_CODES = {191, 132, 396}

# -------------------------------------------------------
# פונקציות עזר
# -------------------------------------------------------
def normalize_date(series):
    def f(x):
        if pd.isna(x): return pd.NaT
        if isinstance(x,(pd.Timestamp,datetime)): return pd.Timestamp(x.date())
        return pd.to_datetime(x,dayfirst=True,errors="coerce").normalize()
    return series.apply(f)

def to_num(s):
    s = (s.astype(str)
           .str.replace(",","",regex=False)
           .str.replace("₪","",regex=False)
           .str.replace("\u200f","",regex=False)
           .str.replace("\u200e","",regex=False)
           .str.strip())
    return pd.to_numeric(s,errors="coerce")

def ref_ovrc(v): 
    if not isinstance(v,str): return False
    t=v.strip().upper()
    return (t.startswith("OV") or t.startswith("RC"))

def exact_col(df, names):
    for n in names:
        if n in df.columns: return n
    for n in names:
        for c in df.columns:
            if isinstance(c,str) and n in c:
                return c
    return None

def ws_to_df(ws):
    rows=list(ws.iter_rows(values_only=True))
    header=[str(x) if x else "" for x in rows[0]]
    data=[list(r[:len(header)]) for r in rows[1:]]
    return pd.DataFrame(data, columns=header)

# שמות עמודות אפשריים
MATCH_COLS=["מס.התאמה","מס. התאמה","מס התאמה","מספר התאמה","התאמה"]
BANK_CODES=["קוד פעולת בנק","קוד פעולה","קוד פעולת"]
BANK_AMTS=["סכום בדף","סכום דף","סכום בבנק"]
BOOKS_AMTS=["סכום בספרים","סכום בספר"]
REF1S=["אסמכתא 1","אסמכתא1","אסמכתא","אסמכתה"]
REF2S=["אסמכתא 2","אסמכתא2","אסמכתא-2","אסמכתה 2"]
DATES=["תאריך מאזן","תאריך ערך","תאריך"]
DETAILS=["פרטים","תיאור","שם ספק"]

# -------------------------------------------------------
# לוגיקה
# -------------------------------------------------------
def process_file(file):
    wb=load_workbook(file, data_only=True)
    ws=wb["DataSheet"] if "DataSheet" in wb.sheetnames else wb.worksheets[0]

    df=ws_to_df(ws)
    if df.empty:
        st.error("אין נתונים")
        return None,None

    # איתור עמודות
    col_match = exact_col(df, MATCH_COLS) or df.columns[0]
    col_code  = exact_col(df, BANK_CODES)
    col_bamt  = exact_col(df, BANK_AMTS)
    col_aamt  = exact_col(df, BOOKS_AMTS)
    col_ref1  = exact_col(df, REF1S)
    col_ref2  = exact_col(df, REF2S)
    col_date  = exact_col(df, DATES)
    col_det   = exact_col(df, DETAILS)

    match = df[col_match].fillna(0).astype(int)
    code  = to_num(df[col_code]) if col_code else pd.Series([np.nan]*len(df))
    bamt  = to_num(df[col_bamt])
    aamt  = to_num(df[col_aamt]) if col_aamt else pd.Series([np.nan]*len(df))
    datev = normalize_date(pd.to_datetime(df[col_date],errors="coerce")) if col_date else pd.Series([pd.NaT]*len(df))
    det   = df[col_det].astype(str).fillna("")

    # ---------------- התאמה 1 OV/RC ----------------
    if col_ref1:
        for i in range(len(df)):
            if match.iat[i]!=0: continue
            if ref_ovrc(df[col_ref1].iat[i]):
                for j in range(len(df)):
                    if i==j: continue
                    if match.iat[j]!=0: continue
                    if not ref_ovrc(df[col_ref1].iat[j]): continue
                    if datev.iat[i]==datev.iat[j] and abs(bamt.iat[i])==abs(aamt.iat[j]):
                        match.iat[i]=match.iat[j]=1
                        break

    # ---------------- התאמה 2 הוראות קבע ----------------
    for i in range(len(df)):
        if match.iat[i]==0 and code.iat[i] in STANDING_CODES:
            match.iat[i]=2

    # ---------------- התאמה 3 העברות (פשטני) ----------------
    for i in range(len(df)):
        if match.iat[i]==0 and code.iat[i]==TRANSFER_CODE and TRANSFER_PHRASE in det.iat[i]:
            match.iat[i]=3

    # ---------------- התאמה 4 שיקים ספקים 493 ----------------
    for i in range(len(df)):
        if match.iat[i]!=0: continue
        if code.iat[i]==RULE4_CODE and col_ref2:
            for j in range(len(df)):
                if i==j: continue
                if match.iat[j]!=0: continue
                if df[col_ref2].iat[j]==df[col_ref1].iat[i] and abs(bamt.iat[i])==abs(aamt.iat[j]):
                    match.iat[i]=match.iat[j]=4
                    break

    # ---------------- התאמה 5 עמלות ----------------
    mask5 = (match==0) & (code.isin(list(RULE5_CODES))) & (bamt>0) & (bamt<=500)
    match.loc[mask5] = 5

    # ---------------- התאמה 6 פאיימי ----------------
    mask6 = (match==0) & (code==175) & (bamt<0) & (det.str.contains(RULE6_COMPANY,regex=False))
    match.loc[mask6] = 6

    # ---------------- התאמה 7 שיקים ממשמרת ----------------
    mask7 = (match==0) & (code==RULE7_CODE) & (bamt<0) & (det==RULE7_PHRASE)
    match.loc[mask7] = 7

    # ---------------- התאמה 8 הפק' שיק-שידור ----------------
    mask8 = (match==0) & (code==RULE8_CODE) & (bamt<0) & (det==RULE8_PHRASE)
    match.loc[mask8] = 8

    # ---------------- התאמה 9 הפק.שיק במכונה ----------------
    mask9 = (match==0) & (code==RULE9_CODE) & (bamt<0) & (det==RULE9_PHRASE)
    match.loc[mask9] = 9

    # ---------------- התאמה 10 קודים נוספים ----------------
    mask10 = (match==0) & (code.isin(list(RULE10_CODES))) & (bamt!=0)
    match.loc[mask10] = 10

    df[col_match]=match
    counts=match.value_counts().sort_index()

    # יצוא לקובץ הורדה
    output=io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as wr:
        df.to_excel(wr,index=False,sheet_name="DataSheet")
        pd.DataFrame({"מס":counts.index,"כמות":counts.values}).to_excel(wr,index=False,sheet_name="סיכום")
    return output.getvalue(), counts


# -------------------------------------------------------
# UI הפעלה
# -------------------------------------------------------
file=st.file_uploader("בחרי קובץ מקור – DataSheet בלבד", type=["xlsx"])

if st.button("הרצה 1–10"):
    if file:
        with st.spinner("עובד..."):
            out, cnt = process_file(file)
        st.success("✅ מוכן!")
        st.dataframe(pd.DataFrame({"מס":cnt.index,"כמות":cnt.values}), use_container_width=True)
        st.download_button("📥 הורדה", data=out,
                           file_name="התאמות_1_10.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.error("נא להעלות קובץ")
