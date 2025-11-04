# -*- coding: utf-8 -*-
"""
התאמות בנק – מימוש מלא
1–4: OV/RC, הוראות קבע (כולל VLOOKUP), העברות, שיקים ספקים
5–10: עמלות, פאיימי, שיקים ממשמרת, הפק' שיק-שידור, הפק.שיק במכונה (קוד 205), וקודים 191/132/396
11–12: מחזיקי מקום – לא מסמנים דבר עד קבלת לוגיקה
כללים 5–12 מסמנים אך ורק שורות שמספר ההתאמה בהן 0 (כך שלא דורכים על 1–4).
"""

import io, os, re, json
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# -------------------------------------------------------
# הגדרות UI (RTL)
# -------------------------------------------------------
st.set_page_config(page_title="התאמות בנק – 1 עד 12", page_icon="✅", layout="centered")
st.markdown("""
<style>
html, body, [class*="css"] { direction: rtl; text-align: right; }
.block-container { padding-top: 1rem; max-width: 1100px; }
</style>
""", unsafe_allow_html=True)
st.title("התאמות בנק – 1 עד 12")

# -------------------------------------------------------
# קבועים – לוגיקות קיימות
# -------------------------------------------------------
STANDING_CODES = {469, 515}                                     # כלל 2
OVRC_CODES     = {120, 175}                                     # כלל 1
TRANSFER_CODE  = 485                                             # כלל 3
TRANSFER_PHRASE = "העב' במקבץ-נט"                                # כלל 3
RULE4_CODE     = 493                                             # כלל 4 (שיקים ספקים)
RULE4_EPS      = 0.50

# כללים 5–10 (מאושרים)
RULE5_CODES = {453, 472, 473, 124}                               # עמלות – חיובי ועד 500
RULE6_COMPANY = 'פאיימי בע"מ'                                    # 175, שלילי, פרטים מכיל בדיוק
RULE7_CODE = 143; RULE7_PHRASE = "שיקים ממשמרת"                  # שלילי, פרטים בדיוק
RULE8_CODE = 191; RULE8_PHRASE = "הפק' שיק-שידור"                 # שלילי, פרטים בדיוק
RULE9_CODE = 205; RULE9_PHRASE = "הפק.שיק במכונה"                # שלילי, פרטים בדיוק
RULE10_CODES = {191, 132, 396}                                   # כל סכום ≠ 0

# כללים 11–12 – מחזיק מקום (עד שתישלח לוגיקה)
def rule11_placeholder(df, match_col, code_col, bamt_col, details_col):
    return df[match_col]  # לא מסמן כלום

def rule12_placeholder(df, match_col, code_col, bamt_col, details_col):
    return df[match_col]  # לא מסמן כלום

# -------------------------------------------------------
# זיהוי עמודות ונרמול
# -------------------------------------------------------
MATCH_COLS = ["מס.התאמה","מס. התאמה","מס התאמה","מספר התאמה","התאמה"]
BANK_CODES = ["קוד פעולת בנק","קוד פעולה","קוד פעולת","Bank Code"]
BANK_AMTS  = ["סכום בדף","סכום דף","סכום בבנק","סכום תנועת בנק","Bank Amount"]
BOOKS_AMTS = ["סכום בספרים","סכום בספר","סכום ספרים","Books Amount"]
REF1S      = ["אסמכתא 1","אסמכתא1","אסמכתא","אסמכתה","Ref1"]
REF2S      = ["אסמכתא 2","אסמכתא2","אסמכתא-2","אסמכתה 2","Ref2"]
DATES      = ["תאריך מאזן","תאריך ערך","תאריך","Date"]
DETAILS    = ["פרטים","תיאור","שם ספק","Details","תאור"]

def pick_col(df, names):
    for n in names:
        if n in df.columns: return n
    for n in names:
        for c in df.columns:
            if isinstance(c,str) and n in c: return c
    return None

def to_num(s):
    s = (s.astype(str)
         .str.replace(",","",regex=False)
         .str.replace("₪","",regex=False)
         .str.replace("\u200f","",regex=False)
         .str.replace("\u200e","",regex=False)
         .str.strip())
    return pd.to_numeric(s, errors="coerce")

def norm_date(series):
    def f(x):
        if pd.isna(x): return pd.NaT
        if isinstance(x, (pd.Timestamp, datetime)): return pd.Timestamp(x.date())
        return pd.to_datetime(x, dayfirst=True, errors="coerce").normalize()
    return series.apply(f)

def ws_to_df(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows: return pd.DataFrame()
    header = [str(x) if x is not None else "" for x in rows[0]]
    data   = [list(r[:len(header)]) for r in rows[1:]]
    return pd.DataFrame(data, columns=header)

def only_digits(s):
    return re.sub(r"\D","", str(s)).lstrip("0") or "0"

# -------------------------------------------------------
# VLOOKUP – הוראת קבע ספקים (שמירה וייבוא)
# -------------------------------------------------------
VK_FILE = "rules_store.json"

def vk_load():
    if os.path.exists(VK_FILE):
        try:
            with open(VK_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"name_map": {}, "amount_map": {}}

def vk_save(store):
    with open(VK_FILE, "w", encoding="utf-8") as f:
        json.dump(store, f, ensure_ascii=False, indent=2)

def build_vlookup_sheet(datasheet_df: pd.DataFrame) -> pd.DataFrame:
    store = vk_load()
    name_map   = {str(k): v for k, v in store.get("name_map", {}).items()}
    amount_map = {float(k): v for k, v in store.get("amount_map", {}).items()}

    col_match = pick_col(datasheet_df, MATCH_COLS) or datasheet_df.columns[0]
    col_bamt  = pick_col(datasheet_df, BANK_AMTS)
    col_det   = pick_col(datasheet_df, DETAILS)

    match = pd.to_numeric(datasheet_df[col_match], errors="coerce").fillna(0).astype(int)
    bamt  = to_num(datasheet_df[col_bamt]) if col_bamt else pd.Series([np.nan]*len(datasheet_df))
    det   = datasheet_df[col_det].astype(str).fillna("")

    vk = datasheet_df.loc[match==2, [col_det, col_bamt]].rename(columns={col_det:"פרטים", col_bamt:"סכום"}).copy()
    if vk.empty:
        return pd.DataFrame(columns=["פרטים","סכום","מס' ספק","סכום חובה","סכום זכות"])

    def pick_supplier(row):
        s = str(row["פרטים"])
        # עדיפות התאמת-שם (contains)
        for k,v in name_map.items():
            if k and k in s:
                return v
        # אם לא, לפי סכום אבסולוטי
        try:
            key = round(abs(float(row["סכום"])),2)
            return amount_map.get(key, "")
        except Exception:
            return ""

    vk["מס' ספק"] = vk.apply(pick_supplier, axis=1)
    vk["סכום חובה"] = vk["סכום"].apply(lambda x: abs(x) if pd.notna(x) and x<0 else 0.0)
    vk["סכום זכות"] = vk["סכום"].apply(lambda x: abs(x) if pd.notna(x) and x>0 else 0.0)
    return vk

# -------------------------------------------------------
# לוגיקה 1–4 (כפי שסיכמנו; 5–12 לעולם לא דורסים אותם)
# -------------------------------------------------------
def apply_rules_1_4(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    col_match = pick_col(out, MATCH_COLS) or out.columns[0]
    col_code  = pick_col(out, BANK_CODES)
    col_bamt  = pick_col(out, BANK_AMTS)
    col_aamt  = pick_col(out, BOOKS_AMTS)
    col_ref1  = pick_col(out, REF1S)
    col_ref2  = pick_col(out, REF2S)
    col_date  = pick_col(out, DATES)
    col_det   = pick_col(out, DETAILS)

    if col_match not in out.columns:
        out[col_match] = 0

    match = pd.to_numeric(out[col_match], errors="coerce").fillna(0).astype(int)
    code  = to_num(out[col_code]) if col_code else pd.Series([np.nan]*len(out))
    bamt  = to_num(out[col_bamt]) if col_bamt else pd.Series([np.nan]*len(out))
    aamt  = to_num(out[col_aamt]) if col_aamt else pd.Series([np.nan]*len(out))
    datev = norm_date(pd.to_datetime(out[col_date], errors="coerce")) if col_date else pd.Series([pd.NaT]*len(out))
    det   = out[col_det].astype(str).fillna("") if col_det else pd.Series([""]*len(out))
    ref1  = out[col_ref1].astype(str).fillna("") if col_ref1 else pd.Series([""]*len(out))
    ref2  = out[col_ref2].astype(str).fillna("") if col_ref2 else pd.Series([""]*len(out))

    # ---- כלל 1: OV/RC התאמה 1:1 לפי תאריך וסכום ----
    bank_keys, books_keys = {}, {}
    for i in range(len(out)):
        if match.iat[i]!=0: continue
        if pd.notna(code.iat[i]) and int(code.iat[i]) in OVRC_CODES and pd.notna(bamt.iat[i]) and bamt.iat[i] < 0 and pd.notna(datev.iat[i]):
            k=(round(abs(float(bamt.iat[i])),2), datev.iat[i])
            bank_keys.setdefault(k, []).append(i)
    for j in range(len(out)):
        if match.iat[j]!=0: continue
        if pd.notna(aamt.iat[j]) and aamt.iat[j]>0 and pd.notna(datev.iat[j]) and str(ref1.iat[j]).upper().startswith(("OV","RC")):
            k=(round(abs(float(aamt.iat[j])),2), datev.iat[j])
            books_keys.setdefault(k, []).append(j)
    for k, bidx in bank_keys.items():
        if len(bidx)==1 and len(books_keys.get(k,[]))==1:
            i=bidx[0]; j=books_keys[k][0]
            if match.iat[i]==0 and match.iat[j]==0:
                match.iat[i]=1; match.iat[j]=1

    # ---- כלל 2: הוראות קבע (סימון בנק; VLOOKUP ייעשה בגיליון ייעודי) ----
    for i in range(len(out)):
        if match.iat[i]==0 and pd.notna(code.iat[i]) and int(code.iat[i]) in STANDING_CODES:
            match.iat[i]=2

    # ---- כלל 3: העברות (ע"פ טקסט וקוד; מיזוג מול קובץ עזר יתבצע מחוץ לפונקציה במידת הצורך) ----
    for i in range(len(out)):
        if match.iat[i]==0 and pd.notna(code.iat[i]) and int(code.iat[i])==TRANSFER_CODE and det.iat[i] and TRANSFER_PHRASE in det.iat[i]:
            match.iat[i]=3

    # ---- כלל 4: שיקים ספקים 493 – מיפוי Ref1 (בנק) ↔ Ref2 (ספרים), סכומים בטולרנס ----
    bank_idx = [i for i in range(len(out)) if match.iat[i]==0 and pd.notna(code.iat[i]) and int(code.iat[i])==RULE4_CODE and str(ref1.iat[i]).strip() and pd.notna(bamt.iat[i])]
    books_idx= [j for j in range(len(out)) if match.iat[j]==0 and str(ref1.iat[j]).upper().startswith("CH") and str(ref2.iat[j]).strip() and pd.notna(aamt.iat[j])]
    used=set()
    for i in bank_idx:
        ref_b = only_digits(ref1.iat[i]); ab = abs(float(bamt.iat[i]))
        for j in books_idx:
            if j in used or match.iat[j]!=0: continue
            if only_digits(ref2.iat[j]) != ref_b: continue
            aj = abs(float(aamt.iat[j]))
            if abs(aj - ab) <= RULE4_EPS:
                match.iat[i]=4; match.iat[j]=4; used.add(j); break

    out[col_match] = match
    return out

# -------------------------------------------------------
# לוגיקה 5–12 (לא דורסים 1–4)
# -------------------------------------------------------
def apply_rules_5_12(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    col_match = pick_col(out, MATCH_COLS) or out.columns[0]
    col_code  = pick_col(out, BANK_CODES)
    col_bamt  = pick_col(out, BANK_AMTS)
    col_det   = pick_col(out, DETAILS)

    if col_match not in out.columns:
        out[col_match] = 0

    match = pd.to_numeric(out[col_match], errors="coerce").fillna(0).astype(int)
    code  = to_num(out[col_code]) if col_code else pd.Series([np.nan]*len(out))
    bamt  = to_num(out[col_bamt]) if col_bamt else pd.Series([np.nan]*len(out))
    det   = out[col_det].astype(str).fillna("") if col_det else pd.Series([""]*len(out))

    # 5: עמלות – חיובי ועד 500
    m5 = (match==0) & (code.isin(list(RULE5_CODES))) & (bamt>0) & (bamt<=500); match.loc[m5]=5

    # 6: פאיימי בע"מ – קוד 175, שלילי, פרטים בדיוק
    m6 = (match==0) & (code==175) & (bamt<0) & (det.str.contains(RULE6_COMPANY, regex=False, na=False)); match.loc[m6]=6

    # 7: שיקים ממשמרת – קוד 143, שלילי, פרטים בדיוק
    m7 = (match==0) & (code==RULE7_CODE) & (bamt<0) & (det==RULE7_PHRASE); match.loc[m7]=7

    # 8: הפק' שיק-שידור – קוד 191, שלילי, פרטים בדיוק
    m8 = (match==0) & (code==RULE8_CODE) & (bamt<0) & (det==RULE8_PHRASE); match.loc[m8]=8

    # 9: הפק.שיק במכונה – קוד 205, שלילי, פרטים בדיוק
    m9 = (match==0) & (code==RULE9_CODE) & (bamt<0) & (det==RULE9_PHRASE); match.loc[m9]=9

    # 10: הפק.שיק במכונה – קודים 191/132/396, כל סכום ≠ 0
    m10 = (match==0) & (code.isin(list(RULE10_CODES))) & (bamt.notna()) & (bamt!=0); match.loc[m10]=10

    # 11–12: מחזיקי מקום (לוגיקה תתווסף כשנשלחת)
    match = rule11_placeholder(out.assign(**{col_match:match}), col_match, col_code, col_bamt, pick_col(out, DETAILS) or "")
    match = rule12_placeholder(out.assign(**{col_match:match}), col_match, col_code, col_bamt, pick_col(out, DETAILS) or "")

    out[col_match] = match
    return out

# -------------------------------------------------------
# עיבוד קובץ מקור – DataSheet אחד (אפשר להרחיב לגיליונות נוספים)
# -------------------------------------------------------
def process_workbook(main_bytes: bytes):
    wb = load_workbook(io.BytesIO(main_bytes), data_only=True)
    ws = wb["DataSheet"] if "DataSheet" in wb.sheetnames else wb.worksheets[0]
    df = ws_to_df(ws)
    if df.empty: return None, None, None

    # 1–4
    df = apply_rules_1_4(df)
    # 5–12 (רק על 0)
    df = apply_rules_5_12(df)

    # סיכום
    col_match = pick_col(df, MATCH_COLS) or df.columns[0]
    counts = pd.to_numeric(df[col_match], errors="coerce").fillna(0).astype(int).value_counts().sort_index()

    # גיליון הוראת קבע ספקים (מבוסס על כללי 2)
    vk_df = build_vlookup_sheet(df)

    # יצוא
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as wr:
        df.to_excel(wr, index=False, sheet_name="DataSheet")
        counts_df = pd.DataFrame({"מס": counts.index, "כמות": counts.values})
        counts_df.to_excel(wr, index=False, sheet_name="סיכום")
        vk_df.to_excel(wr, index=False, sheet_name="הוראת קבע ספקים")
    return df, vk_df, out.getvalue()

# -------------------------------------------------------
# UI – העלאות והרצה
# -------------------------------------------------------
c1, c2 = st.columns([2,2])
main_file = c1.file_uploader("בחרי קובץ מקור – DataSheet בלבד", type=["xlsx"])
# קובץ עזר – נשמר לשימוש עתידי/ייבוא מפות; אינו חובה להרצה
aux_file  = c2.file_uploader("⬆️ קובץ עזר (אופציונלי) – לשימוש בעדכוני מפות VLOOKUP/העברות", type=["xlsx"])

st.caption("טיפ: מפות VLOOKUP נשמרות/נקראות מקובץ rules_store.json בתיקיית האפליקציה.")

if st.button("הרצה 1–12"):
    if not main_file:
        st.error("נא להעלות קובץ מקור.")
    else:
        with st.spinner("מעבד..."):
            df_out, vk_out, out_bytes = process_workbook(main_file.read())
        if df_out is None:
            st.error("לא נמצאו נתונים בגיליון.")
        else:
            st.success("מוכן!")
            st.subheader("סיכום התאמות")
            col_match = pick_col(df_out, MATCH_COLS) or df_out.columns[0]
            cnt = pd.to_numeric(df_out[col_match], errors="coerce").fillna(0).astype(int).value_counts().sort_index()
            st.dataframe(pd.DataFrame({"מס": cnt.index, "כמות": cnt.values}), use_container_width=True)
            st.download_button("📥 הורד קובץ מעודכן", data=out_bytes,
                               file_name="התאמות_1_עד_12.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.divider()
st.subheader("🔎 VLOOKUP – הוראת קבע ספקים (עריכה ושמירה)")
store = vk_load()
with st.expander("הגדרות מיפוי (נשמר ל-rules_store.json)", expanded=False):
    t1, t2 = st.columns([2,1])
    nm = t1.text_input("מיפוי לפי פרטים (מחרוזת שמופיעה בעמודת 'פרטים')")
    sp = t2.text_input("מס' ספק")
    if st.button("➕ הוסף/עדכן לפי שם"):
        if nm and sp:
            store["name_map"][nm] = sp
            vk_save(store)
            st.success("נשמר לפי שם.")
    t3, t4 = st.columns([1,1])
    amt = t3.number_input("מיפוי לפי סכום (ערך מוחלט)", step=0.01, format="%.2f")
    sp2 = t4.text_input("מס' ספק", key="vk2")
    if st.button("➕ הוסף/עדכן לפי סכום"):
        try:
            store["amount_map"][str(round(abs(float(amt)),2))] = sp2
            vk_save(store)
            st.success("נשמר לפי סכום.")
        except Exception as e:
            st.error(str(e))
    st.caption("אפשר לייצא/לייבא את הקובץ JSON מהשרת (rules_store.json).")

# כפתור בניית גיליון VLOOKUP נפרד מההרצה הראשית (אם רוצים לבדוק בלבד)
if st.button("בני גיליון 'הוראת קבע ספקים' מקובץ מקור (ללא הרצה)"):
    if not main_file:
        st.error("נא להעלות קובץ מקור.")
    else:
        wb_tmp = load_workbook(main_file, data_only=True)
        ws_tmp = wb_tmp["DataSheet"] if "DataSheet" in wb_tmp.sheetnames else wb_tmp.worksheets[0]
        df_tmp = ws_to_df(ws_tmp)
        vk_tmp = build_vlookup_sheet(df_tmp)
        st.dataframe(vk_tmp, use_container_width=True, height=320)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as wr:
            vk_tmp.to_excel(wr, index=False, sheet_name="הוראת קבע ספקים")
        st.download_button("📥 הורד גיליון הוראת קבע ספקים", data=buf.getvalue(),
                           file_name="הוראת_קבע_ספקים.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
