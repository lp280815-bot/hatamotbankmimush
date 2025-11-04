# -*- coding: utf-8 -*-
"""
התאמות בנק – 1 עד 12
1–4: OV/RC, הוראות קבע (VLOOKUP), העברות (#3), שיקים ספקים (#4)
5–10: עמלות, פאיימי, שיקים ממשמרת, הפק' שיק-שידור, הפק.שיק במכונה, קודים
11–12: placeholders (לא מסמנים עד שתקבלי לוגיקה)
תיקונים:
- כלל 3 מסמן גם צד ספרים לפי קובץ עזר (מס' תשלום) + סכומים לפי תאריך.
- גיליון 'הוראת קבע ספקים' צובע כתום שורות בלי מס' ספק.
- פריסת עמוד A4 לרוחב, Fit-to-width=1, שוליים, RTL.
"""

import io, os, re, json
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.worksheet.page import PageMargins

# ---------- UI ----------
st.set_page_config(page_title="התאמות בנק – 1 עד 12", page_icon="✅", layout="centered")
st.markdown("""
<style>
html, body, [class*="css"] { direction: rtl; text-align: right; }
.block-container { padding-top: 1rem; max-width: 1100px; }
</style>
""", unsafe_allow_html=True)
st.title("התאמות בנק – 1 עד 12")

# ---------- Constants ----------
STANDING_CODES = {469, 515}
OVRC_CODES     = {120, 175}
TRANSFER_CODE  = 485
TRANSFER_PHRASE = "העב' במקבץ-נט"
RULE4_CODE     = 493
RULE4_EPS      = 0.50

RULE5_CODES = {453, 472, 473, 124}
RULE6_COMPANY = 'פאיימי בע"מ'
RULE7_CODE = 143; RULE7_PHRASE = "שיקים ממשמרת"
RULE8_CODE = 191; RULE8_PHRASE = "הפק' שיק-שידור"
RULE9_CODE = 205; RULE9_PHRASE = "הפק.שיק במכונה"
RULE10_CODES = {191, 132, 396}

# placeholders 11–12
def rule11_placeholder(df, match_col, code_col, bamt_col, details_col): return df[match_col]
def rule12_placeholder(df, match_col, code_col, bamt_col, details_col): return df[match_col]

# ---------- Column maps ----------
MATCH_COLS = ["מס.התאמה","מס. התאמה","מס התאמה","מספר התאמה","התאמה"]
BANK_CODES = ["קוד פעולת בנק","קוד פעולה","קוד פעולת","Bank Code"]
BANK_AMTS  = ["סכום בדף","סכום דף","סכום בבנק","סכום תנועת בנק","Bank Amount"]
BOOKS_AMTS = ["סכום בספרים","סכום בספר","סכום ספרים","Books Amount"]
REF1S      = ["אסמכתא 1","אסמכתא1","אסמכתא","אסמכתה","Ref1"]
REF2S      = ["אסמכתא 2","אסמכתא2","אסמכתא-2","אסמכתה 2","Ref2"]
DATES      = ["תאריך מאזן","תאריך ערך","תאריך","Date"]
DETAILS    = ["פרטים","תיאור","שם ספק","Details","תאור"]

# aux (עזר להעברות)
AUX_DATE_KEYS = ["תאריך פריקה","תאריך","פריקה"]
AUX_AMT_KEYS  = ["אחרי ניכוי","אחרי","סכום"]
AUX_PAYNO_KEYS= ["מס' תשלום","מס תשלום","מספר תשלום"]

# ---------- Helpers ----------
def pick_col(df, names):
    for n in names:
        if n in df.columns: return n
    for n in names:
        for c in df.columns:
            if isinstance(c,str) and n in c: return c
    return None

def to_num(s):
    s = (s.astype(str).str.replace(",","",regex=False)
                  .str.replace("₪","",regex=False)
                  .str.replace("\u200f","",regex=False)
                  .str.replace("\u200e","",regex=False)
                  .str.strip())
    return pd.to_numeric(s, errors="coerce")

def norm_date(series):
    def f(x):
        if pd.isna(x): return pd.NaT
        if isinstance(x,(pd.Timestamp,datetime)): return pd.Timestamp(x.date())
        return pd.to_datetime(x, dayfirst=True, errors="coerce").normalize()
    return series.apply(f)

def ws_to_df(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows: return pd.DataFrame()
    header = [str(x) if x is not None else "" for x in rows[0]]
    data   = [list(r[:len(header)]) for r in rows[1:]]
    return pd.DataFrame(data, columns=header)

def only_digits(s): return re.sub(r"\D","", str(s)).lstrip("0") or "0"

# ---------- VLOOKUP store ----------
VK_FILE = "rules_store.json"
def vk_load():
    if os.path.exists(VK_FILE):
        try:
            with open(VK_FILE,"r",encoding="utf-8") as f: return json.load(f)
        except Exception: pass
    return {"name_map": {}, "amount_map": {}}
def vk_save(store):
    with open(VK_FILE,"w",encoding="utf-8") as f: json.dump(store,f,ensure_ascii=False,indent=2)

def build_vlookup_sheet(datasheet_df: pd.DataFrame) -> pd.DataFrame:
    store = vk_load()
    name_map   = {str(k): v for k,v in store.get("name_map",{}).items()}
    amount_map = {float(k): v for k,v in store.get("amount_map",{}).items()}

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
        for k,v in name_map.items():
            if k and k in s: return v
        try:
            key = round(abs(float(row["סכום"])),2)
            return amount_map.get(key,"")
        except Exception:
            return ""

    vk["מס' ספק"]   = vk.apply(pick_supplier, axis=1)
    vk["סכום חובה"] = vk["סכום"].apply(lambda x: abs(x) if pd.notna(x) and x<0 else 0.0)
    vk["סכום זכות"] = vk["סכום"].apply(lambda x: abs(x) if pd.notna(x) and x>0 else 0.0)
    return vk

# ---------- Rules 1–4 ----------
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

    if col_match not in out.columns: out[col_match]=0

    match = pd.to_numeric(out[col_match], errors="coerce").fillna(0).astype(int)
    code  = to_num(out[col_code]) if col_code else pd.Series([np.nan]*len(out))
    bamt  = to_num(out[col_bamt]) if col_bamt else pd.Series([np.nan]*len(out))
    aamt  = to_num(out[col_aamt]) if col_aamt else pd.Series([np.nan]*len(out))
    datev = norm_date(pd.to_datetime(out[col_date], errors="coerce")) if col_date else pd.Series([pd.NaT]*len(out))
    det   = out[col_det].astype(str).fillna("") if col_det else pd.Series([""]*len(out))
    ref1  = out[col_ref1].astype(str).fillna("") if col_ref1 else pd.Series([""]*len(out))
    ref2  = out[col_ref2].astype(str).fillna("") if col_ref2 else pd.Series([""]*len(out))

    # 1: OV/RC 1:1
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

    # 2: Standing orders
    for i in range(len(out)):
        if match.iat[i]==0 and pd.notna(code.iat[i]) and int(code.iat[i]) in STANDING_CODES:
            match.iat[i]=2

    # 3: Transfers – בנק + ספרים (ספרים לפי מס' תשלום מהעזר)
    # בנק: קוד 485, סכום חיובי, פרטים מכילים את הטקסט
    # ספרים: אסמכתא 1 ∈ קבוצת מס' תשלום לאותו יום
    # סימון לא דורס ערכים !=0
    # => סימון שני הצדדים = 3
    # (הטעינה של העזר נעשית בהמשך ב-process_workbook)
    out[col_match] = match
    return out

# ---------- Rules 5–12 ----------
def apply_rules_5_12(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    col_match = pick_col(out, MATCH_COLS) or out.columns[0]
    col_code  = pick_col(out, BANK_CODES)
    col_bamt  = pick_col(out, BANK_AMTS)
    col_det   = pick_col(out, DETAILS)

    if col_match not in out.columns: out[col_match]=0
    match = pd.to_numeric(out[col_match], errors="coerce").fillna(0).astype(int)
    code  = to_num(out[col_code]) if col_code else pd.Series([np.nan]*len(out))
    bamt  = to_num(out[col_bamt]) if col_bamt else pd.Series([np.nan]*len(out))
    det   = out[col_det].astype(str).fillna("") if col_det else pd.Series([""]*len(out))

    m5  = (match==0) & (code.isin(list(RULE5_CODES))) & (bamt>0) & (bamt<=500); match.loc[m5]=5
    m6  = (match==0) & (code==175) & (bamt<0) & (det.str.contains(RULE6_COMPANY, regex=False, na=False)); match.loc[m6]=6
    m7  = (match==0) & (code==RULE7_CODE) & (bamt<0) & (det==RULE7_PHRASE); match.loc[m7]=7
    m8  = (match==0) & (code==RULE8_CODE) & (bamt<0) & (det==RULE8_PHRASE); match.loc[m8]=8
    m9  = (match==0) & (code==RULE9_CODE) & (bamt<0) & (det==RULE9_PHRASE); match.loc[m9]=9
    m10 = (match==0) & (code.isin(list(RULE10_CODES))) & (bamt.notna()) & (bamt!=0); match.loc[m10]=10

    match = rule11_placeholder(out.assign(**{col_match:match}), col_match, pick_col(out,BANK_CODES), pick_col(out,BANK_AMTS), pick_col(out,DETAILS))
    match = rule12_placeholder(out.assign(**{col_match:match}), col_match, pick_col(out,BANK_CODES), pick_col(out,BANK_AMTS), pick_col(out,DETAILS))

    out[col_match]=match
    return out

# ---------- Styling ----------
def style_and_print(wb):
    for ws in wb.worksheets:
        ws.sheet_view.rightToLeft = True
        # A4 landscape, fit to width, margins
        ws.page_setup.paperSize = 9           # A4
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.6, bottom=0.6)
    # כתום לשורות בלי 'מס' ספק' בגיליון הוראת קבע
    if "הוראת קבע ספקים" in wb.sheetnames:
        ws = wb["הוראת קבע ספקים"]
        header = {cell.value: idx for idx, cell in enumerate(ws[1], start=1)}
        col_supplier = header.get("מס' ספק")
        if col_supplier:
            orange = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            for r in range(2, ws.max_row+1):
                v = ws.cell(row=r, column=col_supplier).value
                if v in ("", None):
                    for c in range(1, ws.max_column+1):
                        ws.cell(row=r, column=c).fill = orange
        # בולד לשורה האחרונה
        for cell in ws[ws.max_row]:
            cell.font = Font(bold=True)

# ---------- Processing ----------
def process_workbook(main_bytes: bytes, aux_bytes: bytes|None):
    # load main
    wb = load_workbook(io.BytesIO(main_bytes), data_only=True)
    ws = wb["DataSheet"] if "DataSheet" in wb.sheetnames else wb.worksheets[0]
    df = ws_to_df(ws)
    if df.empty: return None, None, None

    # 1–4
    df = apply_rules_1_4(df)

    # === Rule 3 (with AUX) — בנק + ספרים ===
    if aux_bytes is not None:
        aux_wb = load_workbook(io.BytesIO(aux_bytes), data_only=True)
        a_ws   = aux_wb.worksheets[0]
        a_df   = ws_to_df(a_ws)
        c_dt   = pick_col(a_df, AUX_DATE_KEYS)
        c_amt  = pick_col(a_df, AUX_AMT_KEYS)
        c_pay  = pick_col(a_df, AUX_PAYNO_KEYS)

        if c_dt and c_amt:
            a_dt  = norm_date(pd.to_datetime(a_df[c_dt], errors="coerce"))
            a_amt = pd.to_numeric(a_df[c_amt], errors="coerce").round(2)
            groups = (pd.DataFrame({"dt":a_dt, "amt":a_amt})
                        .dropna(subset=["dt"])
                        .groupby("dt")["amt"].sum().round(2)
                        .to_dict())
            pays = {}
            if c_pay:
                pays = (pd.DataFrame({"dt":a_dt, "pay":a_df[c_pay].astype(str).str.strip()})
                          .groupby("dt")["pay"]
                          .apply(lambda s: set(s.dropna().astype(str)))
                          .to_dict())

            # locate columns in df
            col_match = pick_col(df, MATCH_COLS) or df.columns[0]
            col_code  = pick_col(df, BANK_CODES)
            col_bamt  = pick_col(df, BANK_AMTS)
            col_det   = pick_col(df, DETAILS)
            col_ref1  = pick_col(df, REF1S)

            match = pd.to_numeric(df[col_match], errors="coerce").fillna(0).astype(int)
            code  = to_num(df[col_code])
            bamt  = to_num(df[col_bamt]).round(2)
            det   = df[col_det].astype(str).fillna("")
            # books side:
            ref1  = df[col_ref1].astype(str).fillna("") if col_ref1 else pd.Series([""]*len(df))
            datev = norm_date(pd.to_datetime(df[pick_col(df, DATES)], errors="coerce")) if pick_col(df, DATES) else pd.Series([pd.NaT]*len(df))

            bank_mask = (code==TRANSFER_CODE) & (bamt>0) & (det.str.contains(TRANSFER_PHRASE, na=False))
            for gdt, gsum in groups.items():
                # bank rows match by amount & date
                rows_bank = df.index[ bank_mask & (bamt.abs()==abs(gsum)) & (datev==gdt) & (match==0) ].tolist()
                for i in rows_bank: match.iat[i]=3
                # books rows match by pay numbers (if available)
                payset = pays.get(gdt, set())
                if payset and len(payset)>0 and col_ref1:
                    rows_books = df.index[(ref1.astype(str).isin(payset)) & (match==0)].tolist()
                    for j in rows_books: match.iat[j]=3
            df[col_match]=match

    # 5–12 (רק על 0)
    df = apply_rules_5_12(df)

    # build vlookup sheet
    vk_df = build_vlookup_sheet(df)

    # export with styling & print setup
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as wr:
        df.to_excel(wr, index=False, sheet_name="DataSheet")
        pd.DataFrame({"מס": pd.to_numeric(df[pick_col(df,MATCH_COLS) or df.columns[0]], errors="coerce")
                       .fillna(0).astype(int).value_counts().sort_index().index,
                      "כמות": pd.to_numeric(df[pick_col(df,MATCH_COLS) or df.columns[0]], errors="coerce")
                       .fillna(0).astype(int).value_counts().sort_index().values
                     }).to_excel(wr, index=False, sheet_name="סיכום")
        vk_df.to_excel(wr, index=False, sheet_name="הוראת קבע ספקים")
    wb_out = load_workbook(io.BytesIO(buffer.getvalue()))
    style_and_print(wb_out)
    final = io.BytesIO(); wb_out.save(final)
    return df, vk_df, final.getvalue()

# ---------- UI ----------
c1, c2 = st.columns([2,2])
main_file = c1.file_uploader("בחרי קובץ מקור – DataSheet בלבד", type=["xlsx"])
aux_file  = c2.file_uploader("⬆️ קובץ עזר להעברות (אופציונלי)", type=["xlsx"])
st.caption("VLOOKUP שומר מפות ב-rules_store.json (שם/סכום → מס' ספק).")

if st.button("הרצה 1–12"):
    if not main_file:
        st.error("נא להעלות קובץ מקור.")
    else:
        with st.spinner("מעבד..."):
            df_out, vk_out, out_bytes = process_workbook(main_file.read(), aux_file.read() if aux_file else None)
        if df_out is None:
            st.error("לא נמצאו נתונים.")
        else:
            st.success("מוכן!")
            col_match = pick_col(df_out, MATCH_COLS) or df_out.columns[0]
            cnt = pd.to_numeric(df_out[col_match], errors="coerce").fillna(0).astype(int).value_counts().sort_index()
            st.dataframe(pd.DataFrame({"מס":cnt.index,"כמות":cnt.values}), use_container_width=True)
            st.download_button("📥 הורד קובץ מעודכן", data=out_bytes,
                               file_name="התאמות_1_עד_12.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.divider()
st.subheader("🔎 VLOOKUP – הוראת קבע ספקים (עריכה ושמירה)")
store = vk_load()
with st.expander("מפות מיפוי (נשמר ל-rules_store.json)", expanded=False):
    t1, t2 = st.columns([2,1])
    nm = t1.text_input("מיפוי לפי 'פרטים' (contains)")
    sp = t2.text_input("מס' ספק")
    if st.button("➕ הוסף/עדכן לפי שם"):
        if nm and sp:
            store["name_map"][nm] = sp; vk_save(store); st.success("נשמר לפי שם.")
    t3, t4 = st.columns([1,1])
    amt = t3.number_input("מיפוי לפי סכום (ערך מוחלט)", step=0.01, format="%.2f")
    sp2 = t4.text_input("מס' ספק", key="vk2")
    if st.button("➕ הוסף/עדכן לפי סכום"):
        try:
            store["amount_map"][str(round(abs(float(amt)),2))] = sp2; vk_save(store); st.success("נשמר לפי סכום.")
        except Exception as e:
            st.error(str(e))
