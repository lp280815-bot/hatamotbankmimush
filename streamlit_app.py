# streamlit_app.py
# -*- coding: utf-8 -*-
import io, re
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font

# ---------- UI ----------
st.set_page_config(page_title="התאמות לקוחות – OV/RC + הוראות קבע", page_icon="✅", layout="centered")
st.markdown("""
<style>
  html, body, [class*="css"] { direction: rtl; text-align: right; }
  .block-container { padding-top: 1.3rem; }
</style>
""", unsafe_allow_html=True)

st.title("התאמות לקוחות – OV/RC + הוראות קבע (כללי VLOOKUP קבועים)")
st.caption("האפליקציה משתמשת בכללי VLOOKUP קבועים שמוטמעים בקוד. אין צורך להעלות קובץ כללים.")

# ---------- כללי VLOOKUP קבועים ----------
DEFAULT_NAME_MAP = {
    "עיריית אשדוד": 30056,
    "ישראכרט מור": 34002,
}
DEFAULT_AMOUNT_MAP = {
    8520.0: 30247,
    10307.3: 30038,
}

def rules_excel_bytes():
    # מאפשר להוריד את טבלת הכללים הקבועה (רק לנוחות)
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        pd.DataFrame(
            {"by_name": list(DEFAULT_NAME_MAP.keys()),
             "מס' ספק": list(DEFAULT_NAME_MAP.values())}
        ).to_excel(w, index=False, sheet_name="by_name")
        pd.DataFrame(
            {"סכום": list(DEFAULT_AMOUNT_MAP.keys()),
             "מס' ספק": list(DEFAULT_AMOUNT_MAP.values())}
        ).to_excel(w, index=False, sheet_name="by_amount")
    return out.getvalue()

# ---------- Helpers (כמו בגירסה הקודמת) ----------
MATCH_COL_CANDS = ["מס.התאמה","מס. התאמה","מס התאמה","מספר התאמה","התאמה"]
BANK_CODE_CANDS = ["קוד פעולת בנק","קוד פעולה","קוד פעולת"]
BANK_AMT_CANDS  = ["סכום בדף","סכום דף","סכום בבנק","סכום תנועת בנק"]
BOOKS_AMT_CANDS = ["סכום בספרים","סכום בספר","סכום ספרים"]
REF_CANDS       = ["אסמכתא 1","אסמכתא1","אסמכתא","אסמכתה"]
DATE_CANDS      = ["תאריך מאזן","תאריך ערך","תאריך"]
DETAILS_CANDS   = ["פרטים","תיאור","שם ספק"]

def ws_to_df(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows: return pd.DataFrame()
    header = None; start = 0
    for i, r in enumerate(rows):
        if any(x is not None for x in r):
            header = [str(x).strip() if x is not None else "" for x in r]; start = i+1; break
    if header is None: return pd.DataFrame()
    data = [tuple(list(row)[:len(header)]) for row in rows[start:]]
    return pd.DataFrame(data, columns=header)

def exact_or_contains(df, names):
    for n in names:
        if n in df.columns: return n
    for n in names:
        for c in df.columns:
            if isinstance(c,str) and n in c: return c
    return None

def normalize_date(s):
    def f(x):
        if pd.isna(x): return pd.NaT
        if isinstance(x,(pd.Timestamp,datetime)): return pd.Timestamp(x.date())
        return pd.to_datetime(x, dayfirst=True, errors="coerce").normalize()
    return s.apply(f)

def to_number(s):
    return pd.to_numeric(s.astype(str).str.replace(",","").str.replace("₪","").str.strip(), errors="coerce")

def ref_starts_with_ov_rc(v):
    t = (str(v) if v is not None else "").strip().upper()
    return t.startswith("OV") or t.startswith("RC")

def normalize_text(s):
    if s is None: return ""
    t = str(s).replace("'", "").replace('"', "").replace("’","").replace("`","")
    t = t.replace("-", " ").replace("–"," ").replace("־"," ")
    return re.sub(r"\s+", " ", t).strip()

def process_workbook(xlsx_bytes: bytes):
    name_map   = DEFAULT_NAME_MAP.copy()
    amount_map = DEFAULT_AMOUNT_MAP.copy()

    wb_in = load_workbook(io.BytesIO(xlsx_bytes), data_only=True, read_only=True)
    out_stream = io.BytesIO()
    summary_rows, standing_rows = [], []

    with pd.ExcelWriter(out_stream, engine="xlsxwriter") as writer:
        for ws in wb_in.worksheets:
            df = ws_to_df(ws); df_save = df.copy()
            if df.empty:
                pd.DataFrame().to_excel(writer, index=False, sheet_name=ws.title); continue

            col_match = exact_or_contains(df, MATCH_COL_CANDS) or df.columns[0]
            col_bank_code = exact_or_contains(df, BANK_CODE_CANDS)
            col_bank_amt  = exact_or_contains(df, BANK_AMT_CANDS)
            col_books_amt = exact_or_contains(df, BOOKS_AMT_CANDS)
            col_ref       = exact_or_contains(df, REF_CANDS)
            col_date      = exact_or_contains(df, DATE_CANDS)
            col_details   = exact_or_contains(df, DETAILS_CANDS)

            applied_ovrc = applied_standing = False
            pairs = flagged = 0
            match_values = df_save[col_match].copy() if col_match in df_save.columns else pd.Series([None]*len(df_save))

            _date = normalize_date(pd.to_datetime(df[col_date], errors="coerce")) if col_date else pd.Series([pd.NaT]*len(df))
            _bank_amt  = to_number(df[col_bank_amt])  if col_bank_amt  else pd.Series([np.nan]*len(df))
            _books_amt = to_number(df[col_books_amt]) if col_books_amt else pd.Series([np.nan]*len(df))
            _bank_code = to_number(df[col_bank_code]) if col_bank_code else pd.Series([np.nan]*len(df))
            _ref = df[col_ref].astype(str).fillna("") if col_ref else pd.Series([""]*len(df))
            _details = df[col_details].astype(str).fillna("") if col_details else pd.Series([""]*len(df))

            # OV/RC
            if all([col_bank_code, col_bank_amt, col_books_amt, col_ref, col_date]):
                applied_ovrc = True
                books_cands = [i for i in range(len(df))
                               if pd.notna(_books_amt.iat[i]) and _books_amt.iat[i] > 0
                               and pd.notna(_date.iat[i]) and ref_starts_with_ov_rc(_ref.iat[i])]
                used = set()
                for i in range(len(df)):
                    if pd.notna(_bank_code.iat[i]) and int(_bank_code.iat[i]) in (175,120) \
                       and pd.notna(_bank_amt.iat[i]) and _bank_amt.iat[i] < 0 and pd.notna(_date.iat[i]):
                        amt = round(abs(float(_bank_amt.iat[i])),2)
                        d   = _date.iat[i]
                        cands = [j for j in books_cands if j not in used
                                 and _date.iat[j]==d and round(float(_books_amt.iat[j]),2)==amt]
                        chosen = None
                        if len(cands)==1: chosen = cands[0]
                        elif len(cands)>1: chosen = min(cands, key=lambda j: abs(j-i))
                        if chosen is not None:
                            match_values.iat[i] = 1; match_values.iat[chosen] = 1
                            used.add(chosen); pairs += 1

            # Standing orders
            if all([col_bank_code, col_details, col_bank_amt]):
                applied_standing = True
                for i in range(len(df)):
                    code = _bank_code.iat[i]
                    if pd.notna(code) and int(code) in (515,469):
                        match_values.iat[i] = 2; flagged += 1
                        standing_rows.append({"פרטים": _details.iat[i], "סכום": _bank_amt.iat[i]})

            df_out = df_save.copy(); df_out[col_match] = match_values
            df_out.to_excel(writer, index=False, sheet_name=ws.title)

            summary_rows.append({"גיליון": ws.title,
                                 "OV/RC בוצע": "כן" if applied_ovrc else "לא",
                                 "זוגות שסומנו 1": pairs,
                                 "הוראת קבע בוצע": "כן" if applied_standing else "לא",
                                 "שורות שסומנו 2": flagged,
                                 "עמודת התאמה": col_match})

        # גיליון הוראת קבע
        st_df = pd.DataFrame(standing_rows)
        if not st_df.empty:
            def map_supplier(n):
                s = normalize_text(n)
                if s in name_map := DEFAULT_NAME_MAP:  # exact
                    return name_map[s]
                for key in sorted(DEFAULT_NAME_MAP.keys(), key=len, reverse=True):
                    if key and key in s: return DEFAULT_NAME_MAP[key]
                return ""
            st_df["מס' ספק"] = st_df["פרטים"].apply(map_supplier)
            def by_amount(row):
                if not row["מס' ספק"]:
                    val = round(abs(float(row["סכום"])),2) if pd.notna(row["סכום"]) else None
                    return DEFAULT_AMOUNT_MAP.get(val, "")
                return row["מס' ספק"]
            st_df["מס' ספק"] = st_df.apply(by_amount, axis=1)
            st_df["סכום חובה"] = st_df["סכום"].apply(lambda x: x if pd.notna(x) and x>0 else 0)
            st_df["סכום זכות"] = st_df["סכום"].apply(lambda x: abs(x) if pd.notna(x) and x<0 else 0)
            st_df = st_df[["פרטים","סכום","מס' ספק","סכום חובה","סכום זכות"]]
        else:
            st_df = pd.DataFrame(columns=["פרטים","סכום","מס' ספק","סכום חובה","סכום זכות"])
        st_df.to_excel(writer, index=False, sheet_name="הוראת קבע ספקים")

    # RTL + צבע + סיכום 20001 (חובה → נרשם בזכות)
    wb_out = load_workbook(io.BytesIO(out_stream.getvalue()))
    for ws in wb_out.worksheets:
        ws.sheet_view.rightToLeft = True
    ws = wb_out["הוראת קבע ספקים"]
    hdr = {c.value: i for i,c in enumerate(ws[1], start=1)}
    col_supplier, col_debit, col_credit, col_details, col_amount = hdr["מס' ספק"], hdr["סכום חובה"], hdr["סכום זכות"], hdr["פרטים"], hdr["סכום"]
    orange = PatternFill(start_color="FFDDBB", end_color="FFDDBB", fill_type="solid")
    for r in range(2, ws.max_row+1):
        if ws.cell(row=r, column=col_supplier).value in ("", None):
            for c in range(1, ws.max_column+1): ws.cell(row=r, column=c).fill = orange
    # remove old 20001
    dels=[]; 
    for r in range(2, ws.max_row+1):
        v = ws.cell(row=r, column=col_supplier).value
        if v==20001 or (isinstance(v,str) and v.strip()=="20001"): dels.append(r)
    for k, r in enumerate(dels): ws.delete_rows(r-k, 1)
    # total debit for rows that have supplier -> write in credit
    total_from_debit = 0.0
    for r in range(2, ws.max_row+1):
        if ws.cell(row=r, column=col_supplier).value not in ("", None):
            try: total_from_debit += float(ws.cell(row=r, column=col_debit).value or 0)
            except: pass
    last = ws.max_row + 1
    ws.cell(row=last, column=col_details,  value='סה"כ זכות – עם מס\' ספק')
    ws.cell(row=last, column=col_amount,   value="")
    ws.cell(row=last, column=col_supplier, value=20001)
    ws.cell(row=last, column=col_debit,    value=0)
    ws.cell(row=last, column=col_credit,   value=round(total_from_debit, 2))
    for c in range(1, ws.max_column+1): ws.cell(row=last, column=c).font = Font(bold=True)

    final = io.BytesIO(); wb_out.save(final)
    return final.getvalue(), pd.DataFrame(summary_rows)

# ---------- UI ----------
uploaded = st.file_uploader("בחרי קובץ אקסל מקור (xlsx)", type=["xlsx"])
st.download_button("⬇️ הורדת קובץ כללים קבוע", rules_excel_bytes(),
                   file_name="VLOOKUP_rules_fixed.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if st.button("הרצה") and uploaded is not None:
    with st.spinner("מעבד..."):
        out_bytes, summary = process_workbook(uploaded.read())
    st.success("מוכן! אפשר להוריד את הקובץ המעודכן.")
    st.dataframe(summary, use_container_width=True)
    st.download_button("⬇️ הורדת קובץ מעודכן",
                       data=out_bytes,
                       file_name="התאמות_והוראת_קבע.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("ה-VLOOKUP קבוע בקוד. רק מעלים קובץ מקור ומריצים.")
