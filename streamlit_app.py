# streamlit_app.py
# -*- coding: utf-8 -*-

import io
import re
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font

# ===== UI base (RTL) =====
st.set_page_config(page_title="התאמות לקוחות – OV/RC + הוראות קבע", page_icon="✅", layout="centered")
st.markdown("""
<style>
  html, body, [class*="css"] { direction: rtl; text-align: right; }
  .block-container { padding-top: 1.3rem; }
</style>
""", unsafe_allow_html=True)

st.title("התאמות לקוחות – OV/RC + הוראות קבע (גרסה מעודכנת)")
st.caption("מעלה קובץ אקסל ״דוגמה 2״ (או כל קובץ דומה), ומקבל קובץ מעובד עם עמודת ״מס. התאמה״ + גיליון ״הוראת קבע ספקים״, צביעת חסרים וסיכום 20001.")

# ====================== Helpers ======================

MATCH_COL_CANDS = ["מס.התאמה","מס. התאמה","מס התאמה","מספר התאמה","התאמה"]
BANK_CODE_CANDS = ["קוד פעולת בנק","קוד פעולה","קוד פעולת"]
BANK_AMT_CANDS  = ["סכום בדף","סכום דף","סכום בבנק","סכום תנועת בנק"]
BOOKS_AMT_CANDS = ["סכום בספרים","סכום בספר","סכום ספרים"]
REF_CANDS       = ["אסמכתא 1","אסמכתא1","אסמכתא","אסמכתה"]
DATE_CANDS      = ["תאריך מאזן","תאריך ערך","תאריך"]
DETAILS_CANDS   = ["פרטים","תיאור","שם ספק"]

def ws_to_df(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return pd.DataFrame()
    header = None; start_idx = 0
    for i, r in enumerate(rows):
        if any(x is not None for x in r):
            header = [str(x).strip() if x is not None else "" for x in r]
            start_idx = i + 1
            break
    if header is None:
        return pd.DataFrame()
    data_rows = [tuple(list(row)[:len(header)]) for row in rows[start_idx:]]
    return pd.DataFrame(data_rows, columns=header)

def exact_or_contains(df: pd.DataFrame, names: list) -> str | None:
    for n in names:
        if n in df.columns:
            return n
    for n in names:
        for c in df.columns:
            if isinstance(c, str) and n in c:
                return c
    return None

def normalize_date(series: pd.Series) -> pd.Series:
    def to_date(x):
        if pd.isna(x): return pd.NaT
        if isinstance(x,(pd.Timestamp, datetime)): return pd.Timestamp(x.date())
        return pd.to_datetime(x, dayfirst=True, errors="coerce").normalize()
    return series.apply(to_date)

def to_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series.astype(str).str.replace(",","").str.replace("₪","").str.strip(), errors="coerce")

def ref_starts_with_ov_rc(val) -> bool:
    t = (str(val) if val is not None else "").strip().upper()
    return t.startswith("OV") or t.startswith("RC")

def normalize_text(s: str) -> str:
    if s is None: return ""
    t = str(s)
    t = t.replace("'", "").replace('"', "").replace("’","").replace("`","")
    t = t.replace("-", " ").replace("–"," ").replace("־"," ")
    t = re.sub(r"\s+", " ", t).strip()
    return t

# ===== Default VLOOKUP rules (can be extended by user file) =====
DEFAULT_NAME_MAP = {
    "עיריית אשדוד": 30056,
    "ישראכרט מור": 34002,
}
DEFAULT_AMOUNT_MAP = {
    # סכום → מס' ספק
    8520.0: 30247,
    10307.3: 30038,
}

def build_lookup_maps(lookup_file: bytes | None):
    name_map = dict(DEFAULT_NAME_MAP)
    amount_map = dict(DEFAULT_AMOUNT_MAP)
    if lookup_file:
        xls = pd.ExcelFile(io.BytesIO(lookup_file))
        if "by_name" in xls.sheet_names:
            df = pd.read_excel(io.BytesIO(lookup_file), sheet_name="by_name")
            ncol, icol = df.columns[0], df.columns[1]
            for _, r in df[[ncol, icol]].dropna(subset=[ncol]).iterrows():
                name_map[normalize_text(r[ncol])] = r[icol]
        if "by_amount" in xls.sheet_names:
            df2 = pd.read_excel(io.BytesIO(lookup_file), sheet_name="by_amount")
            acol, icol2 = df2.columns[0], df2.columns[1]
            for _, r in df2[[acol, icol2]].dropna(subset=[acol]).iterrows():
                try:
                    amount_map[round(abs(float(r[acol])),2)] = r[icol2]
                except Exception:
                    pass
    return name_map, amount_map

def process_workbook(xlsx_bytes: bytes, lookup_bytes: bytes | None):
    name_map, amount_map = build_lookup_maps(lookup_bytes)

    wb_in = load_workbook(io.BytesIO(xlsx_bytes), data_only=True, read_only=True)
    out_stream = io.BytesIO()
    summary_rows = []
    standing_rows = []

    with pd.ExcelWriter(out_stream, engine="xlsxwriter") as writer:
        for ws in wb_in.worksheets:
            df = ws_to_df(ws)
            df_save = df.copy()
            if df.empty:
                pd.DataFrame().to_excel(writer, index=False, sheet_name=ws.title)
                continue

            col_match     = exact_or_contains(df, MATCH_COL_CANDS) or df.columns[0]
            col_bank_code = exact_or_contains(df, BANK_CODE_CANDS)
            col_bank_amt  = exact_or_contains(df, BANK_AMT_CANDS)
            col_books_amt = exact_or_contains(df, BOOKS_AMT_CANDS)
            col_ref       = exact_or_contains(df, REF_CANDS)
            col_date      = exact_or_contains(df, DATE_CANDS)
            col_details   = exact_or_contains(df, DETAILS_CANDS)

            applied_ovrc = False
            applied_standing = False
            pairs = 0
            flagged = 0

            match_values = df_save[col_match].copy() if col_match in df_save.columns else pd.Series([None]*len(df_save))
            _date = normalize_date(pd.to_datetime(df[col_date], errors="coerce")) if col_date else pd.Series([pd.NaT]*len(df))
            _bank_amt  = to_number(df[col_bank_amt])  if col_bank_amt  else pd.Series([np.nan]*len(df))
            _books_amt = to_number(df[col_books_amt]) if col_books_amt else pd.Series([np.nan]*len(df))
            _bank_code = to_number(df[col_bank_code]) if col_bank_code else pd.Series([np.nan]*len(df))
            _ref = df[col_ref].astype(str).fillna("") if col_ref else pd.Series([""]*len(df))
            _details = df[col_details].astype(str).fillna("") if col_details else pd.Series([""]*len(df))

            # Rule 1: OV/RC (set 1)
            if all([col_bank_code, col_bank_amt, col_books_amt, col_ref, col_date]):
                applied_ovrc = True
                books_candidates = [i for i in range(len(df))
                                    if pd.notna(_books_amt.iat[i]) and _books_amt.iat[i] > 0
                                    and pd.notna(_date.iat[i]) and ref_starts_with_ov_rc(_ref.iat[i])]
                used_books = set()
                for i in range(len(df)):
                    if pd.notna(_bank_code.iat[i]) and int(_bank_code.iat[i]) in (175,120) \
                       and pd.notna(_bank_amt.iat[i]) and _bank_amt.iat[i] < 0 and pd.notna(_date.iat[i]):
                        target_amt = round(abs(float(_bank_amt.iat[i])),2)
                        target_date = _date.iat[i]
                        cands = [j for j in books_candidates
                                 if j not in used_books
                                 and _date.iat[j] == target_date
                                 and round(float(_books_amt.iat[j]),2) == target_amt]
                        chosen = None
                        if len(cands)==1:
                            chosen = cands[0]
                        elif len(cands)>1:
                            chosen = min(cands, key=lambda j: abs(j-i))
                        if chosen is not None:
                            match_values.iat[i] = 1
                            match_values.iat[chosen] = 1
                            used_books.add(chosen)
                            pairs += 1

            # Rule 2: Standing orders (515/469) → set 2 + collect rows
            if all([col_bank_code, col_details, col_bank_amt]):
                applied_standing = True
                for i in range(len(df)):
                    code = _bank_code.iat[i]
                    if pd.notna(code) and int(code) in (515,469):
                        match_values.iat[i] = 2
                        flagged += 1
                        standing_rows.append({
                            "פרטים": _details.iat[i],
                            "סכום":  _bank_amt.iat[i]
                        })

            df_out = df_save.copy()
            df_out[col_match] = match_values
            df_out.to_excel(writer, index=False, sheet_name=ws.title)

            summary_rows.append({
                "גיליון": ws.title,
                "OV/RC בוצע": "כן" if applied_ovrc else "לא",
                "זוגות שסומנו 1": pairs,
                "הוראת קבע בוצע": "כן" if applied_standing else "לא",
                "שורות שסומנו 2": flagged,
                "עמודת התאמה": col_match
            })

        # Build standing orders sheet (name priority -> amount)
        st_df = pd.DataFrame(standing_rows)
        if not st_df.empty:
            def map_supplier_name(name):
                s = normalize_text(name)
                if s in name_map:
                    return name_map[s]
                for key in sorted(name_map.keys(), key=len, reverse=True):
                    if key and key in s:
                        return name_map[key]
                return ""

            st_df["מס' ספק"] = st_df["פרטים"].apply(map_supplier_name)

            def fill_by_amount(row):
                if row["מס' ספק"] in ("", None) or (isinstance(row["מס' ספק"], float) and pd.isna(row["מס' ספק"])):
                    amt = row["סכום"]
                    if pd.notna(amt):
                        val = round(abs(float(amt)), 2)
                        if val in amount_map:
                            return amount_map[val]
                return row["מס' ספק"]

            st_df["מס' ספק"] = st_df.apply(fill_by_amount, axis=1)
            st_df["סכום חובה"] = st_df["סכום"].apply(lambda x: x if pd.notna(x) and x>0 else 0)
            st_df["סכום זכות"] = st_df["סכום"].apply(lambda x: abs(x) if pd.notna(x) and x<0 else 0)
            st_df = st_df[["פרטים","סכום","מס' ספק","סכום חובה","סכום זכות"]]
        else:
            st_df = pd.DataFrame(columns=["פרטים","סכום","מס' ספק","סכום חובה","סכום זכות"])

        st_df.to_excel(writer, index=False, sheet_name="הוראת קבע ספקים")

    # Add RTL + coloring + summary-row (20001)
    wb_out = load_workbook(io.BytesIO(out_stream.getvalue()))
    for ws in wb_out.worksheets:
        ws.sheet_view.rightToLeft = True

    ws = wb_out["הוראת קבע ספקים"]

    # 1) צבע כתום בהיר למי שאין מס' ספק
    headers = {cell.value: idx for idx, cell in enumerate(ws[1], start=1)}
    col_supplier = headers.get("מס' ספק")
    col_debit    = headers.get("סכום חובה")
    col_credit   = headers.get("סכום זכות")
    col_details  = headers.get("פרטים")
    col_amount   = headers.get("סכום")

    orange = PatternFill(start_color="FFDDBB", end_color="FFDDBB", fill_type="solid")
    if col_supplier:
        for r in range(2, ws.max_row+1):
            v = ws.cell(row=r, column=col_supplier).value
            if v in ("", None):
                for c in range(1, ws.max_column+1):
                    ws.cell(row=r, column=c).fill = orange

    # 2) סיכום 20001 – סכום חובה (רק לשורות שיש להן מס' ספק) – נרשם בעמודת "סכום זכות"
    #    קודם מוחקים 20001 קודמים אם יש.
    to_delete = []
    for r in range(2, ws.max_row+1):
        v = ws.cell(row=r, column=col_supplier).value
        if v == 20001 or (isinstance(v,str) and v.strip()=="20001"):
            to_delete.append(r)
    for j, r in enumerate(to_delete):
        ws.delete_rows(r-j, 1)

    total_from_debit = 0.0
    for r in range(2, ws.max_row+1):
        supplier_val = ws.cell(row=r, column=col_supplier).value
        if supplier_val not in (None, ""):
            try:
                total_from_debit += float(ws.cell(row=r, column=col_debit).value or 0)
            except Exception:
                pass

    last = ws.max_row + 1
    if col_details:  ws.cell(row=last, column=col_details,  value="סה\"כ זכות – עם מס' ספק")
    if col_amount:   ws.cell(row=last, column=col_amount,   value="")
    if col_supplier: ws.cell(row=last, column=col_supplier, value=20001)
    if col_debit:    ws.cell(row=last, column=col_debit,    value=0)
    if col_credit:   ws.cell(row=last, column=col_credit,   value=round(total_from_debit, 2))
    for c in range(1, ws.max_column+1):
        ws.cell(row=last, column=c).font = Font(bold=True)

    final_bytes = io.BytesIO()
    wb_out.save(final_bytes)
    return final_bytes.getvalue(), pd.DataFrame(summary_rows)

# ====================== UI ======================
uploaded = st.file_uploader("בחרי קובץ אקסל מקור (xlsx)", type=["xlsx"])
lookup   = st.file_uploader("קובץ VLOOKUP (לא חובה): גיליונות by_name + by_amount", type=["xlsx"])

if st.button("הרצה") and uploaded is not None:
    with st.spinner("מעבד..."):
        out_bytes, summary = process_workbook(uploaded.read(), lookup.read() if lookup else None)
    st.success("מוכן! אפשר להוריד את הקובץ המעודכן.")
    st.dataframe(summary, use_container_width=True)
    st.download_button("⬇️ הורדת קובץ מעודכן", data=out_bytes,
                       file_name="התאמות_והוראת_קבע.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("טיפ: אפשר לצרף גם קובץ VLOOKUP מותאם (by_name/by_amount). אם לא מצורף – ייעשה שימוש בכללי ברירת המחדל.")
