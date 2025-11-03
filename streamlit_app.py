# streamlit_app.py
# -*- coding: utf-8 -*-
"""
התאמות לקוחות/בנק – 1+2+3+4 + גיליון "הוראת קבע ספקים"
A2: לא דורסת ערך קיים ב"מס. התאמה" – מסמן רק אם ריק/0.
כולל כללי VLOOKUP (שם/פרטים → מספר ספק, סכום → מספר ספק) עם שמירה/ייבוא/ייצוא JSON.
RTL + פריסת עמוד להדפסה.
קובץ עזר (אופציונלי): אתחול כללים מגיליונות "שם→ספק" ו-"סכום→ספק".
"""

import io, os, re, json
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

# ---------------- קבועים ----------------
STANDING_CODES = {469, 515}     # כלל 2 – הוראת קבע
CHECK_CODE = 493                # כלל 4 – שיקים
AMOUNT_TOL_4 = 0.20             # טולרנס בכלל 4
RULES_PATH = "rules_store.json" # קובץ כללים לשמירה

# מועמדי שמות עמודות (זיהוי חכם)
MATCH_COL_CANDS = ["מס. התאמה", "מס התאמה", "התאמה", "מס' התאמה"]
BANK_CODE_CANDS = ["קוד פעולת בנק", "קוד פעולה בנק", "קוד פעולה", "קוד בנק"]
BANK_AMT_CANDS  = ["סכום בדף", "סכום בבנק", "סכום בנק"]
BOOKS_AMT_CANDS = ["סכום בספרים", "ספרים סכום", "סכום ספרים"]
REF1_CANDS      = ["אסמכתא 1", "אסמכתא", "מס' אסמכתא", "Reference 1"]
REF2_CANDS      = ["אסמכתא 2", "אסמכתא2", "אסמכתא נוספת", "Reference 2"]
DATE_CANDS      = ["תאריך מאזן", "תאריך", "ת. מאזן", "תאריך מסמך"]
DETAILS_CANDS   = ["פרטים", "תיאור", "תאור", "Details"]

HEADER_BLUE = PatternFill(start_color="9BC2E6", end_color="9BC2E6", fill_type="solid")
THIN  = Side(style='thin', color='CCCCCC')
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

# ---------------- כללי VLOOKUP – זיכרון ----------------
if "name_map" not in st.session_state:   st.session_state.name_map = {}
if "amount_map" not in st.session_state: st.session_state.amount_map = {}

def normalize_text(s: str) -> str:
    if s is None: return ""
    return str(s).replace("\u200f","").replace("\u200e","").strip()

def to_number(x) -> float:
    if pd.isna(x): return np.nan
    s = str(x).replace(",", "")
    try:
        return float(s)
    except Exception:
        m = re.findall(r"[-+]?\d+(?:\.\d+)?", s)
        return float(m[0]) if m else np.nan

def normalize_date(s: pd.Series) -> pd.Series:
    out = pd.to_datetime(s, errors="coerce")
    return pd.to_datetime(out.dt.date)

def ref_starts_with_ov_rc(s: str) -> bool:
    s = "" if s is None else str(s).strip().upper()
    return s.startswith("OV") or s.startswith("RC")

def exact_or_contains(df: pd.DataFrame, cands: List[str]) -> Optional[str]:
    cols = list(map(str, df.columns))
    for c in cands:
        if c in cols: return c
    for c in cands:
        for col in cols:
            if c in col: return col
    return None

def ws_to_df(ws) -> pd.DataFrame:
    rows = list(ws.values)
    if not rows: return pd.DataFrame()
    headers = [normalize_text(h) if h is not None else "" for h in rows[0]]
    return pd.DataFrame(rows[1:], columns=headers)

def load_rules_from_disk():
    if os.path.exists(RULES_PATH):
        try:
            data = json.load(open(RULES_PATH, "r", encoding="utf-8"))
            st.session_state.name_map.update(data.get("name_map", {}))
            st.session_state.amount_map.update(data.get("amount_map", {}))
        except Exception:
            pass

def save_rules_to_disk():
    with open(RULES_PATH, "w", encoding="utf-8") as f:
        json.dump({
            "name_map": st.session_state.name_map,
            "amount_map": st.session_state.amount_map
        }, f, ensure_ascii=False, indent=2)

def _only_digits(s: str) -> str:
    s = "" if s is None else str(s)
    d = re.sub(r"\D", "", s).lstrip("0")
    return d or "0"

load_rules_from_disk()
# ---------------- כלל 1 – OV/RC (סכום+תאריך 1:1) ----------------
def rule_1(df: pd.DataFrame, cols: Dict[str, str], match: pd.Series) -> Tuple[pd.Series, Dict[str,int]]:
    stats = {"pairs": 0}
    bamt = cols.get("bank_amt"); samt = cols.get("books_amt")
    ref  = cols.get("ref");      date = cols.get("date")
    if not all([bamt, samt, ref, date]): return match, stats

    d   = normalize_date(df[date])
    ba  = df[bamt].apply(to_number)
    sa  = df[samt].apply(to_number)
    r1  = df[ref].astype(str).fillna("")

    books = [i for i in df.index if pd.notna(sa.iat[i]) and sa.iat[i] > 0 and pd.notna(d.iat[i]) and ref_starts_with_ov_rc(r1.iat[i])]
    bank_keys, books_keys = {}, {}
    for i in df.index:
        if pd.notna(ba.iat[i]) and pd.notna(d.iat[i]):
            bank_keys.setdefault((abs(ba.iat[i]), d.iat[i]), []).append(i)
    for j in books:
        books_keys.setdefault((abs(sa.iat[j]), d.iat[j]), []).append(j)

    for k, bi in bank_keys.items():
        bj = books_keys.get(k, [])
        if len(bi)==1 and len(bj)==1:
            i, j = bi[0], bj[0]
            if float(match.iat[i] or 0)==0 and float(match.iat[j] or 0)==0:
                match.iat[i]=1; match.iat[j]=1; stats["pairs"]+=1
    return match, stats

# ---------------- כלל 2 – הוראות קבע + איסוף גיליון ----------------
def rule_2_and_collect(df: pd.DataFrame, cols: Dict[str,str], match: pd.Series, sheet_name: str) -> Tuple[pd.Series, Dict[str,int], pd.DataFrame]:
    stats = {"flagged": 0}
    code = cols.get("bank_code"); bamt = cols.get("bank_amt"); det = cols.get("details")
    rows = []
    if not all([code, bamt, det]): return match, stats, pd.DataFrame()

    vcode = df[code].apply(to_number)
    vdet  = df[det].astype(str).fillna("")
    vamt  = df[bamt].apply(to_number)

    for i in df.index:
        c = vcode.iat[i]
        if pd.isna(c): continue
        if int(c) in STANDING_CODES:
            if float(match.iat[i] or 0)==0:
                match.iat[i]=2; stats["flagged"]+=1
            rows.append({"גיליון מקור": sheet_name, "פרטים": vdet.iat[i], "סכום": float(vamt.iat[i]) if pd.notna(vamt.iat[i]) else np.nan})

    def map_supplier(name: str, amount_val) -> str:
        name = normalize_text(name)
        if name in st.session_state.name_map: return st.session_state.name_map[name]
        for key, sup in st.session_state.name_map.items():
            if key and key in name: return sup
        if amount_val is not None and not pd.isna(amount_val):
            s = str(int(amount_val)) if float(amount_val).is_integer() else str(amount_val)
            if s in st.session_state.amount_map: return st.session_state.amount_map[s]
            s_round = str(int(round(float(amount_val))))
            if s_round in st.session_state.amount_map: return st.session_state.amount_map[s_round]
        return ""

    if rows:
        out = pd.DataFrame(rows)
        out["מספר ספק"] = out.apply(lambda r: map_supplier(r.get("פרטים",""), r.get("סכום", np.nan)), axis=1)
        return match, stats, out
    return match, stats, pd.DataFrame()

# ---------------- כלל 3 – סכום+תאריך 1:1 ללא OV/RC ----------------
def rule_3(df: pd.DataFrame, cols: Dict[str,str], match: pd.Series) -> Tuple[pd.Series, Dict[str,int]]:
    stats = {"pairs": 0}
    bamt = cols.get("bank_amt"); samt = cols.get("books_amt")
    date = cols.get("date"); ref = cols.get("ref"); code = cols.get("bank_code")
    if not all([bamt, samt, date]): return match, stats

    d   = normalize_date(df[date])
    ba  = df[bamt].apply(to_number)   # בנק – שלילי
    sa  = df[samt].apply(to_number)   # ספרים – חיובי
    r1  = df[ref].astype(str).fillna("") if (ref and ref in df.columns) else pd.Series([""]*len(df))
    vc  = df[code].apply(to_number)   if (code and code in df.columns) else pd.Series([np.nan]*len(df))

    books = [i for i in df.index if float(match.iat[i] or 0)==0 and pd.notna(sa.iat[i]) and sa.iat[i]>0 and not ref_starts_with_ov_rc(r1.iat[i]) and pd.notna(d.iat[i])]
    banks = [i for i in df.index if float(match.iat[i] or 0)==0 and pd.notna(ba.iat[i]) and ba.iat[i]<0 and pd.notna(d.iat[i])
             and (pd.isna(vc.iat[i]) or (int(vc.iat[i]) not in STANDING_CODES and int(vc.iat[i])!=CHECK_CODE))]

    bank_keys, books_keys = {}, {}
    for i in banks:
        bank_keys.setdefault((abs(ba.iat[i]), d.iat[i]), []).append(i)
    for j in books:
        books_keys.setdefault((abs(sa.iat[j]), d.iat[j]), []).append(j)

    for k, bi in bank_keys.items():
        bj = books_keys.get(k, [])
        if len(bi)==1 and len(bj)==1:
            i, j = bi[0], bj[0]
            if float(match.iat[i] or 0)==0 and float(match.iat[j] or 0)==0:
                match.iat[i]=3; match.iat[j]=3; stats["pairs"]+=1
    return match, stats

# ---------------- כלל 4 – שיקי ספקים ----------------
def rule_4(df: pd.DataFrame, cols: Dict[str,str], match: pd.Series) -> Tuple[pd.Series, Dict[str,int]]:
    stats = {"pairs": 0}
    code = cols.get("bank_code"); bamt = cols.get("bank_amt"); samt = cols.get("books_amt")
    ref1 = cols.get("ref"); ref2 = cols.get("ref2"); det = cols.get("details")
    if not all([code, bamt, samt, ref1, ref2, det]): return match, stats

    vc   = df[code].apply(to_number)
    ba   = df[bamt].apply(to_number)
    sa   = df[samt].apply(to_number)
    ddet = df[det].astype(str).fillna("")
    r1   = df[ref1].astype(str).fillna("")
    r2   = df[ref2].astype(str).fillna("")
    r1n  = r1.apply(_only_digits)
    r2n  = r2.apply(_only_digits)

    bank_idx = [i for i in df.index
                if float(match.iat[i] or 0)==0 and pd.notna(vc.iat[i]) and int(vc.iat[i])==CHECK_CODE
                and "שיק" in ddet.iat[i] and pd.notna(ba.iat[i]) and ba.iat[i] > 0]
    books_idx = [j for j in df.index
                 if float(match.iat[j] or 0)==0 and str(r1.iat[j]).startswith("CH")
                 and "תשלום בהמחאה" in ddet.iat[j] and pd.notna(sa.iat[j]) and sa.iat[j] < 0]

    books_by_id = {}
    for j in books_idx:
        books_by_id.setdefault(r2n.iat[j], []).append(j)

    used = set()
    for i in bank_idx:
        key = r1n.iat[i]
        cands = [j for j in books_by_id.get(key, []) if j not in used]
        chosen = None
        for j in cands:
            if abs(float(ba.iat[i]) + float(sa.iat[j])) <= AMOUNT_TOL_4:
                chosen = j; break
        if chosen is not None:
            match.iat[i]=4; match.iat[chosen]=4; used.add(chosen); stats["pairs"]+=1
    return match, stats
# ---------------- עיבוד קובץ ----------------
def process_workbook(main_bytes: bytes, aux_bytes: Optional[bytes] = None):
    wb_in = load_workbook(io.BytesIO(main_bytes), data_only=True, read_only=True)

    # אתחול כללים מתוך קובץ עזר (אופציונלי)
    if aux_bytes:
        try:
            aux = pd.ExcelFile(io.BytesIO(aux_bytes))
            if "שם→ספק" in aux.sheet_names:
                df1 = pd.read_excel(aux, "שם→ספק")
                for _, r in df1.fillna("").iterrows():
                    k = normalize_text(r.get("שם") or r.get("שם/פרטים") or "")
                    v = normalize_text(r.get("מספר ספק") or r.get("מס' ספק") or "")
                    if k and v: st.session_state.name_map[k] = v
            if "סכום→ספק" in aux.sheet_names:
                df2 = pd.read_excel(aux, "סכום→ספק")
                for _, r in df2.fillna("").iterrows():
                    k = normalize_text(r.get("סכום") or "")
                    v = normalize_text(r.get("מספר ספק") or r.get("מס' ספק") or "")
                    if k and v: st.session_state.amount_map[k] = v
            save_rules_to_disk()
        except Exception:
            pass

    out_stream = io.BytesIO()
    summary_rows = []
    standing_collect = []

    with pd.ExcelWriter(out_stream, engine="xlsxwriter") as writer:
        for ws in wb_in.worksheets:
            df = ws_to_df(ws)
            df0 = df.copy()
            if df.empty:
                pd.DataFrame().to_excel(writer, index=False, sheet_name=ws.title)
                continue

            col_match     = exact_or_contains(df, MATCH_COL_CANDS) or "מס. התאמה"
            col_bank_code = exact_or_contains(df, BANK_CODE_CANDS)
            col_bank_amt  = exact_or_contains(df, BANK_AMT_CANDS)
            col_books_amt = exact_or_contains(df, BOOKS_AMT_CANDS)
            col_ref1      = exact_or_contains(df, REF1_CANDS)
            col_ref2      = exact_or_contains(df, REF2_CANDS)
            col_date      = exact_or_contains(df, DATE_CANDS)
            col_details   = exact_or_contains(df, DETAILS_CANDS)

            if col_match not in df0.columns:
                df0[col_match] = 0
            match_values = df0[col_match].copy().fillna(0)

            cols = {"bank_code": col_bank_code, "bank_amt": col_bank_amt, "books_amt": col_books_amt,
                    "ref": col_ref1, "ref2": col_ref2, "date": col_date, "details": col_details}

            # כלל 1
            match_values, s1 = rule_1(df, cols, match_values)
            # כלל 2 + איסוף
            match_values, s2, rows2 = rule_2_and_collect(df, cols, match_values, ws.title)
            if not rows2.empty: standing_collect.append(rows2)
            # כלל 3
            match_values, s3 = rule_3(df, cols, match_values)
            # כלל 4
            match_values, s4 = rule_4(df, cols, match_values)

            df_out = df0.copy()
            # A2: לא לדרוס סימונים קיימים – match_values כבר שומר על זה (מסמן רק אם 0)
            df_out[col_match] = match_values
            df_out.to_excel(writer, index=False, sheet_name=ws.title)

            summary_rows.append({
                "גיליון": ws.title,
                "זוגות שסומנו 1": s1.get("pairs", 0),
                "שורות שסומנו 2": s2.get("flagged", 0),
                "זוגות שסומנו 3": s3.get("pairs", 0),
                "זוגות שסומנו 4": s4.get("pairs", 0),
                "עמודת התאמה": col_match,
            })

        # גיליון הוראת קבע ספקים
        if standing_collect:
            st_df = pd.concat(standing_collect, ignore_index=True)
        else:
            st_df = pd.DataFrame([{"גיליון מקור": "—", "הודעה": "אין שורות עם מס. התאמה = 2 (הוראת קבע)"}])
        st_df.to_excel(writer, index=False, sheet_name="הוראת קבע ספקים")

        # גיליון סיכום
        pd.DataFrame(summary_rows).to_excel(writer, index=False, sheet_name="סיכום")

    # עיצוב/RTL/פריסת עמוד
    out_bytes = out_stream.getvalue()
    wb = load_workbook(io.BytesIO(out_bytes))
    def style_sheet(ws):
        try:
            ws.sheet_view.rightToLeft = True
            ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
            ws.page_setup.paperSize = ws.PAPERSIZE_A4
            ws.sheet_properties.pageSetUpPr.fitToPage = True
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.5, bottom=0.5, header=0.3, footer=0.3)
        except Exception:
            pass
        if ws.max_row >= 1:
            for c in ws[1]:
                c.font = Font(bold=True); c.fill = HEADER_BLUE
                c.alignment = Alignment(horizontal='center', vertical='center'); c.border = BORDER
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.alignment = Alignment(horizontal='right', vertical='center'); cell.border = BORDER
        for col_idx in range(1, ws.max_column+1):
            col_letter = get_column_letter(col_idx); maxlen = 0
            for r in range(1, ws.max_row+1):
                v = ws.cell(row=r, column=col_idx).value
                if v is None: continue
                maxlen = max(maxlen, len(str(v)))
            ws.column_dimensions[col_letter].width = min(maxlen + 2, 60)
    for ws in wb.worksheets: style_sheet(ws)
    final = io.BytesIO(); wb.save(final)
    return final.getvalue(), pd.DataFrame(summary_rows)
# ---------------- UI ----------------
st.set_page_config(page_title="התאמות 1+2+3+4 + הוראת קבע", page_icon="✅", layout="centered")
st.markdown("""
<style>
  html, body, [class*="css"] { direction: rtl; text-align: right; }
  .block-container { padding-top: 1rem; }
</style>
""", unsafe_allow_html=True)

st.title("התאמות לקוחות/בנק – 1+2+3+4 + גיליון הוראת קבע")
st.subheader("⚙️ עדכון – כללי VLOOKUP (עם שמירה)")

mode = st.radio("סוג עדכון", ["לפי פרטים (שם)", "לפי סכום"], horizontal=True)
colA, colB, colC = st.columns([2,1,1])
if mode == "לפי פרטים (שם)":
    name_key = colA.text_input("שם/פרטים (המפתח לחיפוש)")
    name_val = colB.text_input("מס' ספק")
    if colC.button("הוסף/עדכן שם→ספק"):
        if name_key and name_val:
            st.session_state.name_map[normalize_text(name_key)] = normalize_text(name_val)
            save_rules_to_disk(); st.success("נשמר")
else:
    amt_key = colA.text_input("סכום (כמספר או טקסט)")
    amt_val = colB.text_input("מס' ספק")
    if colC.button("הוסף/עדכן סכום→ספק"):
        if amt_key and amt_val:
            st.session_state.amount_map[str(amt_key).strip()] = normalize_text(amt_val)
            save_rules_to_disk(); st.success("נשמר")

c1, c2, c3, c4 = st.columns([1,1,1,2])
if c1.button("נקה הכללים"):
    st.session_state.name_map = {}; st.session_state.amount_map = {}
    save_rules_to_disk(); st.warning("נוקה ונשמר")

c2.download_button("⬇️ ייצוא JSON",
    data=json.dumps({"name_map": st.session_state.name_map, "amount_map": st.session_state.amount_map},
                    ensure_ascii=False, indent=2).encode("utf-8"),
    file_name="rules_store.json", mime="application/json")
up_rules = c3.file_uploader("⬆️ ייבוא JSON", type=["json"], label_visibility="collapsed")
if c4.button("ייבוא והחלפה") and up_rules is not None:
    data = json.loads(up_rules.getvalue().decode("utf-8"))
    st.session_state.name_map = data.get("name_map", {}); st.session_state.amount_map = data.get("amount_map", {})
    save_rules_to_disk(); st.success("נטען ונשמר")

# טבלאות כללים (ללא מיון מספרי)
st.dataframe(pd.DataFrame({"שם/פרטים": list(st.session_state.name_map.keys()),
                           "מס' ספק":  list(st.session_state.name_map.values())}), use_container_width=True, height=220)
st.dataframe(pd.DataFrame({"סכום": list(st.session_state.amount_map.keys()),
                           "מס' ספק": list(st.session_state.amount_map.values())}), use_container_width=True, height=220)

st.divider()

# העלאה והרצה
cx1, cx2 = st.columns([2,2])
main_file = cx1.file_uploader("בחרי קובץ אקסל מקור (xlsx)", type=["xlsx"])
aux_file  = cx2.file_uploader("(אופציונלי) קובץ עזר להעברות/מיפויים (גיליונות: שם→ספק, סכום→ספק)", type=["xlsx"])

run = st.button("▶️ הרצה – הפקת קובץ מסומן + גיליון הוראת קבע")
if run and main_file is not None:
    main_bytes = main_file.read()
    aux_bytes  = aux_file.read() if aux_file is not None else None
    out_bytes, summary_df = process_workbook(main_bytes, aux_bytes)
    st.success("הקובץ עובד! אפשר להוריד.")
    st.download_button("⬇️ הורדה – תוצאה סופית (1+2+3+4 + הוראת קבע)",
        data=out_bytes,
        file_name="תוצאה_סופית_התאמות_1_2_3_4_והוראת_קבע.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with st.expander("תצוגת סיכום"):
        st.dataframe(summary_df, use_container_width=True)
else:
    st.info("בחרי קובץ מקור (ואופציונלית קובץ עזר) ולחצי הרצה.")
