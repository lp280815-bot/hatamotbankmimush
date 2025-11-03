# streamlit_app.py
# -*- coding: utf-8 -*-
"""
אפליקציית Streamlit לסימון התאמות 1+2+3+4 בקבצי Excel (לקוחות/בנקים),
כולל הפקת גיליון "הוראת קבע ספקים", פריסת עמוד/RTL,
וכפתור "עדכון – כללי VLOOKUP קבועים ומורחבים (עם שמירה)".

A2: לא דורסים ערכים קיימים בעמודה "מס. התאמה" – מסמנים רק אם ריק/0.

כללים:
1) OV/RC: התאמה 1:1 לפי (abs סכום, אותו תאריך) בין ספרים(+) ↔ בנק(−) כשאסמכתא1 מתחילה OV/RC.
2) הוראות קבע: מסמן 2 על שורות בנק עם קודים {469, 515} + בניית גיליון "הוראת קבע ספקים".
3) (רזרבה להרחבה).
4) שיקים ספקים: בנק (קוד 493, "שיק", סכום בדף>0) ↔ ספרים (CH..., "תשלום בהמחאה", סכום בספרים<0),
   התאמה ע"פ אסמכתא2 בספרים = אסמכתא1 בבנק (ספרות בלבד), טולרנס ₪0.20, בלי בדיקת תאריכים.

VLOOKUP קבוע/מורחב: מפות name→supplier ו-amount→supplier, שמירה/ייבוא/ייצוא ל- rules_store.json,
ושימוש במפות למילוי "מספר ספק" בגיליון "הוראת קבע ספקים".
"""

import io, os, re, json
from typing import Dict, List, Tuple, Optional

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

# ---------------- הגדרות ----------------
STANDING_CODES = {469, 515}     # קודי פעולה לזיהוי הוראת קבע
AMOUNT_TOL_4   = 0.20           # טולרנס לכלל 4 – שיקים ספקים
RULES_PATH     = "rules_store.json"  # שמירת כללי VLOOKUP

# מועמדי שמות עמודות
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

# ---------------- כללי VLOOKUP – טעינה ושמירה ----------------
if "name_map" not in st.session_state:   st.session_state.name_map = {}
if "amount_map" not in st.session_state: st.session_state.amount_map = {}

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

load_rules_from_disk()

# ---------------- כלי עזר ----------------
def normalize_text(s: str) -> str:
    if s is None: return ""
    s = str(s).replace("\u200f","").replace("\u200e","").strip()
    return s

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

def ref_starts_with_ov_rc(ref: str) -> bool:
    if not isinstance(ref, str):
        ref = str(ref) if ref is not None else ""
    ref = ref.strip().upper()
    return ref.startswith("OV") or ref.startswith("RC")

def exact_or_contains(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols = list(map(str, df.columns))
    for c in candidates:
        if c in cols:
            return c
    for c in candidates:
        for col in cols:
            if c in col:
                return col
    return None

def ws_to_df(ws) -> pd.DataFrame:
    data = ws.values
    try:
        headers = next(data)
    except StopIteration:
        return pd.DataFrame()
    headers = [normalize_text(h) if h is not None else "" for h in headers]
    return pd.DataFrame(list(data), columns=headers)

# ---------------- כלל 1 – OV/RC ----------------
def _apply_rule_1(df: pd.DataFrame, cols: Dict[str, str], match_values: pd.Series) -> Tuple[pd.Series, Dict[str, int]]:
    stats = {"pairs": 0}
    col_bank_amt  = cols.get("bank_amt")
    col_books_amt = cols.get("books_amt")
    col_ref       = cols.get("ref")
    col_date      = cols.get("date")
    if not all([col_bank_amt, col_books_amt, col_ref, col_date]):
        return match_values, stats

    _date      = normalize_date(pd.to_datetime(df[col_date], errors="coerce"))
    _bank_amt  = df[col_bank_amt].apply(to_number)
    _books_amt = df[col_books_amt].apply(to_number)
    _ref       = df[col_ref].astype(str).fillna("")

    books_candidates = [
        j for j in range(len(df))
        if pd.notna(_books_amt.iat[j]) and _books_amt.iat[j] > 0
        and pd.notna(_date.iat[j]) and ref_starts_with_ov_rc(_ref.iat[j])
    ]

    bank_keys, books_keys = {}, {}
    for i in range(len(df)):
        a = _bank_amt.iat[i]; d = _date.iat[i]
        if pd.notna(a) and pd.notna(d):
            bank_keys.setdefault((abs(a), d), []).append(i)
    for j in books_candidates:
        a = _books_amt.iat[j]; d = _date.iat[j]
        if pd.notna(a) and pd.notna(d):
            books_keys.setdefault((abs(a), d), []).append(j)

    for k, b_idx in bank_keys.items():
        if len(b_idx) == 1 and len(books_keys.get(k, [])) == 1:
            i = b_idx[0]; j = books_keys[k][0]
            if float(match_values.iat[i] or 0) == 0 and float(match_values.iat[j] or 0) == 0:
                match_values.iat[i] = 1
                match_values.iat[j] = 1
                stats["pairs"] += 1
    return match_values, stats

# ---------------- כלל 2 – הוראות קבע + איסוף גיליון ----------------
def _apply_rule_2_and_collect(df: pd.DataFrame, cols: Dict[str, str], match_values: pd.Series, sheet_name: str) -> Tuple[pd.Series, Dict[str, int], pd.DataFrame]:
    stats = {"flagged": 0}
    col_bank_code = cols.get("bank_code")
    col_bank_amt  = cols.get("bank_amt")
    col_details   = cols.get("details")
    rows = []
    if not all([col_bank_code, col_details, col_bank_amt]):
        return match_values, stats, pd.DataFrame(rows)

    _bank_code = df[col_bank_code].apply(to_number)
    _details   = df[col_details].astype(str).fillna("")
    _bank_amt  = df[col_bank_amt].apply(to_number)

    for i in range(len(df)):
        code = _bank_code.iat[i]
        if pd.isna(code): continue
        if int(code) in STANDING_CODES:
            if float(match_values.iat[i] or 0) == 0:
                match_values.iat[i] = 2
                stats["flagged"] += 1
            rows.append({
                "גיליון מקור": sheet_name,
                "פרטים": _details.iat[i],
                "סכום": float(_bank_amt.iat[i]) if pd.notna(_bank_amt.iat[i]) else np.nan
            })

    # מיפוי ספק לפי כללים שמורים (שם/תת־מחרוזת/סכום)
    def map_supplier(name: str, amount_val) -> str:
        name = normalize_text(name)
        if name in st.session_state.name_map:
            return st.session_state.name_map[name]
        for key, sup in st.session_state.name_map.items():
            if key and key in name:
                return sup
        if amount_val is not None and not pd.isna(amount_val):
            s = str(int(amount_val)) if float(amount_val).is_integer() else str(amount_val)
            if s in st.session_state.amount_map:
                return st.session_state.amount_map[s]
            s_round = str(int(round(float(amount_val))))
            if s_round in st.session_state.amount_map:
                return st.session_state.amount_map[s_round]
        return ""

    if rows:
        df_rows = pd.DataFrame(rows)
        df_rows["מספר ספק"] = df_rows.apply(lambda r: map_supplier(r.get("פרטים",""), r.get("סכום", np.nan)), axis=1)
        return match_values, stats, df_rows
    return match_values, stats, pd.DataFrame()

# ---------------- כלל 4 – שיקי ספקים ----------------
def _apply_rule_4(df: pd.DataFrame, cols: Dict[str, str], match_values: pd.Series) -> Tuple[pd.Series, Dict[str, int]]:
    stats = {"pairs": 0}
    col_bank_code = cols.get("bank_code")
    col_bank_amt  = cols.get("bank_amt")
    col_books_amt = cols.get("books_amt")
    col_ref1      = cols.get("ref")
    col_ref2      = cols.get("ref2")
    col_details   = cols.get("details")
    if not all([col_bank_code, col_bank_amt, col_books_amt, col_ref1, col_ref2, col_details]):
        return match_values, stats

    def _num_id(x):
        s = "" if x is None else str(x).strip()
        digits = re.sub(r"\D","", s).lstrip("0")
        return digits or "0"

    bank_code = df[col_bank_code].apply(to_number)
    bank_amt  = df[col_bank_amt].apply(to_number)
    books_amt = df[col_books_amt].apply(to_number)
    details   = df[col_details].astype(str).fillna("")
    ref1      = df[col_ref1].astype(str).fillna("")
    ref2      = df[col_ref2].astype(str).fillna("")

    bank_idx  = [i for i in df.index
                 if pd.notna(bank_code.iat[i]) and int(bank_code.iat[i]) == 493
                 and "שיק" in details.iat[i]
                 and pd.notna(bank_amt.iat[i]) and bank_amt.iat[i] > 0
                 and float(match_values.iat[i] or 0) == 0]

    books_idx = [j for j in df.index
                 if str(ref1.iat[j]).startswith("CH")
                 and "תשלום בהמחאה" in details.iat[j]
                 and pd.notna(books_amt.iat[j]) and books_amt.iat[j] < 0
                 and float(match_values.iat[j] or 0) == 0]

    books_by_id: Dict[str, List[int]] = {}
    for j in books_idx:
        k = _num_id(ref2.iat[j])
        books_by_id.setdefault(k, []).append(j)

    used_books = set()
    for i in bank_idx:
        key = _num_id(ref1.iat[i])
        cands = [j for j in books_by_id.get(key, []) if j not in used_books]
        chosen = None
        for j in cands:
            if abs(float(bank_amt.iat[i]) + float(books_amt.iat[j])) <= AMOUNT_TOL_4:
                chosen = j
                break
        if chosen is not None:
            match_values.iat[i] = 4
            match_values.iat[chosen] = 4
            used_books.add(chosen)
            stats["pairs"] += 1

    return match_values, stats

# ---------------- עיבוד חוברת ----------------
def process_workbook(main_bytes: bytes, aux_bytes: Optional[bytes] = None):
    """מחזיר Bytes של אקסל מעודכן + טבלת סיכום."""
    wb_in = load_workbook(io.BytesIO(main_bytes), data_only=True, read_only=True)

    out_stream = io.BytesIO()
    summary_rows = []
    standing_collect = []

    with pd.ExcelWriter(out_stream, engine="xlsxwriter") as writer:
        for ws in wb_in.worksheets:
            df = ws_to_df(ws)
            df_save = df.copy()
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

            if col_match not in df_save.columns:
                df_save[col_match] = 0
            match_values = df_save[col_match].copy().fillna(0)

            cols_map = {
                "bank_code": col_bank_code,
                "bank_amt":  col_bank_amt,
                "books_amt": col_books_amt,
                "ref":       col_ref1,
                "ref2":      col_ref2,
                "date":      col_date,
                "details":   col_details,
            }

            # כלל 1
            match_values, stats1 = _apply_rule_1(df, cols_map, match_values)
            # כלל 2 + איסוף הוראת קבע
            match_values, stats2, st_rows = _apply_rule_2_and_collect(df, cols_map, match_values, ws.title)
            if not st_rows.empty:
                standing_collect.append(st_rows)
            # כלל 3 – רזרבה
            stats3 = {"note": "not_implemented"}
            # כלל 4
            match_values, stats4 = _apply_rule_4(df, cols_map, match_values)

            df_out = df_save.copy()
            df_out[col_match] = match_values
            df_out.to_excel(writer, index=False, sheet_name=ws.title)

            summary_rows.append({
                "גיליון": ws.title,
                "זוגות שסומנו 1": stats1.get("pairs", 0),
                "שורות שסומנו 2": stats2.get("flagged", 0),
                "התאמה 3": stats3.get("note"),
                "זוגות שסומנו 4": stats4.get("pairs", 0),
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
        # Header
        if ws.max_row >= 1:
            for c in ws[1]:
                c.font = Font(bold=True)
                c.fill = HEADER_BLUE
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = BORDER
        # Body + רוחב עמודות
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.alignment = Alignment(horizontal='right', vertical='center')
                cell.border = BORDER
        for col_idx in range(1, ws.max_column+1):
            col_letter = get_column_letter(col_idx)
            maxlen = 0
            for r in range(1, ws.max_row+1):
                v = ws.cell(row=r, column=col_idx).value
                if v is None: continue
                v = str(v)
                if len(v) > maxlen: maxlen = len(v)
            ws.column_dimensions[col_letter].width = min(maxlen + 2, 60)

    for ws in wb.worksheets:
        style_sheet(ws)

    final = io.BytesIO()
    wb.save(final)
    return final.getvalue(), pd.DataFrame(summary_rows)

# ---------------- ממשק משתמש ----------------
st.set_page_config(page_title="התאמות 1+2+3+4 + הוראת קבע", page_icon="✅", layout="centered")
st.markdown("""
<style>
  html, body, [class*="css"] { direction: rtl; text-align: right; }
  .block-container { padding-top: 1rem; }
</style>
""", unsafe_allow_html=True)

st.title("התאמות לקוחות/בנק – 1+2+3+4 + גיליון הוראת קבע")

# ====== עדכון – כללי VLOOKUP (עם שמירה) ======
st.subheader("⚙️ עדכון – כללי VLOOKUP קבועים ומורחבים (עם שמירה)")
mode = st.radio("סוג עדכון", ["לפי פרטים (שם)", "לפי סכום"], horizontal=True)
colA, colB, colC = st.columns([2,1,1])
if mode == "לפי פרטים (שם)":
    name_key = colA.text_input("שם/פרטים (המפתח לחיפוש)")
    name_val = colB.text_input("מס' ספק")
    if colC.button("הוסף/עדכן שם→ספק"):
        if name_key and name_val:
            st.session_state.name_map[normalize_text(name_key)] = name_val
            save_rules_to_disk(); st.success("נשמר")
else:
    amt_key = colA.text_input("סכום (כמספר או טקסט)")
    amt_val = colB.text_input("מס' ספק")
    if colC.button("הוסף/עדכן סכום→ספק"):
        if amt_key and amt_val:
            st.session_state.amount_map[str(amt_key).strip()] = amt_val
            save_rules_to_disk(); st.success("נשמר")

c1, c2, c3, c4 = st.columns([1,1,1,2])
if c1.button("נקה הכללים"):
    st.session_state.name_map = {}
    st.session_state.amount_map = {}
    save_rules_to_disk(); st.warning("נוקה ונשמר")

c2.download_button(
    "⬇️ ייצוא JSON",
    data=json.dumps({"name_map": st.session_state.name_map, "amount_map": st.session_state.amount_map}, ensure_ascii=False, indent=2).encode("utf-8"),
    file_name="rules_store.json",
    mime="application/json"
)
up_rules = c3.file_uploader("⬆️ ייבוא JSON", type=["json"], label_visibility="collapsed")
if c4.button("ייבוא והחלפה"):
    if up_rules is not None:
        data = json.loads(up_rules.getvalue().decode("utf-8"))
        st.session_state.name_map  = data.get("name_map", {})
        st.session_state.amount_map = data.get("amount_map", {})
        save_rules_to_disk(); st.success("נטען ונשמר")

# --- תצוגת טבלאות כללים (עם מיון מספרי בטוח לסכומים) ---
# שם/פרטים → ספק
_df_name = pd.DataFrame({
    "שם/פרטים": list(st.session_state.name_map.keys()),
    "מס' ספק":  list(st.session_state.name_map.values())
})
st.dataframe(_df_name, use_container_width=True, height=220)

# סכום → ספק (מיון מספרי)
_df_amt = pd.DataFrame({
    "סכום":   list(st.session_state.amount_map.keys()),
    "מס' ספק": list(st.session_state.amount_map.values())
})
_df_amt["_סכום_מספרי"] = pd.to_numeric(_df_amt["סכום"].astype(str).str.replace(",", ""), errors="coerce")
_df_amt = _df_amt.sort_values(by="_סכום_מספרי").drop(columns="_סכום_מספרי")
st.dataframe(_df_amt, use_container_width=True, height=220)

st.divider()

# ====== העלאה והרצה ======
cx1, cx2 = st.columns([2,2])
main_file = cx1.file_uploader("בחרי קובץ אקסל מקור (xlsx)", type=["xlsx"])
aux_file  = cx2.file_uploader("(אופציונלי) קובץ עזר להעברות/מיפויים", type=["xlsx"], help="לא חובה. משמש לעזר בלבד.")

run = st.button("▶️ הרצה – הפקת קובץ מסומן + גיליון הוראת קבע")
if run and main_file is not None:
    main_bytes = main_file.read()
    aux_bytes  = aux_file.read() if aux_file is not None else None
    out_bytes, summary_df = process_workbook(main_bytes, aux_bytes)

    st.success("הקובץ עובד! אפשר להוריד.")
    st.download_button(
        "⬇️ הורדה – תוצאה סופית (1+2+3+4 + הוראת קבע)",
        data=out_bytes,
        file_name="תוצאה_סופית_התאמות_1_2_3_4_והוראת_קבע.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    with st.expander("תצוגת סיכום"):
        st.dataframe(summary_df, use_container_width=True)
else:
    st.info("בחרי קובץ מקור ולחצי הרצה.")
