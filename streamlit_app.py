# streamlit_app.py
# -*- coding: utf-8 -*-
"""
אפליקציית Streamlit לסימון התאמות 1+2+3+4 בקבצי אקסל (לקוחות/בנקים),
כולל הפקת גיליון "הוראת קבע ספקים", ופריסת עמוד/RTL.

כללים (ברירת מחדל – לא דורסת ערכים קיימים בעמודה "מס. התאמה"):
1. התאמה 1 – OV/RC קפדנית 1:1 לפי (סכום מוחלט, תאריך זהה) בצד ספרים (+) ↔ בצד בנק (−)
   רק כאשר אסמכתא 1 בצד ספרים מתחילה OV/RC. מסמן 1 בשתי השורות.
2. התאמה 2 – הוראות קבע: מסמן 2 על שורות בנק עם קודי פעולה מוגדרים (ברירת מחדל: {469, 515}).
3. התאמה 3 – העברות/סימון עזר: אפשר להרחיב לפי צורך; כאן נשאר כבסיס (לא מסמן דבר כברירת מחדל).
4. התאמה 4 – שיקים ספקים: התאמת זוגות בין צד בנק וצד ספרים לפי הכללים:
   בנק: קוד פעולה 493, פרטים מכילים "שיק", סכום בדף > 0.
   ספרים: אסמכתא 1 מתחילה CH, פרטים "תשלום בהמחאה", סכום בספרים < 0.
   זיהוי באמצעות אסמכתא 2 בספרים = אסמכתא 1 בבנק (בהשוואה מנורמלת ספרות בלבד),
   וטולרנס סכום עד 0.20 ₪. אין דרישת התאמת תאריכים. מסמן 4 בשתי השורות.

גיליון "הוראת קבע ספקים": נאסף מכל הגיליונות ש"מס. התאמה"=2, עם עיצוב בסיסי ו-RTL.

הערות:
- ניתן לשנות סט קודי פעולה להוראת קבע (STANDING_CODES) למטה לפי הבנק שלך.
- העמודות מזוהות אוטומטית לפי שמות נפוצים בעברית. ניתן להרחיב את הרשימות.
- בכל הכללים ההתנהגות היא A2: מסמנים רק אם "מס. התאמה" ריק/0.
"""

import io
import re
from datetime import datetime
from typing import Dict, List, Tuple, Optional

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.page import PageMargins

# ---------------- הגדרות וכלי עזר ----------------
# קודי פעולה להוראות קבע בבנק (ניתן להתאים)
STANDING_CODES = {469, 515}

# טולרנס להפרש סכומים
AMOUNT_TOL = 0.20

# מועמדים לשמות עמודות (תמיכה בוואריאציות)
MATCH_COL_CANDS   = ["מס. התאמה", "מס התאמה", "התאמה", "מס' התאמה"]
BANK_CODE_CANDS   = ["קוד פעולת בנק", "קוד פעולה בנק", "קוד פעולה", "קוד בנק"]
BANK_AMT_CANDS    = ["סכום בדף", "סכום בבנק", "סכום בנק"]
BOOKS_AMT_CANDS   = ["סכום בספרים", "ספרים סכום", "סכום ספרים"]
REF1_CANDS        = ["אסמכתא 1", "אסמכתא", "מס' אסמכתא", "Reference 1"]
REF2_CANDS        = ["אסמכתא 2", "אסמכתא נוספת", "Reference 2"]
DATE_CANDS        = ["תאריך מאזן", "תאריך", "ת. מאזן", "תאריך מסמך"]
DATE_VAL_CANDS    = ["ת. ערך", "תאריך ערך", "תאריך ערך בנק"]
DETAILS_CANDS     = ["פרטים", "תיאור", "תאור", "Details"]
SHEETDATE_CANDS   = ["תאריך הדף", "תאריך דף"]
CARD_CANDS        = ["קוד חש' בנק/אשראי", "קוד חשבון", "קוד בנק", "מס' חשבון"]

HEADER_BLUE = PatternFill(start_color="9BC2E6", end_color="9BC2E6", fill_type="solid")
THIN = Side(style='thin', color='CCCCCC')
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)


def normalize_text(s: str) -> str:
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\u200f", "").replace("\u200e", "").strip()
    return s


def to_number(x) -> float:
    if pd.isna(x):
        return np.nan
    s = str(x).replace(",", "")
    try:
        return float(s)
    except Exception:
        m = re.findall(r"[-+]?\d+(?:\.\d+)?", s)
        return float(m[0]) if m else np.nan


def normalize_date(s: pd.Series) -> pd.Series:
    # תאריך בלי שעה
    out = pd.to_datetime(s, errors="coerce")
    return pd.to_datetime(out.dt.date)


def ref_starts_with_ov_rc(ref: str) -> bool:
    if not isinstance(ref, str):
        ref = str(ref) if ref is not None else ""
    ref = ref.strip().upper()
    return ref.startswith("OV") or ref.startswith("RC")


def exact_or_contains(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    # קודם התאמה מדויקת לשם עמודה, אם לא – חיפוש לפי הכללה
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
    df = pd.DataFrame(list(data), columns=headers)
    return df

# ---------------- לוגיקת עיבוד מרכזית ----------------

def _apply_rule_1(df: pd.DataFrame, cols: Dict[str, str], match_values: pd.Series) -> Tuple[pd.Series, Dict[str, int]]:
    """התאמה 1 – OV/RC קפדנית 1:1 לפי (abs סכום, תאריך זהה) – A2: רק כשהתאמה ריקה/0"""
    stats = {"pairs": 0}
    col_bank_code = cols.get("bank_code")
    col_bank_amt  = cols.get("bank_amt")
    col_books_amt = cols.get("books_amt")
    col_ref       = cols.get("ref1")
    col_date      = cols.get("date")
    if not all([col_bank_code, col_bank_amt, col_books_amt, col_ref, col_date]):
        return match_values, stats

    _date      = normalize_date(pd.to_datetime(df[col_date], errors="coerce"))
    _bank_amt  = df[col_bank_amt].apply(to_number)
    _books_amt = df[col_books_amt].apply(to_number)
    _ref       = df[col_ref].astype(str).fillna("")

    # מועמדי ספרים: סכום חיובי, OV/RC, תאריך קיים
    books_candidates = [
        j for j in range(len(df))
        if pd.notna(_books_amt.iat[j]) and _books_amt.iat[j] > 0
        and pd.notna(_date.iat[j]) and ref_starts_with_ov_rc(_ref.iat[j])
    ]
    # קבוצות לפי (abs סכום, תאריך)
    bank_keys, books_keys = {}, {}
    for i in range(len(df)):
        a = _bank_amt.iat[i]
        d = _date.iat[i]
        if pd.notna(a) and pd.notna(d):
            bank_keys.setdefault((abs(a), d), []).append(i)
    for j in books_candidates:
        a = _books_amt.iat[j]
        d = _date.iat[j]
        if pd.notna(a) and pd.notna(d):
            books_keys.setdefault((abs(a), d), []).append(j)

    # התאמה קפדנית – רק 1:1 בכל צד
    for k, b_idx in bank_keys.items():
        if len(b_idx) == 1 and len(books_keys.get(k, [])) == 1:
            i = b_idx[0]; j = books_keys[k][0]
            if float(match_values.iat[i] or 0) == 0 and float(match_values.iat[j] or 0) == 0:
                match_values.iat[i] = 1
                match_values.iat[j] = 1
                stats["pairs"] += 1
    return match_values, stats


def _apply_rule_2(df: pd.DataFrame, cols: Dict[str, str], match_values: pd.Series) -> Tuple[pd.Series, Dict[str, int], pd.DataFrame]:
    """התאמה 2 – הוראות קבע: סימון לפי קוד פעולה בנק ובמידת הצורך פרטים. A2 – רק כשהתאמה ריקה/0"""
    stats = {"flagged": 0}
    col_bank_code = cols.get("bank_code")
    col_bank_amt  = cols.get("bank_amt")
    col_details   = cols.get("details")

    standing_rows = []
    if not all([col_bank_code, col_details, col_bank_amt]):
        return match_values, stats, pd.DataFrame(standing_rows)

    _bank_code = df[col_bank_code].apply(to_number)
    _details   = df[col_details].astype(str).fillna("")
    _bank_amt  = df[col_bank_amt].apply(to_number)

    for i in range(len(df)):
        code = _bank_code.iat[i]
        if pd.isna(code):
            continue
        if int(code) in STANDING_CODES:
            if float(match_values.iat[i] or 0) == 0:
                match_values.iat[i] = 2
                stats["flagged"] += 1
                standing_rows.append({
                    "פרטים": _details.iat[i],
                    "סכום": float(_bank_amt.iat[i]) if pd.notna(_bank_amt.iat[i]) else np.nan
                })
    return match_values, stats, pd.DataFrame(standing_rows)


def _apply_rule_4(df: pd.DataFrame, cols: Dict[str, str], match_values: pd.Series) -> Tuple[pd.Series, Dict[str, int]]:
    """התאמה 4 – שיקים ספקים: התאמת בנק↔ספרים לפי אסמכתאות מנורמלות וסכום הפוך. A2 – רק כשהתאמה ריקה/0"""
    stats = {"pairs": 0}
    col_bank_code = cols.get("bank_code")
    col_bank_amt  = cols.get("bank_amt")
    col_books_amt = cols.get("books_amt")
    col_ref1      = cols.get("ref1")
    col_ref2      = cols.get("ref2")
    col_details   = cols.get("details")

    if not all([col_bank_code, col_bank_amt, col_books_amt, col_ref1, col_ref2, col_details]):
        return match_values, stats

    def norm_id(x):
        if x is None: return None
        s = str(x).strip()
        digits = re.sub(r"\D", "", s).lstrip('0')
        return digits or '0'

    bank_code = df[col_bank_code].apply(to_number)
    details   = df[col_details].astype(str).fillna("")
    bank_amt  = df[col_bank_amt].apply(to_number)
    books_amt = df[col_books_amt].apply(to_number)
    ref1      = df[col_ref1].astype(str).fillna("")
    ref2      = df[col_ref2].astype(str).fillna("")

    # אינדקסים
    bank_idx  = [i for i in range(len(df)) if (not pd.isna(bank_code.iat[i]) and int(bank_code.iat[i]) == 493
                                              and "שיק" in details.iat[i]
                                              and pd.notna(bank_amt.iat[i]) and bank_amt.iat[i] > 0
                                              and float(match_values.iat[i] or 0) == 0)]
    books_idx = [j for j in range(len(df)) if (str(ref1.iat[j]).startswith("CH")
                                              and "תשלום בהמחאה" in details.iat[j]
                                              and pd.notna(books_amt.iat[j]) and books_amt.iat[j] < 0
                                              and float(match_values.iat[j] or 0) == 0)]

    # מפה מאסמכתא2 מנורמלת לשורות ספרים
    book_by_id: Dict[str, List[int]] = {}
    for j in books_idx:
        k = norm_id(ref2.iat[j])
        book_by_id.setdefault(k, []).append(j)

    used_books: set = set()
    for i in bank_idx:
        key = norm_id(ref1.iat[i])
        candidates = [j for j in book_by_id.get(key, []) if j not in used_books]
        chosen = None
        for j in candidates:
            if abs(bank_amt.iat[i] + books_amt.iat[j]) <= AMOUNT_TOL:
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
    """מעבד את קובץ המקור + (אופציונלי) קובץ עזר ומחזיר Bytes של אקסל מעודכן + טבלת סיכום.
    """
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

            # זיהוי עמודות
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
            match_values = df_save[col_match].copy()
            match_values = match_values.fillna(0)

            cols_map = {
                "bank_code": col_bank_code,
                "bank_amt":  col_bank_amt,
                "books_amt": col_books_amt,
                "ref1":      col_ref1,
                "ref2":      col_ref2,
                "date":      col_date,
                "details":   col_details,
            }

            # --- התאמה 1 ---
            match_values, stats1 = _apply_rule_1(df, cols_map, match_values)
            # --- התאמה 2 ---
            match_values, stats2, st_rows = _apply_rule_2(df, cols_map, match_values)
            if not st_rows.empty:
                st_rows.insert(0, "גיליון מקור", ws.title)
                standing_collect.append(st_rows)
            # --- התאמה 3 --- (מקום להרחבה – כרגע לא מסמן)
            stats3 = {"note": "not_implemented"}
            # --- התאמה 4 ---
            match_values, stats4 = _apply_rule_4(df, cols_map, match_values)

            # כתיבה
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

        # --- גיליון הוראת קבע ספקים (מריכוז של 2) ---
        if standing_collect:
            st_df = pd.concat(standing_collect, ignore_index=True)
        else:
            st_df = pd.DataFrame([{"גיליון מקור": "—", "הודעה": "אין שורות עם מס. התאמה = 2 (הוראת קבע)"}])

        # עיצוב בסיסי ו-RTL
        st_df.to_excel(writer, index=False, sheet_name="הוראת קבע ספקים")

        # --- גיליון סיכום ---
        pd.DataFrame(summary_rows).to_excel(writer, index=False, sheet_name="סיכום")

    # החלת RTL, פריסת עמוד והדגשות על חוברת ה-XLSXWriter
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
        # כותרת מודגשת
        if ws.max_row >= 1:
            for c in ws[1]:
                c.font = Font(bold=True)
                c.fill = HEADER_BLUE
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = BORDER
        # גוף
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
            for cell in row:
                cell.alignment = Alignment(horizontal='right', vertical='center')
                cell.border = BORDER
        # רוחב עמודות אוטומטי (תקרה 60)
        for col_idx in range(1, ws.max_column+1):
            col_letter = get_column_letter(col_idx)
            maxlen = 0
            for r in range(1, ws.max_row+1):
                v = ws.cell(row=r, column=col_idx).value
                if v is None:
                    continue
                v = str(v)
                if len(v) > maxlen:
                    maxlen = len(v)
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
  html, body, [class*=\"css\"] { direction: rtl; text-align: right; }
  .block-container { padding-top: 1rem; }
</style>
""", unsafe_allow_html=True)

st.title("התאמות לקוחות/בנק – 1+2+3+4 + גיליון הוראת קבע")

# ====== עדכון ושמירת כללי VLOOKUP (קבועים ומורחבים) ======
import json, os
if "name_map" not in st.session_state:
    st.session_state.name_map = {}
if "amount_map" not in st.session_state:
    st.session_state.amount_map = {}

RULES_PATH = "rules_store.json"

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

with st.expander("⚙️ עדכון – כללי VLOOKUP קבועים ומורחבים (עם שמירה)", expanded=False):
    st.write("עדכון לפי **פרטים (שם)** או לפי **סכום**. נשמר ל־`rules_store.json` לשימוש חוזר.")
    mode = st.radio("סוג עדכון", ["לפי פרטים (שם)", "לפי סכום"], horizontal=True)
    if mode == "לפי פרטים (שם)":
        colN1, colN2, colN3 = st.columns([2,1,1])
        name_key = colN1.text_input("שם/פרטים (המפתח לחיפוש)")
        name_val = colN2.text_input("מס' ספק")
        if colN3.button("הוסף/עדכן שם→ספק"):
            if name_key and name_val:
                st.session_state.name_map[name_key] = name_val
                save_rules_to_disk(); st.success("נשמר")
    else:
        colA1, colA2, colA3 = st.columns([1.5,1,1])
        amt_key = colA1.text_input("סכום (כמספר או טקסט)")
        amt_val = colA2.text_input("מס' ספק")
        if colA3.button("הוסף/עדכן סכום→ספק"):
            if amt_key and amt_val:
                st.session_state.amount_map[amt_key] = amt_val
                save_rules_to_disk(); st.success("נשמר")

    st.divider()
    c1, c2, c3, c4 = st.columns([1,1,1,2])
    c1.download_button("⬇️ ייצוא JSON", data=json.dumps({
                        "name_map": st.session_state.name_map,
                        "amount_map": st.session_state.amount_map
                    }, ensure_ascii=False, indent=2).encode("utf-8"),
                    file_name="rules_store.json", mime="application/json")
    uploaded_rules = c2.file_uploader("⬆️ ייבוא JSON", type=["json"], label_visibility="collapsed")
    if c3.button("ייבוא והחלפה"):
        if uploaded_rules is not None:
            data = json.loads(uploaded_rules.getvalue().decode("utf-8"))
            st.session_state.name_map = data.get("name_map", {})
            st.session_state.amount_map = data.get("amount_map", {})
            save_rules_to_disk(); st.success("נטען ונשמר")
    if c4.button("נקה הכללים"):
        st.session_state.name_map = {}
        st.session_state.amount_map = {}
        save_rules_to_disk(); st.warning("נוקה ונשמר")

    # תצוגה
    import pandas as _pd
    st.dataframe(_pd.DataFrame({"שם/פרטים": list(st.session_state.name_map.keys()),
                                "מס' ספק": list(st.session_state.name_map.values())}),
                 use_container_width=True, height=220)
    st.dataframe(_pd.DataFrame({"סכום": list(st.session_state.amount_map.keys()),
                                "מס' ספק": list(st.session_state.amount_map.values())}).sort_values("סכום", errors="ignore"),
                 use_container_width=True, height=220)

# ====== העלאה והרצה ======
c1, c2 = st.columns([2,2])
main_file = c1.file_uploader("בחרי קובץ אקסל מקור (xlsx)", type=["xlsx"])
aux_file  = c2.file_uploader("(אופציונלי) קובץ עזר להעברות/מיפויים", type=["xlsx"], help="לא חובה. משמש לעזר בלבד.")

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
