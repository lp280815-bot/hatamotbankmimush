# streamlit_app.py
# -*- coding: utf-8 -*-
import io, re
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

# ===== הגדרות UI =====
st.set_page_config(page_title="התאמה 4 (A2) – שיקי ספקים", page_icon="✅", layout="centered")
st.markdown("""
<style>
  html, body, [class*="css"] { direction: rtl; text-align: right; }
  .block-container { padding-top: 1rem; }
</style>
""", unsafe_allow_html=True)
st.title("התאמה מס' 4 (A2) – שיקי ספקים, ללא פשרות")
st.caption("מיישם רק את כלל 4 במצב A2 על קובץ אקסל קיים • שומר את כל הגיליונות • RTL + A4 לרוחב")

# ===== פרמטרים לכלל 4 =====
CHECK_CODE   = 493      # קוד פעולה בנק לשיקי ספקים
AMOUNT_TOL   = 0.20     # סטייה מותרת בסכום
BANK_TEXT_PATTERNS  = (r"\bשיק\b", )                 # מילת מפתח בפרטים בצד הבנק
BOOKS_TEXT_PATTERNS = (r"תשלום\s*בהמחאה", r"\bהמחאה\b")  # בצד הספרים

# מועמדי שמות עמודות (נזהה חכם גם אם יש וריאציות)
MATCH_COL_CANDS = ["מס. התאמה","מס התאמה","מספר התאמה","התאמה","מס' התאמה"]
BANK_CODE_CANDS = ["קוד פעולת בנק","קוד פעולה","קוד פעולת","קוד בנק"]
BANK_AMT_CANDS  = ["סכום בדף","סכום בבנק","סכום בנק","סכום תנועת בנק"]
BOOKS_AMT_CANDS = ["סכום בספרים","סכום ספרים","סכום בספר"]
REF1_CANDS      = ["אסמכתא 1","אסמכתא","אסמכתה","Reference 1"]
REF2_CANDS      = ["אסמכתא 2","אסמכתא2","אסמכתא נוספת","Reference 2"]
DETAILS_CANDS   = ["פרטים","תיאור","שם ספק","Details"]

HEADER_BLUE = PatternFill(start_color="9BC2E6", end_color="9BC2E6", fill_type="solid")
THIN = Side(style='thin', color='CCCCCC')
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

# ===== עזר =====
def _safe_str(x): return "" if x is None else str(x)

def _only_digits(x: str) -> str:
    """ספרות בלבד ללא אפסים מובילים (לזיהוי שיק)."""
    s = re.sub(r"\D", "", _safe_str(x)).lstrip("0")
    return s or "0"

def _to_number(x):
    if x is None or (isinstance(x, float) and np.isnan(x)): 
        return np.nan
    s = _safe_str(x).replace(",", "").replace("₪", "").strip()
    m = re.findall(r"[-+]?\d+(?:\.\d+)?", s)
    try:
        return float(m[0]) if m else np.nan
    except Exception:
        return np.nan

def _contains_any(text: str, patterns) -> bool:
    t = _safe_str(text)
    for p in patterns:
        if re.search(p, t):
            return True
    return False

def exact_or_contains(df: pd.DataFrame, cands) -> str|None:
    cols = list(map(str, df.columns))
    for c in cands:
        if c in cols: return c
    for c in cands:
        for col in cols:
            if c in col: return col
    return None

def ws_to_df(ws) -> pd.DataFrame:
    rows = list(ws.iter_rows(values_only=True))
    if not rows: return pd.DataFrame()
    header = [(_safe_str(x)) for x in rows[0]]
    return pd.DataFrame(rows[1:], columns=header)

def ensure_rtl_page(ws):
    try:
        ws.sheet_view.rightToLeft = True
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_margins = PageMargins(left=0.3, right=0.3, top=0.5, bottom=0.5, header=0.3, footer=0.3)
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
    except Exception:
        pass

# ===== כלל 4 – A2 בלבד =====
def apply_rule_4_on_sheet(df: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """מקבל DataFrame של גיליון אחד, ומחזיר DF אחרי סימון התאמה מס' 4 + סטטיסטיקה."""
    stats = {"bank_candidates": 0, "books_candidates": 0, "pairs": 0}

    if df.empty: return df, stats

    # מיפוי עמודות
    col_match     = exact_or_contains(df, MATCH_COL_CANDS) or "מס. התאמה"
    if col_match not in df.columns:
        df[col_match] = 0

    col_bank_code = exact_or_contains(df, BANK_CODE_CANDS)
    col_bank_amt  = exact_or_contains(df, BANK_AMT_CANDS)
    col_books_amt = exact_or_contains(df, BOOKS_AMT_CANDS)
    col_ref1      = exact_or_contains(df, REF1_CANDS)
    col_ref2      = exact_or_contains(df, REF2_CANDS)
    col_det       = exact_or_contains(df, DETAILS_CANDS)

    needed = [col_bank_code, col_bank_amt, col_books_amt, col_ref1, col_ref2, col_det]
    if any(c is None for c in needed):
        return df, stats  # אין כל העמודות הנדרשות – לא נוגעים

    v_match = pd.to_numeric(df[col_match], errors="coerce").fillna(0)
    v_code  = df[col_bank_code].apply(_to_number)
    v_bamt  = df[col_bank_amt].apply(_to_number)
    v_samt  = df[col_books_amt].apply(_to_number)
    v_det   = df[col_det].astype(str).fillna("")
    v_r1    = df[col_ref1].astype(str).fillna("")
    v_r2    = df[col_ref2].astype(str).fillna("")

    # נרמול מזהי שיק
    r1_digits = v_r1.apply(_only_digits)  # בנק
    r2_digits = v_r2.apply(_only_digits)  # ספרים
    r2_fallback = v_r1.apply(lambda s: _only_digits(_safe_str(s).replace("CH","").replace("ch","")))
    r2_final = r2_digits.where(r2_digits.ne("0"), r2_fallback)

    # מועמדים
    bank_idx = [
        i for i in df.index
        if float(v_match.iat[i]) == 0
           and pd.notna(v_code.iat[i]) and int(v_code.iat[i]) == CHECK_CODE
           and _contains_any(v_det.iat[i], BANK_TEXT_PATTERNS)
           and pd.notna(v_bamt.iat[i]) and v_bamt.iat[i] > 0
    ]
    books_idx = [
        j for j in df.index
        if float(v_match.iat[j]) == 0
           and _safe_str(v_r1.iat[j]).upper().startswith("CH")
           and _contains_any(v_det.iat[j], BOOKS_TEXT_PATTERNS)
           and pd.notna(v_samt.iat[j]) and v_samt.iat[j] < 0
    ]
    stats["bank_candidates"]  = len(bank_idx)
    stats["books_candidates"] = len(books_idx)

    # קיבוץ ספרים לפי מזהה
    books_by_id = {}
    for j in books_idx:
        books_by_id.setdefault(r2_final.iat[j], []).append(j)

    used_books = set()
    for i in bank_idx:
        key = r1_digits.iat[i]
        candidates = [j for j in books_by_id.get(key, []) if j not in used_books]
        chosen = None
        for j in candidates:
            if abs(float(v_bamt.iat[i]) + float(v_samt.iat[j])) <= AMOUNT_TOL:
                chosen = j; break
        if chosen is not None:
            v_match.iat[i] = 4
            v_match.iat[chosen] = 4
            used_books.add(chosen)
            stats["pairs"] += 1

    # כתיבה חזרה ל-DF
    out = df.copy()
    out[col_match] = v_match
    return out, stats

def apply_rule_4_workbook(xls_bytes: bytes) -> tuple[bytes, pd.DataFrame]:
    """מריץ התאמה 4 (A2) על כל הגיליונות ושומר את כולם ללא מחיקה/שינוי שמות."""
    wb = load_workbook(io.BytesIO(xls_bytes))
    summary = []

    # מעבר על כל הגיליונות
    for ws in wb.worksheets:
        df = ws_to_df(ws)
        df_out, s = apply_rule_4_on_sheet(df)
        # החלפת תוכן הגיליון במלואו (שומר את השם והסדר)
        # מחיקה של כל השורות למעט שורת כותרת:
        for _ in range(ws.max_row - 1):
            ws.delete_rows(2, 1)
        # כתיבה חזרה
        if df_out.empty:
            pass
        else:
            # ודא שיש כותרת
            ws.cell(row=1, column=1).value = df_out.columns[0]
            for j, col in enumerate(df_out.columns, start=1):
                ws.cell(row=1, column=j).value = col
            for r in df_out.itertuples(index=False):
                ws.append(list(r))

        ensure_rtl_page(ws)
        summary.append({"גיליון": ws.title,
                        "מועמדי בנק": s["bank_candidates"],
                        "מועמדי ספרים": s["books_candidates"],
                        "זוגות שסומנו 4": s["pairs"]})

    # שמירה לבייטס
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), pd.DataFrame(summary)

# ===== UI – העלאה/הרצה/הורדה =====
uploaded = st.file_uploader("בחרי קובץ אקסל קיים (xlsx) – עליו ניישם התאמה 4 (A2)", type=["xlsx"])

if st.button("▶️ החלת התאמה 4 (A2) ושמירה") and uploaded is not None:
    with st.spinner("מריץ התאמה 4 (A2) על כל הגיליונות..."):
        out_bytes, summary_df = apply_rule_4_workbook(uploaded.read())
    st.success("מוכן! אפשר להוריד את הקובץ המעודכן.")
    st.dataframe(summary_df, use_container_width=True)
    st.download_button(
        "⬇️ תוצאה סופית - לפי הקוד + התאמה 4 (A2).xlsx",
        data=out_bytes,
        file_name="תוצאה סופית - לפי הקוד + התאמה 4 (A2).xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
else:
    st.info("העלאי קובץ מקור ולחצי על הכפתור כדי להפעיל את כלל 4 (A2).")
