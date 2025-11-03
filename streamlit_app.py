from __future__ import annotations
# ==== תוספות עבור התאמה 4 (למעלה, באזור ההגדרות) ====
AMOUNT_TOL_4 = 0.20  # טולרנס הפרש סכומים בשיק ספקים

def _only_digits(s: str) -> str:
    """החזרת ספרות בלבד (וללא אפסים מובילים)."""
    s = "" if s is None else str(s)
    d = re.sub(r"\D", "", s).lstrip("0")
    return d or "0"
# ==== סוף תוספות חלק 1 ====
# ==== תוספות עבור התאמה 4: פונקציית סימון ====
def _apply_rule_4(df: pd.DataFrame,
                  col_match: str,
                  col_bank_code: str,
                  col_bank_amt: str,
                  col_books_amt: str,
                  col_ref1: str,
                  col_ref2: str,
                  col_details: str) -> Tuple[pd.Series, dict]:
    """
    התאמה 4 – שיקי ספקים:
    תנאים:
      צד בנק:   קוד פעולה = 493, 'שיק' בפרטים, סכום בדף > 0
      צד ספרים: אסמכתא 1 מתחיל 'CH', 'תשלום בהמחאה' בפרטים, סכום בספרים < 0
      התאמה לפי:  אסמכתא 2 (ספרים) == אסמכתא 1 (בנק) לאחר נירמול ספרות בלבד
      תאריכים:    לא בודקים
      טולרנס:     |סכום בנק + סכום ספרים| <= AMOUNT_TOL_4
    """
    stats = {"pairs": 0}

    # אם חסר עמודה כלשהי – אין כלל 4
    needed = [col_match, col_bank_code, col_bank_amt, col_books_amt, col_ref1, col_ref2, col_details]
    if any(c is None or c not in df.columns for c in needed):
        return df[col_match].fillna(0), stats

    # וקטורים
    v_match   = df[col_match].fillna(0).copy()
    v_bcode   = df[col_bank_code].apply(lambda x: np.nan if pd.isna(x) else int(float(str(x).replace(",",""))))
    v_bamt    = df[col_bank_amt].apply(to_number)
    v_bdet    = df[col_details].astype(str).fillna("")
    v_ref1    = df[col_ref1].astype(str).fillna("")

    v_bksamt  = df[col_books_amt].apply(to_number)
    v_bksdet  = df[col_details].astype(str).fillna("")
    v_ref2    = df[col_ref2].astype(str).fillna("")
    v_ref1_norm = v_ref1.apply(_only_digits)
    v_ref2_norm = v_ref2.apply(_only_digits)

    # אינדקסים רלוונטיים
    bank_idx = [i for i in df.index
                if float(v_match.iat[i] or 0) == 0
                and pd.notna(v_bcode.iat[i]) and v_bcode.iat[i] == 493
                and "שיק" in v_bdet.iat[i]
                and pd.notna(v_bamt.iat[i]) and v_bamt.iat[i] > 0]

    books_idx = [j for j in df.index
                 if float(v_match.iat[j] or 0) == 0
                 and str(df[col_ref1].iat[j]).startswith("CH")
                 and "תשלום בהמחאה" in v_bksdet.iat[j]
                 and pd.notna(v_bksamt.iat[j]) and v_bksamt.iat[j] < 0]

    # קיבוץ ספרים לפי מזהה (אסמכתא 2 מנורמלת)
    books_by_id = {}
    for j in books_idx:
        key = v_ref2_norm.iat[j]
        books_by_id.setdefault(key, []).append(j)

    used_books = set()
    for i in bank_idx:
        key = v_ref1_norm.iat[i]
        candidates = [j for j in books_by_id.get(key, []) if j not in used_books]
        chosen = None
        for j in candidates:
            if abs(float(v_bamt.iat[i]) + float(v_bksamt.iat[j])) <= AMOUNT_TOL_4:
                chosen = j
                break
        if chosen is not None:
            v_match.iat[i] = 4
            v_match.iat[chosen] = 4
            used_books.add(chosen)
            stats["pairs"] += 1

    return v_match, stats
# ==== סוף תוספות חלק 2 ====
# ==== הוספת קריאה לכלל 4 בתוך process_workbook ====
# match_values כבר קיים אצלך מהכללים הקודמים:
df_out = df.copy()
if col_match not in df_out.columns:
    df_out[col_match] = 0
df_out[col_match] = df_out[col_match].fillna(0)

# כלל 4 – שיקי ספקים
match_after_4, stats4 = _apply_rule_4(
    df_out, 
    col_match=col_match,
    col_bank_code=col_bank_code,
    col_bank_amt=col_bank_amt,
    col_books_amt=col_books_amt,
    col_ref1=col_ref1,
    col_ref2=col_ref2,
    col_details=col_details
)

# שלא לדרוס מס' התאמה שכבר קיים (A2): רק אם 0 → נעדכן ל-4
mask_zero = (df_out[col_match].fillna(0).astype(float) == 0) & (match_after_4.astype(float) == 4)
df_out.loc[mask_zero, col_match] = 4

# (אופציונלי) לעדכן סטטיסטיקה/סיכום אם יש לך טבלת summary_rows:
# summary_row["זוגות שסומנו 4"] = stats4.get("pairs", 0)
# ==== סוף תוספות חלק 3 ====
