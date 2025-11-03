# match_excel.py
# -*- coding: utf-8 -*-
import argparse, io, json, math, os, re
from datetime import datetime
from collections import defaultdict

import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font

# ---------------------- קבועים ----------------------
MATCH_COL_CANDS   = ["מס.התאמה","מס. התאמה","מס התאמה","מספר התאמה","התאמה"]
BANK_CODE_CANDS   = ["קוד פעולת בנק","קוד פעולה","קוד פעולת"]
BANK_AMT_CANDS    = ["סכום בדף","סכום דף","סכום בבנק","סכום תנועת בנק"]
BOOKS_AMT_CANDS   = ["סכום בספרים","סכום בספר","סכום ספרים"]
REF_CANDS         = ["אסמכתא 1","אסמכתא1","אסמכתא","אסמכתה"]
DATE_CANDS        = ["תאריך מאזן","תאריך ערך","תאריך"]
DETAILS_CANDS     = ["פרטים","תיאור","שם ספק"]

# מיפוי קבועים לשמות -> מס' ספק
NAME_TO_SUPPLIER = {
    "בזק בינלאומי ב": 30006, "פרי ירוחם חב'": 34714, "סלקום ישראל בע": 30055,
    "בזק-הוראות קבע": 34746, "דרך ארץ הייוי": 34602, "גלובס פבלישר ע": 30067,
    "פלאפון תקשורת": 30030, "מרכז הכוכביות": 30002, "ע.אשדוד-מסים": 30056,
    "א.ש.א(בס\"ד)אחז": 30050, "או.פי.ג'י(מ.כ)": 30047, "רשות האכיפה וה": "67-1",
    "קול ביז מילניו": 30053, "פריוריטי סופטו": 30097, "אינטרנט רימון": 34636,
    "עו\"דכנית בע\"מ": 30018, "עיריית רמת גן": 30065, "פז חברת נפט בע": 34811,
    "ישראכרט": 28002, "חברת החשמל ליש": 30015, "הפניקס ביטוח": 34686,
    "מימון ישיר מקב": 34002, "שלמה טפר": 30247, "נמרוד תבור עורך-דין": 30038,
    "עיריית בית שמש": 34805, "פז קמעונאות וא": 34811, "הו\"ק הלו' רבית": 8004,
    "הו\"ק הלוואה קרן": 23001,
    # תוספות מפורשות
    "עיריית אשדוד": 30056, "ישראכרט מור": 34002,
}
# סכומים -> מס' ספק (VLOOKUP לפי סכום מוחלט)
AMOUNT_TO_SUPPLIER = { 8520.0: 30247, 10307.3: 30038 }

ORANGE = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

# ---------------------- פונקציות עזר ----------------------
def normalize_text(s: str) -> str:
    if s is None:
        return ""
    t = str(s)
    t = t.replace("’","'").replace("`","'")
    t = t.replace("״","").replace("„","").replace('"', "")
    t = t.replace("–","-").replace("־","-")
    t = t.replace("\u200f","").replace("\u200e","")
    t = re.sub(r"\s+", " ", t).strip()
    return t

def normalize_for_contains(s: str) -> str:
    s = normalize_text(s).lower()
    return s.replace("-", "").replace(" ", "").replace("'", "")

def to_number(val):
    if val is None: return math.nan
    s = str(val).replace(",", "").replace("₪", "").strip()
    try: return float(s)
    except: return math.nan

def parse_date(val):
    if val is None or val == "": return None
    if isinstance(val, pd.Timestamp): return val
    try:
        return pd.to_datetime(val, dayfirst=True, errors="coerce")
    except:
        return pd.NaT

def find_header_row(ws):
    for i in range(1, ws.max_row+1):
        row_vals = [c.value for c in ws[i]]
        if any(v is not None for v in row_vals):
            return i, [str(v).strip() if v is not None else "" for v in row_vals]
    return 1, [c.value for c in ws[1]]

def find_col(headers, candidates):
    # חיפוש מדויק
    for cand in candidates:
        if cand in headers: return headers.index(cand) + 1
    # חיפוש 'מכיל'
    for i, h in enumerate(headers):
        if not isinstance(h, str): continue
        for cand in candidates:
            if cand in h: return i + 1
    return None

def startswith_ov_rc(v):
    t = (str(v) if v is not None else "").strip().upper()
    return t.startswith("OV") or t.startswith("RC")

# ---------------------- קריאת קובץ עזר לקיבוץ (תאריך+זמן) ----------------------
def build_aux_amount_to_paynums(aux_path):
    wb = load_workbook(aux_path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    head_row, headers = find_header_row(ws)

    def col_idx(cands):
        idx = find_col(headers, cands if isinstance(cands, list) else [cands])
        return idx

    c_date = col_idx(["תאריך פריקה","תאריך"])
    c_amt  = col_idx(["אחרי ניכוי","אחרי ניכוי מס","סכום אחרי ניכוי"])
    c_pay  = col_idx(["מס' תשלום","מספר תשלום","אסמכתא 1","אסמכתא"])
    c_time = col_idx(["זמן","שעה"])  # אופציונלי

    if not (c_date and c_amt and c_pay):
        raise RuntimeError("בקובץ העזר חייבים הופיע: 'תאריך פריקה', 'אחרי ניכוי', 'מס' תשלום'.")

    # בונים key = תאריך (יום) + זמן (או ריק)
    rows = ws.iter_rows(min_row=head_row+1, values_only=True)
    bucket = defaultdict(lambda: {"amts": [], "pays": []})
    for r in rows:
        dt_raw = r[c_date-1]
        amt    = to_number(r[c_amt-1])
        pay    = "" if r[c_pay-1] is None else str(r[c_pay-1]).strip()
        tm     = ""
        if c_time:
            tm = r[c_time-1]
            if isinstance(tm, pd.Timestamp):
                tm = tm.strftime("%H:%M:%S")
            elif isinstance(tm, datetime):
                tm = tm.strftime("%H:%M:%S")
            else:
                tm = str(tm or "").strip()

        # אם התאריך כולל זמן באותה עמודה – נוציא
        dt = parse_date(dt_raw)
        tm2 = tm
        if tm2 == "" and isinstance(dt_raw, (pd.Timestamp, datetime)):
            tm2 = pd.to_datetime(dt_raw).strftime("%H:%M:%S")
        if isinstance(dt, pd.Timestamp):
            day_key = dt.normalize().strftime("%Y-%m-%d")
        else:
            day_key = ""
        key = f"{day_key} {tm2}".strip()

        if not math.isnan(amt) and pay:
            bucket[key]["amts"].append(abs(round(amt,2)))
            bucket[key]["pays"].append(pay)

    # סכום לכל key, ומיפוי סכום -> סט אסמכתאות
    amount_to_paynums = defaultdict(set)
    for key, d in bucket.items():
        total = round(sum(d["amts"]), 2)
        for p in set(d["pays"]):
            amount_to_paynums[total].add(p)

    return amount_to_paynums

# ---------------------- עיבוד קובץ המקור ----------------------
def process_workbook(source_path, aux_path, phrase, tolerance):
    amount_to_paynums = build_aux_amount_to_paynums(aux_path)

    wb = load_workbook(source_path)  # NOT read_only – נרצה לעדכן
    # נעדכן רק את עמודת "מס.התאמה" ונוסיף גיליון חדש
    for ws in wb.worksheets:
        ws.sheet_view.rightToLeft = True

        head_row, headers = find_header_row(ws)
        c_match = find_col(headers, MATCH_COL_CANDS) or 1  # ברירת מחדל – עמודה ראשונה
        c_code  = find_col(headers, BANK_CODE_CANDS)
        c_bamt  = find_col(headers, BANK_AMT_CANDS)
        c_books = find_col(headers, BOOKS_AMT_CANDS)
        c_ref   = find_col(headers, REF_CANDS)
        c_date  = find_col(headers, DATE_CANDS)
        c_det   = find_col(headers, DETAILS_CANDS)

        # נייצר עותקים נוחים לחישוב
        def cell(r, c):
            return ws.cell(row=r, column=c).value if c else None

        # חוצצים לשני הצדדים בהתאמה 1
        books_candidates = []
        used_books = set()

        # נאתר מראש מועמדים של ספרים (חיובי, OV/RC, עם תאריך)
        for r in range(head_row+1, ws.max_row+1):
            bamt  = to_number(cell(r, c_books)) if c_books else math.nan
            d     = parse_date(cell(r, c_date)) if c_date else None
            refv  = cell(r, c_ref) if c_ref else ""
            if not math.isnan(bamt) and bamt > 0 and d is not None and startswith_ov_rc(refv):
                books_candidates.append(r)

        # התאמה 1: קוד 175/120, סכום בדף במינוס; מחפשים ספרים תואמים
        for r in range(head_row+1, ws.max_row+1):
            code = cell(r, c_code)
            bamt = to_number(cell(r, c_bamt)) if c_bamt else math.nan
            d    = parse_date(cell(r, c_date)) if c_date else None
            if code is None or math.isnan(bamt) or d is None:
                continue
            try:
                code = int(code)
            except:
                continue
            if code in (175, 120) and bamt < 0:
                target_amt  = abs(round(bamt, 2))
                target_date = pd.to_datetime(d).normalize()
                # חפש התאמה בספרים
                chosen = None
                for rr in books_candidates:
                    if rr in used_books: continue
                    d2 = parse_date(cell(rr, c_date))
                    if d2 is None: continue
                    if pd.to_datetime(d2).normalize() != target_date:
                        continue
                    val_books = to_number(cell(rr, c_books))
                    if round(val_books,2) == target_amt:
                        chosen = rr
                        break
                if chosen:
                    if ws.cell(row=r, column=c_match).value in (None, ""):
                        ws.cell(row=r, column=c_match, value=1)
                    if ws.cell(row=chosen, column=c_match).value in (None, ""):
                        ws.cell(row=chosen, column=c_match, value=1)
                    used_books.add(chosen)

        # התאמה 2: קוד 469/515 – סימון 2 + איסוף לגיליון "הוראת קבע ספקים"
        standing_rows = getattr(ws, "_standing_rows", [])
        for r in range(head_row+1, ws.max_row+1):
            code = cell(r, c_code)
            det  = cell(r, c_det)
            bamt = to_number(cell(r, c_bamt)) if c_bamt else math.nan
            try:
                code = int(code)
            except:
                continue
            if code in (469, 515):
                if ws.cell(row=r, column=c_match).value in (None, ""):
                    ws.cell(row=r, column=c_match, value=2)
                standing_rows.append({"פרטים": det, "סכום": bamt})
        ws._standing_rows = standing_rows  # נשמור זמנית על האובייקט

        # התאמה 3: העברות ספקים – קוד 485 + ביטוי בפרטים + התאמה לסכום מקובץ עזר
        phrase_norm = normalize_for_contains(phrase)
        # מיפוי אסמכתא -> שורות
        ref_to_rows = defaultdict(list)
        if c_ref:
            for r in range(head_row+1, ws.max_row+1):
                rv = cell(r, c_ref)
                if rv is not None and str(rv).strip() != "":
                    ref_to_rows[str(rv).strip()].append(r)

        for r in range(head_row+1, ws.max_row+1):
            code = cell(r, c_code)
            det  = cell(r, c_det)
            bamt = to_number(cell(r, c_bamt)) if c_bamt else math.nan
            try:
                code = int(code)
            except:
                continue
            if code != 485 or math.isnan(bamt) or bamt <= 0:
                continue
            # בדיקת ביטוי בפרטים
            if phrase_norm and phrase_norm not in normalize_for_contains(det or ""):
                continue

            # התאמה לפי סכום, עם סבילות
            amount_match = None
            target = round(bamt, 2)
            if target in amount_to_paynums:
                amount_match = target
            else:
                for a in amount_to_paynums.keys():
                    if abs(a - target) <= tolerance:
                        amount_match = a
                        break
            if amount_match is None:
                continue

            # סמן את שורת הבנק
            if ws.cell(row=r, column=c_match).value in (None, ""):
                ws.cell(row=r, column=c_match, value=3)

            # סמן שורות ספרים שהאסמכתא שלהם באותה קבוצה
            for pay in amount_to_paynums[amount_match]:
                for rr in ref_to_rows.get(str(pay), []):
                    if ws.cell(row=rr, column=c_match).value in (None, ""):
                        ws.cell(row=rr, column=c_match, value=3)

    # יצירת גיליון "הוראת קבע ספקים"
    st_all = []
    for ws in wb.worksheets:
        if hasattr(ws, "_standing_rows"):
            st_all.extend(ws._standing_rows)

    if "הוראת קבע ספקים" in wb.sheetnames:
        del wb["הוראת קבע ספקים"]
    ws_new = wb.create_sheet("הוראת קבע ספקים")
    ws_new.sheet_view.rightToLeft = True
    headers = ["פרטים","סכום","מס' ספק","סכום חובה","סכום זכות"]
    for i, h in enumerate(headers, start=1):
        ws_new.cell(row=1, column=i, value=h).font = Font(bold=True)

    def map_supplier_by_name(name):
        s = normalize_text(name)
        # התאמה מדויקת
        if s in NAME_TO_SUPPLIER: return NAME_TO_SUPPLIER[s]
        # “מכיל” – לפי המפתח הארוך ביותר
        for k in sorted(NAME_TO_SUPPLIER.keys(), key=len, reverse=True):
            if k and k in s:
                return NAME_TO_SUPPLIER[k]
        return ""

    r = 2
    for row in st_all:
        det = row.get("פרטים","")
        amt = row.get("סכום", math.nan)
        supplier = map_supplier_by_name(det)
        if supplier in ("", None):
            if not math.isnan(amt):
                supplier = AMOUNT_TO_SUPPLIER.get(round(abs(float(amt)),2), "")
        debit  = amt if not math.isnan(amt) and amt > 0 else 0
        credit = abs(amt) if not math.isnan(amt) and amt < 0 else 0
        ws_new.cell(row=r, column=1, value=det)
        ws_new.cell(row=r, column=2, value=amt)
        ws_new.cell(row=r, column=3, value=supplier)
        ws_new.cell(row=r, column=4, value=debit)
        ws_new.cell(row=r, column=5, value=credit)
        # צביעה בכתום למי שאין מס' ספק
        if supplier in ("", None):
            for c in range(1,6):
                ws_new.cell(row=r, column=c).fill = ORANGE
        r += 1

    # שורת 20001 – זכות = סכום חובה של שורות שיש להן מס' ספק
    total_debit_with_supplier = 0.0
    for rr in range(2, ws_new.max_row+1):
        supplier = ws_new.cell(row=rr, column=3).value
        debit    = ws_new.cell(row=rr, column=4).value or 0
        if supplier not in ("", None):
            try:
                total_debit_with_supplier += float(debit)
            except:
                pass
    last = ws_new.max_row + 1
    ws_new.cell(row=last, column=1, value="סה\"כ זכות – עם מס' ספק").font = Font(bold=True)
    ws_new.cell(row=last, column=2, value="").font = Font(bold=True)
    ws_new.cell(row=last, column=3, value=20001).font = Font(bold=True)
    ws_new.cell(row=last, column=4, value=0).font = Font(bold=True)
    ws_new.cell(row=last, column=5, value=round(total_debit_with_supplier,2)).font = Font(bold=True)

    # שמירה
    out_name = "התאמות_והוראת_קבע_והעברות.xlsx"
    wb.save(out_name)
    print(f"✓ נשמר: {out_name}")

# ---------------------- main ----------------------
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="התאמות 1+2+3 לקובץ אקסל + גיליון הוראת קבע ספקים")
    parser.add_argument("--source", required=True, help="קובץ מקור (xlsx)")
    parser.add_argument("--aux",    required=True, help="קובץ עזר (xlsx)")
    parser.add_argument("--phrase", default="העב' במקבץ-נט", help="ביטוי בפרטים עבור התאמה 3")
    parser.add_argument("--tolerance", type=float, default=0.05, help="סבילות התאמת סכום בהתאמה 3")
    args = parser.parse_args()

    process_workbook(args.source, args.aux, args.phrase, args.tolerance)
