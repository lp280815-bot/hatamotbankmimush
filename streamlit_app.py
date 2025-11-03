# -*- coding: utf-8 -*-
from __future__ import annotations

import io, os, re, json
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# ================= UI =================
st.set_page_config(page_title="התאמות לקוחות – OV/RC + הוראות קבע + העברות", page_icon="✅", layout="centered")
st.markdown("""
<style>
  html, body, [class*="css"] { direction: rtl; text-align: right; }
  .block-container { padding-top: 1.1rem; }
</style>
""", unsafe_allow_html=True)
st.title("התאמות לקוחות – OV/RC + הוראות קבע + העברות")

# --------- ברירות מחדל לכללי VLOOKUP (ניתנים לעריכה ושמירה) ----------
DEFAULT_NAME_MAP = {
    "בזק בינלאומי ב": 30006,
    "פרי ירוחם חב'": 34714,
    "סלקום ישראל בע": 30055,
    "בזק-הוראות קבע": 34746,
    "דרך ארץ הייוי": 34602,
    "גלובס פבלישר ע": 30067,
    "פלאפון תקשורת": 30030,
    "מרכז הכוכביות": 30002,
    "ע.אשדוד-מסים": 30056,
    "א.ש.א(בס\"ד)אחז": 30050,
    "או.פי.ג'י(מ.כ)": 30047,
    "רשות האכיפה וה": "67-1",
    "קול ביז מילניו": 30053,
    "פריוריטי סופטו": 30097,
    "אינטרנט רימון": 34636,
    "עו\"דכנית בע\"מ": 30018,
    "עיריית רמת גן": 30065,
    "פז חברת נפט בע": 34811,
    "ישראכרט": 28002,
    "חברת החשמל ליש": 30015,
    "הפניקס ביטוח": 34686,
    "מימון ישיר מקב": 34002,
    "שלמה טפר": 30247,
    "נמרוד תבור עורך-דין": 30038,
    "עיריית בית שמש": 34805,
    "פז קמעונאות וא": 34811,
    "הו\"ק הלו' רבית": 8004,
    "הו\"ק הלואה קרן": 23001,
    # הרחבות:
    "עיריית אשדוד": 30056,
    "ישראכרט מור": 34002,
}
DEFAULT_AMOUNT_MAP = {
    8520.0: 30247,    # שלמה טפר
    10307.3: 30038,   # נמרוד תבור
}

# --------- מאגר שמות עמודות אפשריים (נתאים אוטומטית, אך נוכל לבחור ידנית) ----------
MATCH_COL_CANDS = ["מס.התאמה","מס. התאמה","מס התאמה","מספר התאמה","התאמה"]
BANK_CODE_CANDS = ["קוד פעולת בנק","קוד פעולת","קוד פעולה"]
BANK_AMT_CANDS  = ["סכום בדף","סכום דף","סכום בבנק","סכום תנועת בנק"]
BOOKS_AMT_CANDS = ["סכום בספרים","סכום בספר","סכום ספרים"]
REF_CANDS       = ["אסמכתא 1","אסמכתא1","אסמכתא","אסמכתה"]
DATE_CANDS      = ["תאריך מאזן","תאריך ערך","תאריך"]
DETAILS_CANDS   = ["פרטים","תיאור","שם ספק","שם הפעולה"]

# קובץ כללים
RULES_FILE = "rules_store.json"

# ================= עזר =================
def normalize_text(s):
    if s is None:
        return ""
    t = str(s)
    t = t.replace("’","").replace("`","").replace('"','').replace("'","")
    t = t.replace("–"," ").replace("־"," ").replace("-"," ")
    t = re.sub(r"\s+"," ",t).strip()
    return t

def to_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series.astype(str).str.replace(",","").str.replace("₪","").str.strip(),
                         errors="coerce")

def normalize_date(series: pd.Series) -> pd.Series:
    def f(x):
        if pd.isna(x): return pd.NaT
        if isinstance(x,(pd.Timestamp, datetime)): return pd.Timestamp(x.date())
        return pd.to_datetime(x, dayfirst=True, errors="coerce").normalize()
    return series.apply(f)

def ws_to_df(ws) -> pd.DataFrame:
    rows = list(ws.iter_rows(values_only=True))
    if not rows: return pd.DataFrame()
    header, start = None, 0
    for i, r in enumerate(rows):
        if any(x is not None for x in r):
            header = [str(x).strip() if x is not None else "" for x in r]
            start = i+1; break
    if header is None: return pd.DataFrame()
    data = [tuple(list(r)[:len(header)]) for r in rows[start:]]
    return pd.DataFrame(data, columns=header)

def first_match(candidates, cols):
    for n in candidates:
        if n in cols: return n
    for n in candidates:
        for c in cols:
            if isinstance(c,str) and n in c: return c
    return cols[0] if cols else None

def ref_starts_with_ov_rc(val) -> bool:
    s = (str(val) if val is not None else "").upper().strip()
    return s.startswith("OV") or s.startswith("RC")

# ================= שמירת/טעינת כללי VLOOKUP =================
def load_rules():
    if os.path.exists(RULES_FILE):
        try:
            data = json.load(open(RULES_FILE,"r",encoding="utf-8"))
            name_map = { normalize_text(k): v for k,v in data.get("name_map",{}).items() }
            amount_map = { float(k): v for k,v in data.get("amount_map",{}).items() }
            return name_map, amount_map
        except Exception:
            pass
    return { normalize_text(k): v for k,v in DEFAULT_NAME_MAP.items() }, dict(DEFAULT_AMOUNT_MAP)

def save_rules(name_map, amount_map):
    with open(RULES_FILE,"w",encoding="utf-8") as f:
        json.dump({"name_map": name_map, "amount_map": amount_map}, f, ensure_ascii=False, indent=2)

if "name_map" not in st.session_state:
    nm, am = load_rules()
    st.session_state.name_map = nm
    st.session_state.amount_map = am

# ================= לשונית: עדכון כללי VLOOKUP =================
with st.expander("⚙️ עדכון – כללי VLOOKUP (שומר לקובץ rules_store.json)", expanded=False):
    mode = st.radio("סוג עדכון", ["לפי פרטים (שם)","לפי סכום"], horizontal=True)
    if mode == "לפי פרטים (שם)":
        name = st.text_input("פרטים (כמו בבנק)")
        sup  = st.text_input("מס' ספק")
        c1, c2, c3 = st.columns([1,1,1])
        if c1.button("➕ הוסף/עדכן"):
            k = normalize_text(name)
            if k and sup:
                st.session_state.name_map[k] = sup
                save_rules(st.session_state.name_map, st.session_state.amount_map)
                st.success("נשמר.")
        if c2.button("🗑️ מחיקה"):
            k = normalize_text(name)
            if k in st.session_state.name_map:
                del st.session_state.name_map[k]
                save_rules(st.session_state.name_map, st.session_state.amount_map)
                st.success("נמחק.")
        if c3.button("💾 שמור ידנית"):
            save_rules(st.session_state.name_map, st.session_state.amount_map)
            st.info("נשמר.")
        st.dataframe(pd.DataFrame({"by_name": list(st.session_state.name_map.keys()),
                                   "מס' ספק": list(st.session_state.name_map.values())}),
                     use_container_width=True, height=240)
    else:
        amt = st.number_input("סכום (יישמר בערך מוחלט)", value=0.0, step=0.01, format="%.2f")
        sup = st.text_input("מס' ספק")
        c1, c2, c3 = st.columns([1,1,1])
        if c1.button("➕ הוסף/עדכן", key="add_amt"):
            key = round(abs(float(amt)),2)
            if key and sup:
                st.session_state.amount_map[key] = sup
                save_rules(st.session_state.name_map, st.session_state.amount_map)
                st.success("נשמר.")
        if c2.button("🗑️ מחיקה", key="del_amt"):
            key = round(abs(float(amt)),2)
            if key in st.session_state.amount_map:
                del st.session_state.amount_map[key]
                save_rules(st.session_state.name_map, st.session_state.amount_map)
                st.success("נמחק.")
        if c3.button("💾 שמור ידנית", key="save_amt"):
            save_rules(st.session_state.name_map, st.session_state.amount_map)
            st.info("נשמר.")
        st.dataframe(pd.DataFrame({"סכום": list(st.session_state.amount_map.keys()),
                                   "מס' ספק": list(st.session_state.amount_map.values())})
                     .sort_values("סכום"), use_container_width=True, height=240)

st.divider()

# ================== קלט הקבצים ==================
st.subheader("1) קובץ מקור (Excel) + (אופציונלי) קובץ עזר להעברות")
main_file = st.file_uploader("קובץ מקור (xlsx)", type=["xlsx"])
aux_file  = st.file_uploader("קובץ עזר: תאריך פריקה, זמן*, אחרי ניכוי, מס' תשלום (xlsx)", type=["xlsx"])

# פרמטרים להעברות
st.subheader("2) פרמטרים להתאמת העברות (מס' התאמה 3)")
p1, p2, p3, p4 = st.columns([1,1.2,1,1.6])
transfer_code   = p1.number_input("קוד פעולה", value=485, step=1)
details_phrase  = p2.text_input("ביטוי בפרטים", value="העב' במקבץ-נט")
amount_tol      = p3.number_input("סבילות סכום (₪)", value=0.05, step=0.01, format="%.2f")
ignore_time     = p4.checkbox("להתעלם משדה זמן בקובץ העזר (קיבוץ לפי תאריך בלבד)", value=False)

st.divider()
run_btn = st.button("▶️ הרצה")

# ================== עיבוד ==================
def build_amount_to_paynums_explicit(aux_df,
                                     col_date: str, col_amount: str, col_paynum: str,
                                     col_time: str | None, ignore_time: bool):
    amt = to_number(aux_df[col_amount]).fillna(0).abs().round(2)
    dt  = normalize_date(aux_df[col_date])
    key = dt.astype(str)
    if col_time and not ignore_time and col_time in aux_df.columns:
        try:
            tm = pd.to_datetime(aux_df[col_time], errors="coerce").dt.strftime("%H:%M:%S")
        except Exception:
            tm = aux_df[col_time].astype(str)
        key = key + " " + tm.fillna("")
    work = pd.DataFrame({"key": key, "amt": amt, "pay": aux_df[col_paynum].astype(str).str.strip()})
    sums = work.groupby("key")["amt"].sum().round(2)
    amount_to_paynums = {}
    for k, total in sums.items():
        pays = set(work.loc[work["key"]==k, "pay"].dropna().astype(str))
        amount_to_paynums.setdefault(float(total), set()).update(pays)
    return amount_to_paynums, sums.reset_index().rename(columns={"key": "קבוצה (תאריך+זמן)", "amt": "סכום אחרי ניכוי"})

def process(main_bytes: bytes, aux_bytes: bytes | None):
    # טוענים את קובץ המקור
    wb_in = load_workbook(io.BytesIO(main_bytes), data_only=True, read_only=True)
    df = ws_to_df(wb_in.worksheets[0])  # גיליון ראשון (כמו בקבצים שלך)

    # זיהוי עמודות (נוכל לשנות ידנית אם נרצה)
    cols = list(df.columns)
    col_match   = first_match(MATCH_COL_CANDS, cols)
    col_code    = first_match(BANK_CODE_CANDS, cols)
    col_bank    = first_match(BANK_AMT_CANDS, cols)
    col_books   = first_match(BOOKS_AMT_CANDS, cols)
    col_ref     = first_match(REF_CANDS, cols)
    col_date    = first_match(DATE_CANDS, cols)
    col_details = first_match(DETAILS_CANDS, cols)

    st.write("**זיהוי עמודות (ניתן לשנות):**")
    c1, c2, c3 = st.columns(3)
    col_match   = c1.selectbox("מס. התאמה", cols, index=cols.index(col_match))
    col_code    = c1.selectbox("קוד פעולת בנק", cols, index=cols.index(col_code))
    col_bank    = c1.selectbox("סכום בדף (בבנק)", cols, index=cols.index(col_bank))

    col_books   = c2.selectbox("סכום בספרים", cols, index=cols.index(col_books))
    col_ref     = c2.selectbox("אסמכתא 1", cols, index=cols.index(col_ref))
    col_date    = c2.selectbox("תאריך", cols, index=cols.index(col_date))

    col_details = c3.selectbox("פרטים", cols, index=cols.index(col_details))

    # ממירים לסדרות עבודה
    s_match = df[col_match].copy()
    s_code  = to_number(df[col_code])
    s_bank  = to_number(df[col_bank])
    s_books = to_number(df[col_books])
    s_ref   = df[col_ref].astype(str)
    s_date  = normalize_date(df[col_date])
    s_det   = df[col_details].astype(str)

    # ------------ (1) התאמות OV/RC = 1 ------------
    pairs = 0
    books_candidates = df.index[(s_books > 0) & s_ref.apply(ref_starts_with_ov_rc) & s_date.notna()]
    used_books = set()
    for i in df.index[(s_code.isin([175,120])) & (s_bank < 0) & s_date.notna()]:
        target_amt  = round(abs(float(s_bank.iat[i])),2)
        target_date = s_date.iat[i]
        cands = [j for j in books_candidates if j not in used_books
                 and s_date.iat[j] == target_date
                 and round(float(s_books.iat[j]),2) == target_amt]
        if cands:
            chosen = min(cands, key=lambda j: abs(j-i))
            if s_match.iat[i] not in (1,2,3): s_match.iat[i] = 1
            if s_match.iat[chosen] not in (1,2,3): s_match.iat[chosen] = 1
            used_books.add(chosen)
            pairs += 1

    # ------------ (2) הוראות קבע = 2 + גיליון סיכום ------------
    standing_rows = []
    for i in df.index[s_code.isin([515,469])]:
        if s_match.iat[i] in (1,3):  # לא לדרוס 1/3
            continue
        s_match.iat[i] = 2
        standing_rows.append({"פרטים": s_det.iat[i], "סכום": s_bank.iat[i]})

    st_df = pd.DataFrame(standing_rows)
    if not st_df.empty:
        def map_supplier(name):
            s = normalize_text(name)
            if s in st.session_state.name_map:
                return st.session_state.name_map[s]
            for key in sorted(st.session_state.name_map.keys(), key=len, reverse=True):
                if key and key in s:
                    return st.session_state.name_map[key]
            return ""
        st_df["מס' ספק"] = st_df["פרטים"].apply(map_supplier)
        def by_amount(row):
            if not row["מס' ספק"]:
                val = round(abs(float(row["סכום"] or 0)),2)
                return st.session_state.amount_map.get(val,"")
            return row["מס' ספק"]
        st_df["מס' ספק"] = st_df.apply(by_amount, axis=1)
        st_df["סכום חובה"] = st_df["סכום"].apply(lambda x: x if pd.notna(x) and x>0 else 0)
        st_df["סכום זכות"] = st_df["סכום"].apply(lambda x: abs(x) if pd.notna(x) and x<0 else 0)
        st_df = st_df[["פרטים","סכום","מס' ספק","סכום חובה","סכום זכות"]]
    else:
        st_df = pd.DataFrame(columns=["פרטים","סכום","מס' ספק","סכום חובה","סכום זכות"])

    # ------------ (3) העברות = 3 (עם קובץ עזר) ------------
    log3 = []
    if aux_bytes is not None:
        wb_aux = load_workbook(io.BytesIO(aux_bytes), data_only=True, read_only=True)
        aux = ws_to_df(wb_aux.worksheets[0])
        st.write("**בחירת עמודות בקובץ העזר:**")
        acols = list(aux.columns)
        aux_date  = st.selectbox("תאריך פריקה", acols, index=acols.index(first_match(["תאריך פריקה","תאריך"], acols)))
        aux_time  = st.selectbox("זמן (לא חובה)", ["(ללא)"]+acols, index=0)
        aux_time  = None if aux_time=="(ללא)" else aux_time
        aux_amt   = st.selectbox("אחרי ניכוי", acols, index=acols.index(first_match(["אחרי ניכוי","אחרי ניכוי מס","סכום אחרי ניכוי"], acols)))
        aux_pay   = st.selectbox("מס' תשלום", acols, index=acols.index(first_match(["מס' תשלום","מספר תשלום","אסמכתא תשלום"], acols)))

        amount_to_paynums, aux_groups = build_amount_to_paynums_explicit(aux, aux_date, aux_amt, aux_pay, aux_time, ignore_time)

        # מועמדות בנק: 485 + ביטוי + סכום חיובי
        bank_idx = df.index[(s_code == float(transfer_code)) & (s_bank > 0) & (s_det.str.contains(details_phrase, na=False))]
        for i in bank_idx:
            amt = round(float(s_bank.iat[i]),2)
            # התאמה עם סבילות
            paynums = set()
            for key_amt,pays in amount_to_paynums.items():
                if abs(key_amt - amt) <= float(amount_tol):
                    paynums |= pays
            if not paynums:
                log3.append({"שורה": int(i+1), "סכום בנק": amt, "מס' תשלום": "", "סטאטוס": "לא נמצא בקובץ עזר"})
                continue
            if s_match.iat[i] not in (1,2):  # לא לדרוס 1/2
                s_match.iat[i] = 3
            mask = s_ref.isin(paynums)
            for j in df.index[mask]:
                if s_match.iat[j] not in (1,2):
                    s_match.iat[j] = 3
            log3.append({"שורה": int(i+1), "סכום בנק": amt, "מס' תשלום": ", ".join(sorted(paynums)), "סטאטוס": "סומן 3 (כולל התאמת ספרים לפי אסמכתא)"})
    else:
        st.info("לא עלה קובץ עזר – התאמה 3 תדלג.")

    # ------------ כתיבה לקובץ ------------
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        df_out = df.copy()
        df_out[col_match] = s_match
        df_out.to_excel(w, index=False, sheet_name="DataSheet")
        st_df.to_excel(w, index=False, sheet_name="הוראת קבע ספקים")
        if log3:
            pd.DataFrame(log3).to_excel(w, index=False, sheet_name="לוג_התאמות_3")
        # RTL יתווסף אחרי שמירה

    # עיצוב נוסף + שורת 20001
    wb_out = load_workbook(io.BytesIO(out.getvalue()))
    for s in wb_out.worksheets:
        s.sheet_view.rightToLeft = True

    if "הוראת קבע ספקים" in wb_out.sheetnames:
        ws = wb_out["הוראת קבע ספקים"]
        headers = {cell.value: idx for idx, cell in enumerate(ws[1], start=1)}
        col_details = headers.get("פרטים")
        col_supplier = headers.get("מס' ספק")
        col_debit = headers.get("סכום חובה")
        col_credit = headers.get("סכום זכות")
        orange = PatternFill(start_color="FFDDBB", end_color="FFDDBB", fill_type="solid")

        # צבע לשורות בלי מס' ספק
        if col_supplier:
            for r in range(2, ws.max_row+1):
                v = ws.cell(row=r, column=col_supplier).value
                if v in ("", None):
                    for c in range(1, ws.max_column+1):
                        ws.cell(row=r, column=c).fill = orange

        # מחיקה קודמת של 20001
        dels = []
        for r in range(2, ws.max_row+1):
            if ws.cell(row=r, column=col_supplier).value == 20001:
                dels.append(r)
        for k, r in enumerate(dels):
            ws.delete_rows(r-k, 1)

        # סה"כ חובה לשורות שיש בהן מס' ספק -> כתיבה בשורת 20001 בזכות
        total_debit = 0.0
        for r in range(2, ws.max_row+1):
            sv = ws.cell(row=r, column=col_supplier).value
            try:
                if sv not in (None, ""):
                    total_debit += float(ws.cell(row=r, column=col_debit).value or 0)
            except Exception:
                pass

        last = ws.max_row + 1
        if col_details: ws.cell(row=last, column=col_details, value='סה"כ זכות – עם מס׳ ספק')
        if col_supplier: ws.cell(row=last, column=col_supplier, value=20001)
        if col_debit: ws.cell(row=last, column=col_debit, value=0)
        if col_credit: ws.cell(row=last, column=col_credit, value=round(total_debit,2))
        for c in range(1, ws.max_column+1):
            ws.cell(row=last, column=c).font = Font(bold=True)

    final = io.BytesIO()
    wb_out.save(final)
    return final.getvalue(), pairs, len(standing_rows), len(log3)

# ================== RUN ==================
if run_btn:
    if main_file is None:
        st.error("נא להעלות קובץ מקור.")
    else:
        with st.spinner("מריץ התאמות..."):
            aux_bytes = aux_file.read() if aux_file is not None else None
            out_bytes, pairs, st_count, tr3 = process(main_file.read(), aux_bytes)
        st.success(f"הסתיים! OV/RC=1: {pairs} זוגות • הוראות קבע=2: {st_count} שורות • העברות=3: {tr3} אירועים")
        st.download_button("⬇️ הורדת הקובץ המעודכן",
                           data=out_bytes,
                           file_name="התאמות_מעודכן.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
