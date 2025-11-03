# streamlit_app.py
# -*- coding: utf-8 -*-
import io, re, os, json
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# ---------------- Page / RTL ----------------
st.set_page_config(page_title="התאמות – OV/RC + הוראות קבע + העברות", page_icon="✅", layout="centered")
st.markdown("""
<style>
  html, body, [class*="css"] { direction: rtl; text-align: right; }
  .block-container { padding-top: 1.2rem; max-width: 1100px; }
</style>
""", unsafe_allow_html=True)

st.title("התאמות לקוחות – OV/RC + הוראות קבע (VLOOKUP קבוע + שמירה) + התאמה 3 – העברות ספקים")

# -------------------- Defaults (VLOOKUP) --------------------
RAW_NAME_MAP = {
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
    # תוספות:
    "עיריית אשדוד": 30056, "ישראכרט מור": 34002,
}
BASE_AMOUNT_MAP = { 8520.0: 30247, 10307.3: 30038 }

# -------------------- Helpers --------------------
MATCH_COL_CANDS = ["מס.התאמה","מס. התאמה","מס התאמה","מספר התאמה","התאמה"]
BANK_CODE_CANDS = ["קוד פעולת בנק","קוד פעולה","קוד פעולת"]
BANK_AMT_CANDS  = ["סכום בדף","סכום דף","סכום בבנק","סכום תנועת בנק"]
BOOKS_AMT_CANDS = ["סכום בספרים","סכום בספר","סכום ספרים"]
REF_CANDS       = ["אסמכתא 1","אסמכתא1","אסמכתא","אסמכתה"]
DATE_CANDS      = ["תאריך מאזן","תאריך ערך","תאריך"]
DETAILS_CANDS   = ["פרטים","תיאור","שם ספק"]

RULES_FILE = "rules_store.json"

def normalize_text(s: str) -> str:
    if s is None: return ""
    t = str(s)
    t = t.replace("’","'").replace("`","'")  # גרשים "חכמים"
    t = t.replace("״","").replace("„","").replace('"', "")
    t = t.replace("–","-").replace("־","-")
    t = t.replace("\u200f","").replace("\u200e","")  # RTL marks
    t = re.sub(r"\s+", " ", t)
    return t.strip()

def normalize_for_contains(s: str) -> str:
    """נורמליזציה קשוחה להשוואת 'מכיל' – בלי גרשים/מקפים/רווחים, אותיות קטנות."""
    s = normalize_text(s)
    s = s.lower()
    s = s.replace("-", "").replace(" ", "").replace("'", "")
    return s

def load_rules_from_disk():
    if os.path.exists(RULES_FILE):
        try:
            with open(RULES_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            name_map = { normalize_text(k): v for k, v in data.get("name_map", {}).items() }
            amount_map = { float(k): v for k, v in data.get("amount_map", {}).items() }
            return name_map, amount_map
        except Exception:
            pass
    return { normalize_text(k): v for k, v in RAW_NAME_MAP.items() }, dict(BASE_AMOUNT_MAP)

def save_rules_to_disk(name_map, amount_map):
    try:
        with open(RULES_FILE, "w", encoding="utf-8") as f:
            json.dump({"name_map": name_map, "amount_map": amount_map}, f, ensure_ascii=False, indent=2)
        return True
    except Exception:
        return False

if "name_map" not in st.session_state or "amount_map" not in st.session_state:
    nm, am = load_rules_from_disk()
    st.session_state.name_map = nm
    st.session_state.amount_map = am

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

def normalize_date(series):
    def f(x):
        if pd.isna(x): return pd.NaT
        if isinstance(x,(pd.Timestamp, datetime)): return pd.Timestamp(x.date())
        return pd.to_datetime(x, dayfirst=True, errors="coerce").normalize()
    return series.apply(f)

def to_number(series):
    s = series.astype(str).str.replace(",","").str.replace("₪","").str.strip()
    return pd.to_numeric(s, errors="coerce")

def ref_starts_with_ov_rc(val):
    t = (str(val) if val is not None else "").strip().upper()
    return t.startswith("OV") or t.startswith("RC")

# ---------------- התאמה 3 – קיבוץ תאריך+זמן ----------------
def build_amount_to_paynums_explicit(aux_df, col_date: str, col_amount: str, col_paynum: str, col_time: str | None):
    amt = to_number(aux_df[col_amount]).fillna(0).abs().round(2)
    if not col_time or col_time == col_date:
        dt_all = pd.to_datetime(aux_df[col_date], dayfirst=True, errors="coerce")
        dt_only = dt_all.dt.normalize()
        tm_only = dt_all.dt.strftime("%H:%M:%S").fillna("")
    else:
        dt_only = pd.to_datetime(aux_df[col_date], dayfirst=True, errors="coerce").dt.normalize()
        tm_only = pd.to_datetime(aux_df[col_time], errors="coerce").dt.strftime("%H:%M:%S")
        tm_only = tm_only.fillna(aux_df[col_time].astype(str).str.strip())
    key = dt_only.astype(str) + " " + tm_only.fillna("")
    work = pd.DataFrame({"key": key,
                         "amt": amt,
                         "pay": aux_df[col_paynum].astype(str).str.strip()})
    sums = work.groupby("key")["amt"].sum().round(2)
    amount_to_paynums = {}
    for k, total in sums.items():
        pays = set(work.loc[work["key"] == k, "pay"].dropna().astype(str))
        amount_to_paynums.setdefault(float(total), set()).update(pays)
    return amount_to_paynums, sums.reset_index().rename(
        columns={"key": "קבוצה (תאריך+זמן)", "amt": "סכום אחרי ניכוי"}
    )

# ---------------- UI התאמה 3 + דיאגנוסטיקה ----------------
with st.expander("🔗 התאמה 3 – העברות ספקים (עם נורמליזציה חזקה לביטוי)", expanded=True):
    c1, c2, c3 = st.columns([1,1,1])
    t3_bank_code = c1.number_input("קוד פעולה בקובץ מקור", value=485, step=1)
    t3_phrase    = c2.text_input("ביטוי בפרטים (למשל: העב' במקבץ-נט)", value="העב' במקבץ-נט")
    t3_tol       = c3.number_input("סבילות סכום (₪)", value=0.05, step=0.01, format="%.2f")

    aux_file = st.file_uploader("קובץ עזר (xlsx) – עמודות חובה: תאריך פריקה (יכול לכלול זמן), אחרי ניכוי, מס' תשלום (+אופציונלי זמן)", type=["xlsx"])
    if aux_file is not None:
        try:
            wb_aux = load_workbook(io.BytesIO(aux_file.read()), data_only=True, read_only=True)
            df_aux = ws_to_df(wb_aux.worksheets[0])

            def fm(names):
                for n in names:
                    if n in df_aux.columns: return n
                for n in names:
                    for c in df_aux.columns:
                        if isinstance(c, str) and n in c: return c
                return None

            col_date = fm(["תאריך פריקה","תאריך"])
            col_amt  = fm(["אחרי ניכוי","אחרי ניכוי מס","סכום אחרי ניכוי"])
            col_pay  = fm(["מס' תשלום","מספר תשלום","אסמכתא תשלום","אסמכתא 1","אסמכתא"])
            col_time = fm(["זמן","שעה"])  # יכול להיות None

            if not all([col_date, col_amt, col_pay]):
                st.error("בקובץ העזר חייבים להופיע: 'תאריך פריקה' (יכול לכלול זמן), 'אחרי ניכוי', 'מס' תשלום'.")
                st.session_state.pop("t3_map", None)
            else:
                t3_map, t3_groups = build_amount_to_paynums_explicit(df_aux, col_date, col_amt, col_pay, col_time)
                st.session_state.t3_map = {"amount_to_paynums": t3_map,
                                           "bank_code": int(t3_bank_code),
                                           "phrase": t3_phrase,
                                           "tol": float(t3_tol)}
                st.success(f"נטענו {len(t3_groups)} קבוצות (תאריך+זמן).")
                st.dataframe(t3_groups.head(200), use_container_width=True)
        except Exception as e:
            st.error(f"שגיאה בקריאת קובץ העזר: {e}")
            st.session_state.pop("t3_map", None)

st.divider()

# ---------------- עדכון כללי VLOOKUP ----------------
def rules_excel_bytes():
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        pd.DataFrame({"by_name": list(st.session_state.name_map.keys()),
                      "מס' ספק": list(st.session_state.name_map.values())}).to_excel(w, index=False, sheet_name="by_name")
        pd.DataFrame({"סכום": list(st.session_state.amount_map.keys()),
                      "מס' ספק": list(st.session_state.amount_map.values())}).to_excel(w, index=False, sheet_name="by_amount")
    return out.getvalue()

with st.expander("⚙️ עדכון – VLOOKUP קבוע (שמירה מתמשכת)", expanded=False):
    st.write("עדכון לפי **פרטים** (שם) או לפי **סכום**. נשמר ל־`rules_store.json`.")
    mode = st.radio("סוג עדכון", ["לפי פרטים (שם)", "לפי סכום"], horizontal=True)

    if mode == "לפי פרטים (שם)":
        name_input = st.text_input("פרטים (כמו שמופיע בדף הבנק)")
        supplier_input = st.text_input("מס' ספק")
        c = st.columns([1,1,1,1])
        if c[0].button("➕ הוסף/עדכן"):
            k = normalize_text(name_input)
            if k and supplier_input:
                st.session_state.name_map[k] = supplier_input
                save_rules_to_disk(st.session_state.name_map, st.session_state.amount_map)
                st.success(f"הכלל נשמר: '{k}' → {supplier_input}")
        if c[1].button("🗑️ מחיקה"):
            k = normalize_text(name_input)
            if k in st.session_state.name_map:
                del st.session_state.name_map[k]
                save_rules_to_disk(st.session_state.name_map, st.session_state.amount_map)
                st.warning(f"הכלל נמחק: '{k}'")
        if c[2].button("💾 שמור ידנית"):
            save_rules_to_disk(st.session_state.name_map, st.session_state.amount_map)
            st.info("נשמר לקובץ rules_store.json")
        st.dataframe(pd.DataFrame({"by_name": list(st.session_state.name_map.keys()),
                                   "מס' ספק": list(st.session_state.name_map.values())}),
                     use_container_width=True, height=260)

    else:
        amount_input = st.number_input("סכום (חיובי/שלילי – יישמר בערך מוחלט)", step=0.01, format="%.2f")
        supplier_input2 = st.text_input("מס' ספק", key="amount_supplier")
        c = st.columns([1,1,1,1])
        if c[0].button("➕ הוסף/עדכן", key="add_amount"):
            key_amt = round(abs(float(amount_input)), 2)
            if key_amt and supplier_input2:
                st.session_state.amount_map[key_amt] = supplier_input2
                save_rules_to_disk(st.session_state.name_map, st.session_state.amount_map)
                st.success(f"הכלל נשמר: {key_amt} → {supplier_input2}")
        if c[1].button("🗑️ מחיקה", key="del_amount"):
            key_amt = round(abs(float(amount_input)), 2)
            if key_amt in st.session_state.amount_map:
                del st.session_state.amount_map[key_amt]
                save_rules_to_disk(st.session_state.name_map, st.session_state.amount_map)
                st.warning(f"הכלל נמחק: {key_amt}")
        if c[2].button("💾 שמור ידנית", key="save_amount"):
            save_rules_to_disk(st.session_state.name_map, st.session_state.amount_map)
            st.info("נשמר לקובץ rules_store.json")
        st.dataframe(pd.DataFrame({"סכום": list(st.session_state.amount_map.keys()),
                                   "מס' ספק": list(st.session_state.amount_map.values())})
                     .sort_values("סכום"), use_container_width=True, height=260)

st.divider()

# ---------------- עיבוד הקובץ (1+2+3) ----------------
def process_workbook(xlsx_bytes, t3_ctx=None):
    wb_in = load_workbook(io.BytesIO(xlsx_bytes), data_only=True, read_only=True)
    out_stream = io.BytesIO()
    summary_rows, standing_rows = [], []

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
            applied_transfers = False
            pairs = 0
            flagged = 0
            matched3_bank = 0
            matched3_books = 0

            match_values = df_save[col_match].copy() if col_match in df_save.columns else pd.Series([None]*len(df))
            _date      = normalize_date(pd.to_datetime(df[col_date], errors="coerce")) if col_date else pd.Series([pd.NaT]*len(df))
            _bank_amt  = to_number(df[col_bank_amt])  if col_bank_amt  else pd.Series([np.nan]*len(df))
            _books_amt = to_number(df[col_books_amt]) if col_books_amt else pd.Series([np.nan]*len(df))
            _bank_code = to_number(df[col_bank_code]) if col_bank_code else pd.Series([np.nan]*len(df))
            _ref       = df[col_ref].astype(str).fillna("") if col_ref else pd.Series([""]*len(df))
            _details   = df[col_details].astype(str).fillna("") if col_details else pd.Series([""]*len(df))

            # -------- התאמה 1: OV/RC --------
            if all([col_bank_code, col_bank_amt, col_books_amt, col_ref, col_date]):
                applied_ovrc = True
                books_candidates = [
                    j for j in range(len(df))
                    if pd.notna(_books_amt.iat[j]) and _books_amt.iat[j] > 0
                    and pd.notna(_date.iat[j]) and ref_starts_with_ov_rc(_ref.iat[j])
                ]
                used_books = set()
                for i in range(len(df)):
                    if pd.notna(_bank_code.iat[i]) and int(_bank_code.iat[i]) in (175, 120) \
                       and pd.notna(_bank_amt.iat[i]) and _bank_amt.iat[i] < 0 \
                       and pd.notna(_date.iat[i]):
                        target_amt = round(abs(float(_bank_amt.iat[i])), 2)
                        target_date = _date.iat[i]
                        cands = [
                            j for j in books_candidates
                            if j not in used_books
                            and _date.iat[j] == target_date
                            and round(float(_books_amt.iat[j]), 2) == target_amt
                        ]
                        chosen = None
                        if len(cands) == 1:
                            chosen = cands[0]
                        elif len(cands) > 1:
                            chosen = min(cands, key=lambda j: abs(j - i))
                        if chosen is not None:
                            if pd.isna(match_values.iat[i]): match_values.iat[i] = 1
                            if pd.isna(match_values.iat[chosen]): match_values.iat[chosen] = 1
                            used_books.add(chosen)
                            pairs += 1

            # -------- התאמה 2: הוראות קבע 469/515 --------
            if all([col_bank_code, col_details, col_bank_amt]):
                applied_standing = True
                for i in range(len(df)):
                    code = _bank_code.iat[i]
                    if pd.notna(code) and int(code) in (515, 469):
                        match_values.iat[i] = 2 if pd.isna(match_values.iat[i]) else match_values.iat[i]
                        flagged += 1
                        standing_rows.append({"פרטים": _details.iat[i], "סכום": _bank_amt.iat[i]})

            # -------- התאמה 3: העברות ספקים --------
            if t3_ctx and all([col_bank_code, col_details, col_bank_amt, col_ref]):
                applied_transfers = True
                code_needed = int(t3_ctx["bank_code"])
                phrase = str(t3_ctx["phrase"]).strip()
                tol = float(t3_ctx["tol"])
                amt2pay = t3_ctx["amount_to_paynums"]

                # נורמליזציה של הביטוי והפרטים
                phrase_norm = normalize_for_contains(phrase)
                details_norm_series = _details.map(normalize_for_contains)

                refs_to_rows = {}
                for j in range(len(df)):
                    r = _ref.iat[j]
                    if r:
                        refs_to_rows.setdefault(str(r).strip(), []).append(j)

                # דיאגנוסטיקה: כמה שורות עומדות בתנאי קוד פעולה?
                code_hits = (pd.notna(_bank_code)) & (_bank_code.astype("Int64") == code_needed)
                st.write(f"**אבחון (גיליון {ws.title}) – התאמה 3**: שורות עם קוד פעולה {code_needed}: {int(code_hits.sum())}")

                for i in range(len(df)):
                    if not code_hits.iat[i]:
                        continue
                    # בקובץ המקור: סכום בדף בפלוס
                    bam = _bank_amt.iat[i]
                    if pd.isna(bam) or float(bam) <= 0:
                        continue

                    # בדיקת ביטוי בפרטים – נורמליזציה חזקה
                    if phrase_norm and phrase_norm not in details_norm_series.iat[i]:
                        continue

                    target = float(bam)
                    matched_amount = None
                    if round(target,2) in amt2pay:
                        matched_amount = round(target,2)
                    else:
                        for a in amt2pay.keys():
                            if abs(a - target) <= tol:
                                matched_amount = a
                                break
                    if matched_amount is None:
                        continue

                    if pd.isna(match_values.iat[i]):
                        match_values.iat[i] = 3
                        matched3_bank += 1

                    for p in amt2pay[matched_amount]:
                        for j in refs_to_rows.get(str(p), []):
                            if pd.isna(match_values.iat[j]):
                                match_values.iat[j] = 3
                                matched3_books += 1

            # כתיבה
            df_out = df_save.copy()
            df_out[col_match] = match_values
            df_out.to_excel(writer, index=False, sheet_name=ws.title)

            summary_rows.append({
                "גיליון": ws.title,
                "OV/RC בוצע": "כן" if applied_ovrc else "לא",
                "זוגות שסומנו 1": pairs,
                "הוראת קבע בוצע": "כן" if applied_standing else "לא",
                "שורות שסומנו 2": flagged,
                "העברות (מס' 3) בוצע": "כן" if applied_transfers else "לא",
                "מס' 3 – בנק": matched3_bank,
                "מס' 3 – ספרים": matched3_books,
                "עמודת התאמה": col_match
            })

        # ---- גיליון "הוראת קבע ספקים" ----
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
                    if pd.notna(row["סכום"]):
                        val = round(abs(float(row["סכום"])), 2)
                        return st.session_state.amount_map.get(val, "")
                    return ""
                return row["מס' ספק"]

            st_df["מס' ספק"] = st_df.apply(by_amount, axis=1)
            st_df["סכום חובה"] = st_df["סכום"].apply(lambda x: x if pd.notna(x) and x > 0 else 0)
            st_df["סכום זכות"] = st_df["סכום"].apply(lambda x: abs(x) if pd.notna(x) and x < 0 else 0)
            st_df = st_df[["פרטים","סכום","מס' ספק","סכום חובה","סכום זכות"]]
        else:
            st_df = pd.DataFrame(columns=["פרטים","סכום","מס' ספק","סכום חובה","סכום זכות"])

        st_df.to_excel(writer, index=False, sheet_name="הוראת קבע ספקים")

    # ---- עיצוב ושורת 20001 ----
    wb_out = load_workbook(io.BytesIO(out_stream.getvalue()))
    for s in wb_out.worksheets:
        s.sheet_view.rightToLeft = True

    if "הוראת קבע ספקים" in wb_out.sheetnames:
        ws = wb_out["הוראת קבע ספקים"]
        headers = {cell.value: idx for idx, cell in enumerate(ws[1], start=1)}
        col_supplier = headers.get("מס' ספק")
        col_details  = headers.get("פרטים")
        col_amount   = headers.get("סכום")
        col_debit    = headers.get("סכום חובה")
        col_credit   = headers.get("סכום זכות")

        orange = PatternFill(start_color="FFDDBB", end_color="FFDDBB", fill_type="solid")
        if col_supplier:
            for r in range(2, ws.max_row+1):
                v = ws.cell(row=r, column=col_supplier).value
                if v in ("", None):
                    for c in range(1, ws.max_column+1):
                        ws.cell(row=r, column=c).fill = orange

        # מחיקת שורות 20001 קודמות (אם היו)
        dels = []
        for r in range(2, ws.max_row+1):
            v = ws.cell(row=r, column=col_supplier).value
            if v == 20001 or (isinstance(v,str) and v.strip() == "20001"):
                dels.append(r)
        for k, r in enumerate(dels):
            ws.delete_rows(r-k, 1)

        # סכום 20001 = סכום חובה של שורות שיש להן "מס' ספק"
        total_from_debit = 0.0
        for r in range(2, ws.max_row+1):
            sv = ws.cell(row=r, column=col_supplier).value
            if sv not in (None, ""):
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

# ---------------- קלט/הרצה ----------------
uploaded = st.file_uploader("בחרי קובץ אקסל מקור (xlsx)", type=["xlsx"])

if st.button("הרצה") and uploaded is not None:
    with st.spinner("מעבד..."):
        t3_ctx = st.session_state.get("t3_map")
        out_bytes, summary = process_workbook(uploaded.read(), t3_ctx=t3_ctx)
    st.success("מוכן! אפשר להוריד את הקובץ המעודכן.")
    st.dataframe(summary, use_container_width=True)
    st.download_button("⬇️ הורדת קובץ מעודכן",
                       data=out_bytes,
                       file_name="התאמות_והוראת_קבע_והעברות.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.caption("טיפ: בהתאמה 3, ההשוואה לביטוי בפרטים מבוצעת בנורמליזציה (ללא גרשים/מקפים/רווחים). 'תאריך פריקה' יכול לכלול זמן באותה עמודה.")
