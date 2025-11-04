# -*- coding: utf-8 -*-
"""
התאמות לקוחות – OV/RC (#1) + הוראות קבע (#2, עם VLOOKUP נשמר) + העברות (#3) + שיקים ספקים (#4)
— גרסה עם כלל 4 מתוקן: התאמה לפי ערך מוחלט ±0.50, בלי בדיקת תאריך,
   מיפוי אסמכתאות: בנק(אסמכתא 1, קוד 493) ↔ ספרים(אסמכתא 2) כאשר אסמכתא 1 בספרים מתחילה CH…

שומר כללי VLOOKUP של "הוראת קבע ספקים" לקובץ rules_store.json (by_name / by_amount)
"""

import io, re, os, json
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# ---------------- UI (RTL + בסיס) ----------------
st.set_page_config(page_title="התאמות לקוחות – 1/2/3/4", page_icon="✅", layout="centered")
st.markdown(
    """
<style>
  html, body, [class*="css"] { direction: rtl; text-align: right; }
  .block-container { padding-top: 1rem; }
</style>
""",
    unsafe_allow_html=True,
)

st.title("התאמות לקוחות – OV/RC + הוראות קבע (VLOOKUP נשמר) + העברות + שיקים ספקים (#4)")

# -------- כללי VLOOKUP ברירת-מחדל --------
RAW_NAME_MAP = {
    "בזק בינלאומי ב": 30006,
    "פרי ירוחם חב'": 34714,
    "סלקום ישראל בע": 30055,
    "בזק-הוראות קבע": 34746,
    "דרך ארץ הייווי": 34602,
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
    "הו\"ק הלוואה קרן": 23001,
    # כלליים
    "עיריית אשדוד": 30056,
    "ישראכרט מור": 34002,
}
BASE_AMOUNT_MAP = {
    8520.0: 30247,    # שלמה טפר
    10307.3: 30038,   # נמרוד תבור עו"ד
}

# -------- מזהי עמודות אפשריים (עברית/וריאציות) --------
MATCH_COL_CANDS = ["מס.התאמה","מס. התאמה","מס התאמה","מספר התאמה","התאמה"]
BANK_CODE_CANDS = ["קוד פעולת בנק","קוד פעולה","קוד פעולת"]
BANK_AMT_CANDS  = ["סכום בדף","סכום דף","סכום בבנק","סכום תנועת בנק"]
BOOKS_AMT_CANDS = ["סכום בספרים","סכום בספר","סכום ספרים"]
REF_CANDS       = ["אסמכתא 1","אסמכתא1","אסמכתא","אסמכתה"]
REF2_CANDS      = ["אסמכתא 2","אסמכתא2","אסמכתא-2","אסמכתה 2","אסמכתא2 "]
DATE_CANDS      = ["תאריך מאזן","תאריך ערך","תאריך"]
DETAILS_CANDS   = ["פרטים","תיאור","שם ספק"]

# התאמה 3 – עמודות בקובץ העזר
AUX_DATE_KEYS   = ["תאריך","פריקה"]       # "תאריך פריקה" (תאריך+שעה)
AUX_AMOUNT_KEYS = ["אחרי","ניכוי"]        # "אחרי ניכוי"
AUX_PAYNO_KEYS  = ["מס","תשלום"]          # "מס' תשלום"

# ביטויים/קבועים ללוגיקה
RULES_FILE = "rules_store.json"
TRANSFER_CODE = 485
TRANSFER_PHRASE = "העב' במקבץ-נט"
STANDING_CODES = {469, 515}
OVRC_CODES = {120, 175}
AMOUNT_EPS = 0.00  # התאמה מדויקת בסכומים

# כלל 4 – שיקים ספקים
RULE4_CODE = 493
RULE4_EPS = 0.50   # ±₪0.50

# ---------------- עזר לנרמול ----------------
def normalize_text(s):
    if s is None: return ""
    t = str(s)
    t = t.replace("'", "").replace('"', "").replace("’", "").replace("`", "")
    t = t.replace("-", " ").replace("–", " ").replace("־", " ")
    t = re.sub(r"\s+", " ", t)
    return t.strip()

def normalize_date(series):
    def f(x):
        if pd.isna(x): return pd.NaT
        if isinstance(x,(pd.Timestamp, datetime)): return pd.Timestamp(x.date())
        return pd.to_datetime(x, dayfirst=True, errors="coerce").normalize()
    return series.apply(f)

def to_number(series):
    s = (series.astype(str)
         .str.replace(",","", regex=False)
         .str.replace("₪","", regex=False)
         .str.replace("\u200f","", regex=False)
         .str.replace("\u200e","", regex=False)
         .str.strip())
    return pd.to_numeric(s, errors="coerce")

def ref_starts_with_ov_rc(val):
    t = (str(val) if val is not None else "").strip().upper()
    return t.startswith("OV") or t.startswith("RC")

def exact_or_contains(df, names):
    for n in names:
        if n in df.columns:
            return n
    for n in names:
        for c in df.columns:
            if isinstance(c, str) and n in c:
                return c
    return None

def ws_to_df(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return pd.DataFrame()
    header = None; start = 0
    for i, r in enumerate(rows):
        if any(x is not None for x in r):
            header = [str(x).strip() if x is not None else "" for x in r]; start = i+1; break
    if header is None:
        return pd.DataFrame()
    data = [tuple(list(row)[:len(header)]) for row in rows[start:]]
    return pd.DataFrame(data, columns=header)

# ---------------- טעינת כללים/שמירה מתמשכת ----------------
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
    # בסיס אם אין קובץ
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

# ייצוא/ייבוא של כללי VLOOKUP
with st.expander("⚙️ עדכון – כללי VLOOKUP קבועים ומורחבים (עם שמירה)", expanded=False):
    st.write("עדכון לפי **פרטים (שם)** או לפי **סכום**. נשמר ל־`rules_store.json` לשימוש חוזר.")

    mode = st.radio("סוג עדכון", ["לפי פרטים (שם)", "לפי סכום"], horizontal=True)

    if mode == "לפי פרטים (שם)":
        name_input = st.text_input("פרטים (כמו בדף הבנק)")
        supplier_input = st.text_input("מס' ספק (יכול להיות גם טקסט, למשל 67-1)")
        cols = st.columns([1,1,1,1])
        if cols[0].button("➕ הוסף/עדכן"):
            k = normalize_text(name_input)
            if k and supplier_input:
                st.session_state.name_map[k] = supplier_input
                save_rules_to_disk(st.session_state.name_map, st.session_state.amount_map)
                st.success(f"הכלל נשמר: '{k}' → {supplier_input}")
        if cols[1].button("🗑️ מחיקה"):
            k = normalize_text(name_input)
            if k in st.session_state.name_map:
                del st.session_state.name_map[k]
                save_rules_to_disk(st.session_state.name_map, st.session_state.amount_map)
                st.warning(f"הכלל נמחק: '{k}'")
        if cols[2].button("💾 שמור ידנית"):
            save_rules_to_disk(st.session_state.name_map, st.session_state.amount_map)
            st.info("נשמר לקובץ rules_store.json")
        st.dataframe(pd.DataFrame({"by_name": list(st.session_state.name_map.keys()),
                                   "מס' ספק": list(st.session_state.name_map.values())}),
                     use_container_width=True, height=260)

    else:  # לפי סכום
        amount_input = st.number_input("סכום (יחושב בערך מוחלט)", step=0.01, format="%.2f")
        supplier_input2 = st.text_input("מס' ספק", key="amount_supplier")
        cols = st.columns([1,1,1,1])
        if cols[0].button("➕ הוסף/עדכן", key="add_amount"):
            key_amt = round(abs(float(amount_input)), 2)
            if key_amt and supplier_input2:
                st.session_state.amount_map[key_amt] = supplier_input2
                save_rules_to_disk(st.session_state.name_map, st.session_state.amount_map)
                st.success(f"הכלל נשמר: {key_amt} → {supplier_input2}")
        if cols[1].button("🗑️ מחיקה", key="del_amount"):
            key_amt = round(abs(float(amount_input)), 2)
            if key_amt in st.session_state.amount_map:
                del st.session_state.amount_map[key_amt]
                save_rules_to_disk(st.session_state.name_map, st.session_state.amount_map)
                st.warning(f"הכלל נמחק: {key_amt}")
        if cols[2].button("💾 שמור ידנית", key="save_amount"):
            save_rules_to_disk(st.session_state.name_map, st.session_state.amount_map)
            st.info("נשמר לקובץ rules_store.json")
        st.dataframe(pd.DataFrame({"סכום": list(st.session_state.amount_map.keys()),
                                   "מס' ספק": list(st.session_state.amount_map.values())})
                     .sort_values("סכום"), use_container_width=True, height=260)

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
            try:
                data = json.loads(uploaded_rules.read().decode("utf-8"))
                nm = { normalize_text(k): v for k, v in data.get("name_map", {}).items() }
                am = { float(k): v for k, v in data.get("amount_map", {}).items() }
                st.session_state.name_map = nm
                st.session_state.amount_map = am
                save_rules_to_disk(nm, am)
                st.success("הכללים יובאו ונשמרו בהצלחה.")
            except Exception as e:
                st.error(f"שגיאה בייבוא: {e}")
    if c4.button("🔄 שמור עדכונים לשימוש עתידי"):
        if save_rules_to_disk(st.session_state.name_map, st.session_state.amount_map):
            st.success("נשמר בהצלחה ל־rules_store.json")

st.divider()

# ---------------- פונקציות לוגיקה ----------------

def process_workbook(main_bytes, aux_bytes=None):
    """מעבד את קובץ המקור + (אופציונלית) קובץ עזר להעברות, ומחזיר Bytes של אקסל מעודכן + טבלת סיכום."""
    # טען את חוברת המקור
    wb_in = load_workbook(io.BytesIO(main_bytes), data_only=True, read_only=True)

    out_stream = io.BytesIO()
    summary_rows, standing_rows = [], []

    # ===== שלב 1: מעבר על כל הגיליונות =====
    with pd.ExcelWriter(out_stream, engine="xlsxwriter") as writer:
        for ws in wb_in.worksheets:
            df = ws_to_df(ws)
            df_save = df.copy()
            if df.empty:
                pd.DataFrame().to_excel(writer, index=False, sheet_name=ws.title)
                continue

            # איתור עמודות
            col_match     = exact_or_contains(df, MATCH_COL_CANDS) or df.columns[0]
            col_bank_code = exact_or_contains(df, BANK_CODE_CANDS)
            col_bank_amt  = exact_or_contains(df, BANK_AMT_CANDS)
            col_books_amt = exact_or_contains(df, BOOKS_AMT_CANDS)
            col_ref       = exact_or_contains(df, REF_CANDS)
            col_ref2      = exact_or_contains(df, REF2_CANDS)
            col_date      = exact_or_contains(df, DATE_CANDS)
            col_details   = exact_or_contains(df, DETAILS_CANDS)

            match_values = df_save[col_match].copy() if col_match in df_save.columns else pd.Series([0]*len(df_save))
            if match_values.isna().any():
                match_values = match_values.fillna(0)

            # נרמול שדות
            _date      = normalize_date(pd.to_datetime(df[col_date], errors="coerce")) if col_date else pd.Series([pd.NaT]*len(df))
            _bank_amt  = to_number(df[col_bank_amt])  if col_bank_amt  else pd.Series([np.nan]*len(df))
            _books_amt = to_number(df[col_books_amt]) if col_books_amt else pd.Series([np.nan]*len(df))
            _bank_code = to_number(df[col_bank_code]) if col_bank_code else pd.Series([np.nan]*len(df))
            _ref       = df[col_ref].astype(str).fillna("") if col_ref else pd.Series([""]*len(df))
            _ref2      = df[col_ref2].astype(str).fillna("") if col_ref2 else pd.Series([""]*len(df))
            _details   = df[col_details].astype(str).fillna("") if col_details else pd.Series([""]*len(df))

            # ===== התאמה 1 – OV/RC קפדנית 1:1 =====
            applied_ovrc = False; pairs = 0
            if all([col_bank_code, col_bank_amt, col_books_amt, col_ref, col_date]):
                applied_ovrc = True
                # מועמדים ספרים: +, OV/RC
                books_candidates = [
                    j for j in range(len(df))
                    if pd.notna(_books_amt.iat[j]) and _books_amt.iat[j] > 0
                    and pd.notna(_date.iat[j]) and ref_starts_with_ov_rc(_ref.iat[j])
                ]
                # קבוצות לפי (סכום מוחלט, תאריך) – חייב 1:1
                bank_keys  = {}
                books_keys = {}
                for i in range(len(df)):
                    if pd.notna(_bank_code.iat[i]) and int(_bank_code.iat[i]) in OVRC_CODES \
                       and pd.notna(_bank_amt.iat[i]) and _bank_amt.iat[i] < 0 \
                       and pd.notna(_date.iat[i]):
                        k = (round(abs(float(_bank_amt.iat[i])),2), _date.iat[i])
                        bank_keys.setdefault(k, []).append(i)
                for j in books_candidates:
                    k = (round(abs(float(_books_amt.iat[j])),2), _date.iat[j])
                    books_keys.setdefault(k, []).append(j)
                # התאמה קפדנית: רק מפתחות שמופיעים פעם אחת בכל צד
                for k, b_idx in bank_keys.items():
                    if len(b_idx) == 1 and len(books_keys.get(k, [])) == 1:
                        i = b_idx[0]; j = books_keys[k][0]
                        if match_values.iat[i] in (0,2) and match_values.iat[j] in (0,2):  # לא לדרוס 3/1
                            match_values.iat[i] = 1
                            match_values.iat[j] = 1
                            pairs += 1

            # ===== התאמה 2 – הוראות קבע (469/515) =====
            applied_standing = False; flagged = 0
            if all([col_bank_code, col_details, col_bank_amt]):
                applied_standing = True
                for i in range(len(df)):
                    code = _bank_code.iat[i]
                    if pd.notna(code) and int(code) in STANDING_CODES:
                        if match_values.iat[i] in (0,):   # לא לדרוס 1/3
                            match_values.iat[i] = 2
                            flagged += 1
                            standing_rows.append({"פרטים": _details.iat[i],
                                                  "סכום": float(_bank_amt.iat[i]) if pd.notna(_bank_amt.iat[i]) else np.nan})

            # סיום גיליון
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

        # ===== גיליון הוראת קבע ספקים (מהשורות שסומנו 2) =====
        st_df = pd.DataFrame(standing_rows)
        if not st_df.empty:
            def map_supplier(name, amount):
                # 1) לפי שם
                s = normalize_text(name)
                if s in st.session_state.name_map:
                    return st.session_state.name_map[s]
                for key in sorted(st.session_state.name_map.keys(), key=len, reverse=True):
                    if key and key in s:
                        return st.session_state.name_map[key]
                # 2) לפי סכום מוחלט
                try:
                    val = round(abs(float(amount)), 2)
                    return st.session_state.amount_map.get(val, "")
                except Exception:
                    return ""

            st_df["מס' ספק"] = st_df.apply(lambda r: map_supplier(r["פרטים"], r["סכום"]), axis=1)
            # שורות רגילות: חובה בלבד; שורת סיכום 20001 תהיה בזכות
            st_df["סכום חובה"] = st_df["סכום"].apply(lambda x: abs(x) if pd.notna(x) else 0.0)
            st_df["סכום זכות"] = 0.0

            # סכום חובה רק לשורות שיש להן מס' ספק
            total_hova_with_supplier = st_df.loc[st_df["מס' ספק"].astype(str).str.len()>0, "סכום חובה"].sum()

            vk = st_df[["פרטים","סכום","מס' ספק","סכום חובה","סכום זכות"]].copy()
            # שורת סיכום 20001 – זכות בלבד
            vk = pd.concat([vk, pd.DataFrame([{
                "פרטים":"סה\"כ זכות – עם מס' ספק",
                "סכום":0.0,
                "מס' ספק":20001,
                "סכום חובה":0.0,
                "סכום זכות":round(total_hova_with_supplier,2)
            }])], ignore_index=True)
        else:
            vk = pd.DataFrame(columns=["פרטים","סכום","מס' ספק","סכום חובה","סכום זכות"])

        vk.to_excel(writer, index=False, sheet_name="הוראת קבע ספקים")

    # ===== שלב 2: עיצוב (RTL, צביעה, שורת 20001 מודגשת) =====
    wb_out = load_workbook(io.BytesIO(out_stream.getvalue()))
    for s in wb_out.worksheets:
        s.sheet_view.rightToLeft = True

    ws_so = wb_out["הוראת קבע ספקים"]
    headers = {cell.value: idx for idx, cell in enumerate(ws_so[1], start=1)}
    col_supplier = headers.get("מס' ספק")

    orange = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    if col_supplier:
        # צביעה כתומה לשורות ללא מס' ספק (למעט השורה האחרונה אם היא 20001)
        for r in range(2, ws_so.max_row+1):
            v = ws_so.cell(row=r, column=col_supplier).value
            if v in ("", None):
                for c in range(1, ws_so.max_column+1):
                    ws_so.cell(row=r, column=c).fill = orange

    # מחיקה של 20001 כפולים אם קיימים
    dels = []
    for r in range(2, ws_so.max_row+1):
        v = ws_so.cell(row=r, column=col_supplier).value
        if v == 20001 or (isinstance(v,str) and v.strip() == "20001"):
            dels.append(r)
    for k, r in enumerate(dels[:-1]):  # נשאיר את האחרון
        ws_so.delete_rows(r-k, 1)

    # הדגשה לשורה האחרונה (סיכום)
    for cell in ws_so[ws_so.max_row]:
        cell.font = Font(bold=True)

    # ===== שלב 3: התאמה 3 (העברות ספקים) =====
    if aux_bytes is not None:
        aux_xl = load_workbook(io.BytesIO(aux_bytes), data_only=True, read_only=True)
        aux_ws = aux_xl.worksheets[0]
        aux_df = ws_to_df(aux_ws)

        def pick_col(df, keys):
            for c in df.columns:
                s = str(c)
                if all(k in s for k in keys):
                    return c
            return None

        c_dt  = pick_col(aux_df, AUX_DATE_KEYS)
        c_amt = pick_col(aux_df, AUX_AMOUNT_KEYS)
        c_pay = pick_col(aux_df, AUX_PAYNO_KEYS)

        aux_dt  = pd.to_datetime(aux_df[c_dt], errors="coerce") if c_dt else pd.Series([pd.NaT]*len(aux_df))
        aux_amt = pd.to_numeric(aux_df[c_amt], errors="coerce").round(2) if c_amt else pd.Series([np.nan]*len(aux_df))
        aux_pay = aux_df[c_pay].astype(str).str.strip() if c_pay else pd.Series([""]*len(aux_df))

        grouped = (pd.DataFrame({"_dt": aux_dt, "_amt": aux_amt})
                     .dropna(subset=["_dt"])\
                     .groupby("_dt")["_amt"].sum().round(2)
                     .to_dict())
        pays_by_dt = (pd.DataFrame({"_dt": aux_dt, "_pay": aux_pay})
                        .groupby("_dt")["_pay"]
                        .apply(lambda s: set(s.dropna().astype(str)))
                        .to_dict())

        # נטען DF של DataSheet
        ds_ws = wb_out["DataSheet"]
        ds_df = ws_to_df(ds_ws)

        ds_col_match     = exact_or_contains(ds_df, MATCH_COL_CANDS) or ds_df.columns[0]
        ds_col_bank_code = exact_or_contains(ds_df, BANK_CODE_CANDS)
        ds_col_bank_amt  = exact_or_contains(ds_df, BANK_AMT_CANDS)
        ds_col_details   = exact_or_contains(ds_df, DETAILS_CANDS)
        ds_col_ref       = exact_or_contains(ds_df, REF_CANDS)

        ds_match   = pd.to_numeric(ds_df[ds_col_match], errors="coerce").fillna(0).astype(int)
        ds_code    = to_number(ds_df[ds_col_bank_code])
        ds_amt     = to_number(ds_df[ds_col_bank_amt]).round(2)
        ds_details = ds_df[ds_col_details].astype(str).fillna("")
        ds_ref     = ds_df[ds_col_ref].astype(str).fillna("")

        bank_candidates = (ds_code == TRANSFER_CODE) & \
                          (ds_amt > 0) & \
                          (ds_details.str.contains(TRANSFER_PHRASE, na=False))

        mark_bank = set(); mark_link = set()
        for dt, gsum in grouped.items():
            hits = ds_df.index[bank_candidates & (ds_amt.abs().round(2) == abs(gsum))].tolist()
            if hits:
                mark_bank.update(hits)
                payset = pays_by_dt.get(dt, set())
                if payset:
                    link_rows = ds_df.index[ds_ref.astype(str).isin(payset)].tolist()
                    mark_link.update(link_rows)

        for i in sorted(mark_bank):
            if ds_match.iat[i] in (0,2):
                ds_match.iat[i] = 3
        for i in sorted(mark_link):
            if ds_match.iat[i] in (0,2):
                ds_match.iat[i] = 3

        # כתיבה חזרה לגליון DataSheet בלבד
        ds_df_out = ds_df.copy()
        ds_df_out[ds_col_match] = ds_match
        # מחליפים תוכן גליון
        for _ in range(ds_ws.max_row, 1, -1):
            ds_ws.delete_rows(2, 1)
        for r in ds_df_out.itertuples(index=False):
            ds_ws.append(list(r))

    # ===== שלב 4: התאמה 4 – שיקים ספקים (מתוקן) =====
    # נטען DF של DataSheet (ייתכן שכבר נטען בשלב 3)
    ds_ws = wb_out["DataSheet"] if "DataSheet" in wb_out.sheetnames else wb_out.worksheets[0]
    ds_df = ws_to_df(ds_ws)

    ds_col_match     = exact_or_contains(ds_df, MATCH_COL_CANDS) or ds_df.columns[0]
    ds_col_bank_code = exact_or_contains(ds_df, BANK_CODE_CANDS)
    ds_col_bank_amt  = exact_or_contains(ds_df, BANK_AMT_CANDS)
    ds_col_books_amt = exact_or_contains(ds_df, BOOKS_AMT_CANDS)
    ds_col_details   = exact_or_contains(ds_df, DETAILS_CANDS)
    ds_col_ref       = exact_or_contains(ds_df, REF_CANDS)
    ds_col_ref2      = exact_or_contains(ds_df, REF2_CANDS)

    ds_match   = pd.to_numeric(ds_df[ds_col_match], errors="coerce").fillna(0).astype(int)
    ds_code    = to_number(ds_df[ds_col_bank_code])
    ds_bamt    = to_number(ds_df[ds_col_bank_amt]).round(2)
    ds_aamt    = to_number(ds_df[ds_col_books_amt]).round(2) if ds_col_books_amt else pd.Series([np.nan]*len(ds_df))
    ds_ref     = ds_df[ds_col_ref].astype(str).fillna("")
    ds_ref2    = ds_df[ds_col_ref2].astype(str).fillna("") if ds_col_ref2 else pd.Series([""]*len(ds_df))

    def strip_leading_zeros_digits_only(s):
        s = re.sub(r"\D", "", str(s))
        return s.lstrip("0") or "0"

    # בנק: קוד 493 + יש אסמכתא 1 (כל סימן); ספרים: אסמכתא 1 מתחיל CH + יש אסמכתא 2
    bank_idx = [i for i in range(len(ds_df))
                if pd.notna(ds_code.iat[i]) and int(ds_code.iat[i]) == RULE4_CODE
                and str(ds_ref.iat[i]).strip() != ""
                and pd.notna(ds_bamt.iat[i])]

    books_idx = [j for j in range(len(ds_df))
                 if str(ds_ref.iat[j]).upper().startswith("CH")
                 and str(ds_ref2.iat[j]).strip() != ""
                 and pd.notna(ds_aamt.iat[j])]

    used_books = set()
    for i in bank_idx:
        if ds_match.iat[i] != 0:  # לא לדרוס התאמות 1–3
            continue
        ref_b_clean = strip_leading_zeros_digits_only(ds_ref.iat[i])
        amt_b = abs(float(ds_bamt.iat[i])) if pd.notna(ds_bamt.iat[i]) else np.nan
        candidates = []
        for j in books_idx:
            if j in used_books:
                continue
            if ds_match.iat[j] != 0:
                continue
            ref2_clean = strip_leading_zeros_digits_only(ds_ref2.iat[j])
            if ref_b_clean != ref2_clean:
                continue
            amt_a = abs(float(ds_aamt.iat[j])) if pd.notna(ds_aamt.iat[j]) else np.nan
            if pd.isna(amt_a) or pd.isna(amt_b):
                continue
            if abs(amt_a - amt_b) <= RULE4_EPS:
                candidates.append(j)
        if candidates:
            j = candidates[0]  # מסמנים אחד בלבד
            ds_match.iat[i] = 4
            ds_match.iat[j] = 4
            used_books.add(j)

    # כתיבה חזרה ל-DataSheet
    ds_df_out = ds_df.copy()
    ds_df_out[ds_col_match] = ds_match
    for _ in range(ds_ws.max_row, 1, -1):
        ds_ws.delete_rows(2, 1)
    for r in ds_df_out.itertuples(index=False):
        ds_ws.append(list(r))

    # החזרת Bytes + סיכום
    final_bytes = io.BytesIO()
    wb_out.save(final_bytes)
    summary_df = pd.DataFrame(summary_rows)
    return final_bytes.getvalue(), summary_df

# ---------------- UI – העלאות והרצה ----------------
colA, colB = st.columns([2,2])
uploaded_main = colA.file_uploader("בחרי קובץ אקסל מקור (xlsx) – כולל DataSheet", type=["xlsx"])
uploaded_aux  = colB.file_uploader("(אופציונלי) קובץ עזר להעברות – 'תאריך פריקה' (תאריך+שעה), 'אחרי ניכוי', 'מס' תשלום'", type=["xlsx"])

if st.button("הרצה"):
    if uploaded_main is None:
        st.error("נא להעלות קובץ מקור (xlsx) עם גיליון DataSheet.")
    else:
        with st.spinner("מעבד..."):
            out_bytes, summary = process_workbook(uploaded_main.read(), uploaded_aux.read() if uploaded_aux else None)
        st.success("מוכן! אפשר להוריד את הקובץ המעודכן.")
        if not summary.empty:
            st.dataframe(summary, use_container_width=True)
        st.download_button("⬇️ הורדת קובץ מעודכן", data=out_bytes, file_name="התאמות_1_2_3_4.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.caption("טיפ: כללי VLOOKUP נשמרים אוטומטית ל־rules_store.json. אפשר גם לייצא/לייבא JSON לגיבוי.")
