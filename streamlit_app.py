# -*- coding: utf-8 -*-
"""
התאמות לקוחות – OV/RC (#1) + הוראות קבע (#2) + העברות (#3-ללא התאמת תאריך) + גיליון 'הוראת קבע ספקים'
היגיון התאמה 3 (מעודכן):
- צד בנק: קוד בנק 485, סכום בדף > 0, 'פרטים' מכיל "העב' במקבץ-נט", וללא דרישת התאמת תאריך.
           מסומן 3 אם הסכום שווה (בדיוק) לסכום 'אחרי ניכוי' מאוחד מהקובץ העזר (לפי אירוע/תאריך-שעה).
- צד ספרים: מסומן 3 אם 'אסמכתא 1' שווה לאחד ממספרי התשלום שמופיעים בקובץ העזר לאותו אירוע.
שאר הלוגיקות (#1, #2) + בניית גיליון 'הוראת קבע ספקים' – על בסיס הקובץ המקורי.  (שימור סגנון ו־RTL)
"""

import io, re, os, json
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# ---------------- UI (RTL) ----------------
st.set_page_config(page_title="התאמות – 1/2/3 (3 ללא תאריך)", page_icon="✅", layout="centered")
st.markdown("""
<style>
  html, body, [class*="css"] { direction: rtl; text-align: right; }
  .block-container { padding-top: 1rem; max-width: 1100px; }
</style>
""", unsafe_allow_html=True)
st.title("התאמות: 1–2–3 (התאמה 3 ללא דרישת תאריך)")

# ---------------- קבועים ----------------
MATCH_COL_CANDS = ["מס.התאמה","מס. התאמה","מס התאמה","מספר התאמה","התאמה"]
BANK_CODE_CANDS = ["קוד פעולת בנק","קוד פעולה","קוד פעולת"]
BANK_AMT_CANDS  = ["סכום בדף","סכום דף","סכום בבנק","סכום תנועת בנק"]
BOOKS_AMT_CANDS = ["סכום בספרים","סכום בספר","סכום ספרים"]
REF_CANDS       = ["אסמכתא 1","אסמכתא1","אסמכתא","אסמכתה"]
DATE_CANDS      = ["תאריך מאזן","תאריך ערך","תאריך"]
DETAILS_CANDS   = ["פרטים","תיאור","שם ספק"]

# קובץ עזר – זיהוי עמודות
AUX_DATE_KEYS   = ["תאריך","פריקה"]       # "תאריך פריקה" (עם שעה)
AUX_AMOUNT_KEYS = ["אחרי","ניכוי"]        # "אחרי ניכוי"
AUX_PAYNO_KEYS  = ["מס","תשלום"]          # "מס' תשלום"

# קבועי לוגיקה
TRANSFER_CODE   = 485
TRANSFER_PHRASE = "העב' במקבץ-נט"
STANDING_CODES  = {469, 515}
OVRC_CODES      = {120, 175}
AMOUNT_EPS      = 0.00  # התאמת סכומים מדויקת

# VLOOKUP persistent maps (כמו במקור)
RULES_FILE = "rules_store.json"
RAW_NAME_MAP = {}         # אפשר להשאיר ריק; הכל ניהול מה-UI/JSON
BASE_AMOUNT_MAP = {}

# ---------------- עזרים ----------------
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

# ---------------- כללי VLOOKUP – טעינה/שמירה ----------------
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

# ---------------- לוגיקה מרכזית ----------------
def process_workbook(main_bytes, aux_bytes=None):
    """מעבד את קובץ המקור + (אופציונלי) קובץ עזר להעברות, ומחזיר Bytes של אקסל מעודכן + סיכום."""
    wb_in = load_workbook(io.BytesIO(main_bytes), data_only=True, read_only=True)
    out_stream = io.BytesIO()
    summary_rows, standing_rows = [], []

    with pd.ExcelWriter(out_stream, engine="xlsxwriter") as writer:
        for ws in wb_in.worksheets:
            df = ws_to_df(ws)
            df_save = df.copy()
            if df.empty:
                pd.DataFrame().to_excel(writer, index=False, sheet_name=ws.title)
                continue

            # עמודות
            col_match     = exact_or_contains(df, MATCH_COL_CANDS) or df.columns[0]
            col_bank_code = exact_or_contains(df, BANK_CODE_CANDS)
            col_bank_amt  = exact_or_contains(df, BANK_AMT_CANDS)
            col_books_amt = exact_or_contains(df, BOOKS_AMT_CANDS)
            col_ref       = exact_or_contains(df, REF_CANDS)
            col_date      = exact_or_contains(df, DATE_CANDS)
            col_details   = exact_or_contains(df, DETAILS_CANDS)

            match_values = df_save[col_match].copy() if col_match in df_save.columns else pd.Series([0]*len(df_save))
            if match_values.isna().any(): match_values = match_values.fillna(0)

            # נרמול
            _date      = normalize_date(pd.to_datetime(df[col_date], errors="coerce")) if col_date else pd.Series([pd.NaT]*len(df))
            _bank_amt  = to_number(df[col_bank_amt])  if col_bank_amt  else pd.Series([np.nan]*len(df))
            _books_amt = to_number(df[col_books_amt]) if col_books_amt else pd.Series([np.nan]*len(df))
            _bank_code = to_number(df[col_bank_code]) if col_bank_code else pd.Series([np.nan]*len(df))
            _ref       = df[col_ref].astype(str).fillna("") if col_ref else pd.Series([""]*len(df))
            _details   = df[col_details].astype(str).fillna("") if col_details else pd.Series([""]*len(df))

            # ===== התאמה 1 – OV/RC קפדנית 1:1 =====
            applied_ovrc = False; pairs = 0
            if all([col_bank_code, col_bank_amt, col_books_amt, col_ref, col_date]):
                applied_ovrc = True
                books_candidates = [
                    j for j in range(len(df))
                    if pd.notna(_books_amt.iat[j]) and _books_amt.iat[j] > 0
                    and pd.notna(_date.iat[j]) and ref_starts_with_ov_rc(_ref.iat[j])
                ]
                bank_keys, books_keys = {}, {}
                for i in range(len(df)):
                    if pd.notna(_bank_code.iat[i]) and int(_bank_code.iat[i]) in OVRC_CODES \
                       and pd.notna(_bank_amt.iat[i]) and _bank_amt.iat[i] < 0 \
                       and pd.notna(_date.iat[i]):
                        k = (round(abs(float(_bank_amt.iat[i])),2), _date.iat[i])
                        bank_keys.setdefault(k, []).append(i)
                for j in books_candidates:
                    k = (round(abs(float(_books_amt.iat[j])),2), _date.iat[j])
                    books_keys.setdefault(k, []).append(j)
                for k, b_idx in bank_keys.items():
                    if len(b_idx) == 1 and len(books_keys.get(k, [])) == 1:
                        i = b_idx[0]; j = books_keys[k][0]
                        if match_values.iat[i] in (0,2) and match_values.iat[j] in (0,2):
                            match_values.iat[i] = 1; match_values.iat[j] = 1; pairs += 1

            # ===== התאמה 2 – הוראות קבע (469/515) =====
            applied_standing = False; flagged = 0
            if all([col_bank_code, col_details, col_bank_amt]):
                applied_standing = True
                for i in range(len(df)):
                    code = _bank_code.iat[i]
                    if pd.notna(code) and int(code) in STANDING_CODES:
                        if match_values.iat[i] in (0,):
                            match_values.iat[i] = 2
                            flagged += 1
                            standing_rows.append({"פרטים": _details.iat[i],
                                                  "סכום": float(_bank_amt.iat[i]) if pd.notna(_bank_amt.iat[i]) else np.nan})

            # כתיבה ראשונית של הגיליון עם 1–2
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

        # ===== בניית גיליון 'הוראת קבע ספקים' =====
        st_df = pd.DataFrame(standing_rows)
        if not st_df.empty:
            # מיפוי ספק: קודם לפי שם (contains), אח"כ לפי סכום מוחלט
            def map_supplier(name, amount):
                s = normalize_text(name)
                # by name – התאמה מלאה או contains יורד לפי אורך
                for key in sorted(st.session_state.name_map.keys(), key=len, reverse=True):
                    if key and key in s:
                        return st.session_state.name_map[key]
                # by amount – ערך מוחלט
                try:
                    val = round(abs(float(amount)), 2)
                    return st.session_state.amount_map.get(val, "")
                except Exception:
                    return ""

            st_df["מס' ספק"] = st_df.apply(lambda r: map_supplier(r["פרטים"], r["סכום"]), axis=1)
            st_df["סכום חובה"] = st_df["סכום"].apply(lambda x: abs(x) if pd.notna(x) else 0.0)
            st_df["סכום זכות"] = 0.0

            total_hova_with_supplier = st_df.loc[st_df["מס' ספק"].astype(str).str.len()>0, "סכום חובה"].sum()
            vk = st_df[["פרטים","סכום","מס' ספק","סכום חובה","סכום זכות"]].copy()
            # שורת סיכום 20001 בזכות בלבד (כמו שביקשת מוקדם יותר)
            vk = pd.concat([vk, pd.DataFrame([{
                "פרטים":"סה\"כ זכות – עם מס' ספק",
                "סכום":0.0,
                "מס' ספק":20001,
                "סכום חובה":0.0,
                "סכום זכות":round(float(total_hova_with_supplier),2)
            }])], ignore_index=True)
        else:
            vk = pd.DataFrame(columns=["פרטים","סכום","מס' ספק","סכום חובה","סכום זכות"])

        vk.to_excel(writer, index=False, sheet_name="הוראת קבע ספקים")

    # ===== עיצוב + התאמה 3 (לאחר כתיבה) =====
    wb_out = load_workbook(io.BytesIO(out_stream.getvalue()))
    # RTL + צביעה כתומה לשורות ללא מס' ספק
    for s in wb_out.worksheets:
        s.sheet_view.rightToLeft = True
    if "הוראת קבע ספקים" in wb_out.sheetnames:
        ws_so = wb_out["הוראת קבע ספקים"]
        headers = {cell.value: idx for idx, cell in enumerate(ws_so[1], start=1)}
        col_supplier = headers.get("מס' ספק")
        if col_supplier:
            orange = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            for r in range(2, ws_so.max_row+1):
                v = ws_so.cell(row=r, column=col_supplier).value
                if v in ("", None):
                    for c in range(1, ws_so.max_column+1):
                        ws_so.cell(row=r, column=c).fill = orange
        for cell in ws_so[ws_so.max_row]:
            cell.font = Font(bold=True)

    # ===== התאמה 3 – ללא התאמת תאריך בצד הבנק =====
    if aux_bytes is not None:
        aux_xl = load_workbook(io.BytesIO(aux_bytes), data_only=True, read_only=True)
        aux_ws = aux_xl.worksheets[0]
        aux_df = ws_to_df(aux_ws)

        # עמודות בעזר
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

        # קיבוץ לפי חותמת זמן מלאה (אירוע) → סכום 'אחרי ניכוי'
        grouped  = (pd.DataFrame({"_dt": aux_dt, "_amt": aux_amt})
                      .dropna(subset=["_dt"])
                      .groupby("_dt")["_amt"].sum().round(2).to_dict())
        pays_by_dt = (pd.DataFrame({"_dt": aux_dt, "_pay": aux_pay})
                        .groupby("_dt")["_pay"].apply(lambda s: set(s.dropna().astype(str))).to_dict())

        # טען DataSheet מהקובץ שזה עתה ייצרנו
        ds_ws = wb_out["DataSheet"] if "DataSheet" in wb_out.sheetnames else wb_out.worksheets[0]
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

        # בנק – מועמדים עם 485, סכום>0, טקסט העברה; **אין תנאי על תאריך**:
        bank_candidates = (ds_code == TRANSFER_CODE) & (ds_amt > 0) & (ds_details.str.contains(TRANSFER_PHRASE, na=False))

        # התאמה: לפי סכום בלבד (לעבור על כל אירוע מהעזר)
        mark_bank = set(); mark_books = set()
        for dt, gsum in grouped.items():
            # צד בנק: סכום שווה מדויק (±0.00), בלי לבדוק תאריך
            hits = ds_df.index[ bank_candidates & (ds_amt.abs() == abs(gsum)) ].tolist()
            if hits:
                mark_bank.update(hits)
                # צד ספרים: לפי 'מס' תשלום' של אותו אירוע
                payset = pays_by_dt.get(dt, set())
                if payset:
                    link_rows = ds_df.index[ ds_ref.astype(str).isin(payset) ].tolist()
                    mark_books.update(link_rows)

        for i in sorted(mark_bank):
            if ds_match.iat[i] in (0,2):  # לא לדרוס 1/2
                ds_match.iat[i] = 3
        for j in sorted(mark_books):
            if ds_match.iat[j] in (0,2):
                ds_match.iat[j] = 3

        # כתיבה חזרה ל-DataSheet
        ds_df_out = ds_df.copy()
        ds_df_out[ds_col_match] = ds_match
        # נקה והחזר
        for _ in range(ds_ws.max_row, 1, -1):
            ds_ws.delete_rows(2, 1)
        for r in ds_df_out.itertuples(index=False):
            ds_ws.append(list(r))

    # החזרת Bytes + סיכום
    final_bytes = io.BytesIO()
    wb_out.save(final_bytes)
    summary_df = pd.DataFrame(summary_rows)
    return final_bytes.getvalue(), summary_df

# ---------------- UI ----------------
colA, colB = st.columns([2,2])
uploaded_main = colA.file_uploader("קובץ מקור (xlsx) – כולל DataSheet", type=["xlsx"])
uploaded_aux  = colB.file_uploader("(אופציונלי) קובץ עזר להעברות – 'תאריך פריקה', 'אחרי ניכוי', 'מס' תשלום'", type=["xlsx"])

if st.button("הרצה"):
    if uploaded_main is None:
        st.error("נא להעלות קובץ מקור.")
    else:
        with st.spinner("מעבד..."):
            out_bytes, summary = process_workbook(uploaded_main.read(),
                                                  uploaded_aux.read() if uploaded_aux else None)
        st.success("מוכן להורדה.")
        if not summary.empty:
            st.dataframe(summary, use_container_width=True)
        st.download_button("⬇️ הורד קובץ", data=out_bytes,
                           file_name="התאמות_1_2_3_ללא_תאריך.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.divider()
st.subheader("🔎 VLOOKUP – הוראת קבע ספקים (ניהול כללים)")
# ניהול מפות (שם/סכום → מס' ספק) עם שמירה ל-rules_store.json
name_map_df = pd.DataFrame({"by_name": list(st.session_state.name_map.keys()),
                            "מס' ספק": list(st.session_state.name_map.values())})
amount_map_df = pd.DataFrame({"סכום": list(st.session_state.amount_map.keys()),
                              "מס' ספק": list(st.session_state.amount_map.values())}).sort_values("סכום")
st.write("כלל לפי שם:")
nm_col1, nm_col2 = st.columns([2,1])
_nm = nm_col1.text_input("מחרוזת מתוך 'פרטים' (contains)")
_sp = nm_col2.text_input("מס' ספק")
if st.button("➕ הוסף/עדכן לפי שם"):
    k = normalize_text(_nm)
    if k and _sp:
        st.session_state.name_map[k] = _sp
        save_rules_to_disk(st.session_state.name_map, st.session_state.amount_map)
        st.success("נשמר.")
st.dataframe(name_map_df, use_container_width=True, height=200)

st.write("כלל לפי סכום:")
am_col1, am_col2 = st.columns([1,1])
_amt = am_col1.number_input("סכום (ערך מוחלט)", step=0.01, format="%.2f")
_sp2 = am_col2.text_input("מס' ספק", key="vk_amt_sup")
if st.button("➕ הוסף/עדכן לפי סכום"):
    key_amt = round(abs(float(_amt)), 2)
    if key_amt and _sp2:
        st.session_state.amount_map[key_amt] = _sp2
        save_rules_to_disk(st.session_state.name_map, st.session_state.amount_map)
        st.success("נשמר.")
st.dataframe(amount_map_df, use_container_width=True, height=200)
