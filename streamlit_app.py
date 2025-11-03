# streamlit_app.py
# -*- coding: utf-8 -*-
import io, os, re, json
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

# ---------- UI בסיס ----------
st.set_page_config(page_title="התאמות 1+2+3+4 + הוראת קבע + VLOOKUP + עזר", page_icon="✅", layout="centered")
st.markdown("""
<style>
html, body, [class*="css"] { direction: rtl; text-align: right; }
.block-container{padding-top:1rem}
</style>
""", unsafe_allow_html=True)
st.title("התאמות לקוחות/בנק – 1+2+3+4 (A2) + „הוראת קבע ספקים” + VLOOKUP + קובץ עזר (XLSX)")

# ---------- קבועים ----------
OVRC_CODES       = {120, 175}     # כלל 1 – OV/RC
STANDING_CODES   = {469, 515}     # כלל 2 – הוראות קבע
CHECK_CODE       = 493            # כלל 4 – שיקים ספקים
AMOUNT_TOL_4     = 0.20
RULES_PATH       = "rules_store.json"

MATCH_COL_CANDS  = ["מס. התאמה","מס התאמה","מספר התאמה","מס' התאמה","התאמה"]
BANK_CODE_CANDS  = ["קוד פעולת בנק","קוד פעולה","קוד פעולת","קוד בנק"]
BANK_AMT_CANDS   = ["סכום בדף","סכום בבנק","סכום בנק","סכום תנועת בנק"]
BOOKS_AMT_CANDS  = ["סכום בספרים","סכום ספרים","סכום בספר"]
REF1_CANDS       = ["אסמכתא 1","אסמכתא","אסמכתה","Reference 1"]
REF2_CANDS       = ["אסמכתא 2","אסמכתא2","אסמכתא נוספת","Reference 2"]
DATE_CANDS       = ["תאריך מאזן","תאריך","ת. מאזן","תאריך מסמך","תאריך ערך","ת. ערך"]
DETAILS_CANDS    = ["פרטים","תיאור","שם ספק","Details"]

HEADER_BLUE = PatternFill(start_color="9BC2E6", end_color="9BC2E6", fill_type="solid")
THIN = Side(style='thin', color='CCCCCC')
BORDER = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

# ---------- עזר כללי ----------
def _s(x):  return "" if x is None else str(x)
def _digits(x): 
    s = re.sub(r"\D","",_s(x)).lstrip("0"); return s or "0"
def _num(x):
    if x is None or (isinstance(x,float) and np.isnan(x)): return np.nan
    m = re.findall(r"[-+]?\d+(?:\.\d+)?", _s(x).replace(",","").replace("₪",""))
    return float(m[0]) if m else np.nan

def _find(df, cands):
    cols = list(map(str, df.columns))
    for c in cands:
        if c in cols: return c
    for c in cands:
        for col in cols:
            if c in col: return col
    return None

def ws_to_df(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows: return pd.DataFrame()
    header = [ _s(x) for x in rows[0] ]
    return pd.DataFrame(rows[1:], columns=header)

def style_sheet(ws):
    try:
        ws.sheet_view.rightToLeft = True
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 0
        ws.page_margins = PageMargins(left=0.3,right=0.3,top=0.5,bottom=0.5,header=0.3,footer=0.3)
        if ws.max_row>=1:
            for c in ws[1]:
                c.font=Font(bold=True); c.fill=HEADER_BLUE
                c.alignment=Alignment(horizontal='center',vertical='center'); c.border=BORDER
        for row in ws.iter_rows(min_row=2,max_row=ws.max_row,min_col=1,max_col=ws.max_column):
            for cell in row:
                cell.alignment=Alignment(horizontal='right',vertical='center'); cell.border=BORDER
        for j in range(1, ws.max_column+1):
            w=0
            for r in range(1,ws.max_row+1):
                v=ws.cell(row=r,column=j).value
                if v is not None: w=max(w,len(str(v)))
            ws.column_dimensions[get_column_letter(j)].width=min(w+2,60)
    except Exception: pass

# ---------- כללי VLOOKUP בזיכרון ----------
if "name_map" not in st.session_state:   st.session_state.name_map = {}
if "amount_map" not in st.session_state: st.session_state.amount_map = {}

def load_rules_from_disk():
    if os.path.exists(RULES_PATH):
        try:
            data = json.load(open(RULES_PATH,"r",encoding="utf-8"))
            st.session_state.name_map.update(data.get("name_map", {}))
            st.session_state.amount_map.update(data.get("amount_map", {}))
        except Exception:
            pass
load_rules_from_disk()

def save_rules_to_disk():
    with open(RULES_PATH,"w",encoding="utf-8") as f:
        json.dump({"name_map":st.session_state.name_map,
                   "amount_map":st.session_state.amount_map}, f, ensure_ascii=False, indent=2)

def map_supplier_by_rules(name_text: str, amount_val) -> str:
    """מאתר מס' ספק לפי שם/פרטים או לפי סכום מוחלט (עם גיבוי 'עגול')."""
    sname = _s(name_text).strip()
    if sname in st.session_state.name_map:
        return st.session_state.name_map[sname]
    for k,v in st.session_state.name_map.items():
        if k and k in sname:
            return v
    if amount_val is not None and not pd.isna(amount_val):
        key = str(round(abs(float(amount_val)), 2))
        if key in st.session_state.amount_map:
            return st.session_state.amount_map[key]
        key2 = str(int(round(abs(float(amount_val)))))
        return st.session_state.amount_map.get(key2, "")
    return ""

# ---------- ייבוא כללים מקובץ עזר (XLSX) ----------
def import_rules_from_excel_bytes(xls_bytes: bytes) -> int:
    """
    קורא קובץ עזר XLSX ומוסיף כללים:
      גיליון 'שם→ספק'   עם עמודות:  'שם/פרטים' | 'מס' ספק'
      גיליון 'סכום→ספק' עם עמודות:  'סכום'     | 'מס' ספק'
    מחזיר כמה כללים הוטמעו.
    """
    try:
        buf = io.BytesIO(xls_bytes)
        x = pd.ExcelFile(buf)
        added = 0

        if "שם→ספק" in x.sheet_names:
            df_n = pd.read_excel(x, "שם→ספק").fillna("")
            name_col = next((c for c in df_n.columns if "שם" in str(c) or "פרטים" in str(c)), None)
            supp_col = next((c for c in df_n.columns if "ספק" in str(c)), None)
            if name_col and supp_col:
                for _, r in df_n.iterrows():
                    k = str(r[name_col]).strip()
                    v = str(r[supp_col]).strip()
                    if k and v:
                        st.session_state.name_map[k] = v
                        added += 1

        if "סכום→ספק" in x.sheet_names:
            df_a = pd.read_excel(x, "סכום→ספק").fillna("")
            amt_col = next((c for c in df_a.columns if "סכום" in str(c)), None)
            supp_col = next((c for c in df_a.columns if "ספק" in str(c)), None)
            if amt_col and supp_col:
                for _, r in df_a.iterrows():
                    raw = str(r[amt_col]).replace(",", "").strip()
                    try:
                        key = str(round(abs(float(raw)), 2))
                    except Exception:
                        key = raw
                    v = str(r[supp_col]).strip()
                    if key and v:
                        st.session_state.amount_map[key] = v
                        added += 1

        save_rules_to_disk()
        return added
    except Exception:
        return 0

# ---------- כללי התאמה ----------
def rule_1(df, cols, match):
    stats={"pairs":0}
    bamt,samt,ref,date,code = cols["bank_amt"],cols["books_amt"],cols["ref"],cols["date"],cols["bank_code"]
    if not all([bamt,samt,ref,date,code]): return match,stats
    d = pd.to_datetime(df[date], errors="coerce").dt.date
    ba = df[bamt].apply(_num); sa=df[samt].apply(_num)
    r  = df[ref].astype(str).fillna(""); c=df[code].apply(_num)

    books = [i for i in df.index if sa.iat[i]>0 and pd.notna(d.iat[i]) and (r.iat[i].upper().startswith(("OV","RC")))]
    bank_keys,books_keys={},{}
    for i in df.index:
        if pd.notna(c.iat[i]) and int(c.iat[i]) in OVRC_CODES and ba.iat[i]<0 and pd.notna(d.iat[i]):
            key=(abs(round(ba.iat[i],2)), d.iat[i]); bank_keys.setdefault(key,[]).append(i)
    for j in books:
        key=(abs(round(sa.iat[j],2)), d.iat[j]); books_keys.setdefault(key,[]).append(j)
    for key, bi in bank_keys.items():
        bj = books_keys.get(key,[])
        if len(bi)==1 and len(bj)==1:
            i,j=bi[0],bj[0]
            if float(match.iat[i]) in (0,2,4) and float(match.iat[j]) in (0,2,4):
                match.iat[i]=1; match.iat[j]=1; stats["pairs"]+=1
    return match,stats

def rule_2_collect(df, cols, match, sheet_name):
    stats={"flagged":0}; rows=[]
    code,bamt,det = cols["bank_code"],cols["bank_amt"],cols["details"]
    if not all([code,bamt,det]): return match,stats,pd.DataFrame()
    c=df[code].apply(_num); ba=df[bamt].apply(_num); de=df[det].astype(str).fillna("")
    for i in df.index:
        if pd.notna(c.iat[i]) and int(c.iat[i]) in STANDING_CODES:
            if float(match.iat[i]) in (0,3,4):
                match.iat[i]=2; stats["flagged"]+=1
            rows.append({"גיליון מקור":sheet_name, "קוד פעולת בנק (מקורי)":df[code].iat[i],
                         "פרטים":de.iat[i], "סכום":float(ba.iat[i]) if pd.notna(ba.iat[i]) else np.nan})
    return match,stats,(pd.DataFrame(rows) if rows else pd.DataFrame())

def rule_3(df, cols, match):
    stats={"pairs":0}
    bamt,samt,date,ref,code = cols["bank_amt"],cols["books_amt"],cols["date"],cols["ref"],cols["bank_code"]
    if not all([bamt,samt,date]): return match,stats
    d=pd.to_datetime(df[date], errors="coerce").dt.date
    ba=df[bamt].apply(_num); sa=df[samt].apply(_num)
    r =df[ref].astype(str).fillna("") if ref else pd.Series([""]*len(df))
    c =df[code].apply(_num) if code else pd.Series([np.nan]*len(df))

    books=[i for i in df.index if sa.iat[i]>0 and pd.notna(d.iat[i]) and not r.iat[i].upper().startswith(("OV","RC"))]
    banks=[i for i in df.index if ba.iat[i]<0 and pd.notna(d.iat[i]) and (pd.isna(c.iat[i]) or (int(c.iat[i]) not in STANDING_CODES and int(c.iat[i])!=CHECK_CODE))]
    bank_keys,books_keys={},{}
    for i in banks: bank_keys.setdefault((abs(round(ba.iat[i],2)), d.iat[i]), []).append(i)
    for j in books: books_keys.setdefault((abs(round(sa.iat[j],2)), d.iat[j]), []).append(j)
    for key, bi in bank_keys.items():
        bj=books_keys.get(key,[])
        if len(bi)==1 and len(bj)==1:
            i,j=bi[0],bj[0]
            if float(match.iat[i]) in (0,2,4) and float(match.iat[j]) in (0,2,4):
                match.iat[i]=3; match.iat[j]=3; stats["pairs"]+=1
    return match,stats

def rule_4(df, cols, match):
    """שיקים ספקים – ללא בדיקת תאריכים, A2, נרמול אסמכתאות, טולרנס ₪0.20."""
    stats={"pairs":0}
    code,bamt,samt,ref1,ref2,det = cols["bank_code"],cols["bank_amt"],cols["books_amt"],cols["ref"],cols["ref2"],cols["details"]
    if not all([code,bamt,samt,ref1,ref2,det]): return match,stats
    c=df[code].apply(_num); ba=df[bamt].apply(_num); sa=df[samt].apply(_num)
    de=df[det].astype(str).fillna(""); r1=df[ref1].astype(str).fillna(""); r2=df[ref2].astype(str).fillna("")
    r1d=r1.apply(_digits); r2d=r2.apply(_digits)
    r2fb=r1.apply(lambda s:_digits(_s(s).replace("CH","").replace("ch","")))
    rid = r2d.where(r2d.ne("0"), r2fb)

    bank_idx=[i for i in df.index if float(match.iat[i]) in (0,2,3)
              and pd.notna(c.iat[i]) and int(c.iat[i])==CHECK_CODE and ("שיק" in de.iat[i]) and ba.iat[i]>0]
    books_idx=[j for j in df.index if float(match.iat[j]) in (0,2,3)
               and _s(r1.iat[j]).upper().startswith("CH") and (("תשלום בהמחאה" in de.iat[j]) or ("המחאה" in de.iat[j])) and sa.iat[j]<0]
    books_by_id={}
    for j in books_idx: books_by_id.setdefault(rid.iat[j], []).append(j)
    used=set()
    for i in bank_idx:
        key=r1d.iat[i]
        cands=[j for j in books_by_id.get(key,[]) if j not in used]
        choose=None
        for j in cands:
            if abs(float(ba.iat[i])+float(sa.iat[j]))<=AMOUNT_TOL_4:
                choose=j; break
        if choose is not None:
            match.iat[i]=4; match.iat[choose]=4; used.add(choose); stats["pairs"]+=1
    return match,stats

# ---------- עיבוד חוברת ----------
def process_workbook(xls_bytes: bytes):
    wb = load_workbook(io.BytesIO(xls_bytes))
    standing_all = []
    summary = []

    for ws in wb.worksheets:
        df = ws_to_df(ws)
        if df.empty: 
            style_sheet(ws); 
            summary.append({"גיליון":ws.title,"זוגות 1":0,"מסומני 2":0,"זוגות 3":0,"זוגות 4":0,"עמודת התאמה":"—"})
            continue

        match_col = _find(df, MATCH_COL_CANDS) or "מס. התאמה"
        if match_col not in df.columns: df[match_col]=0
        match = pd.to_numeric(df[match_col], errors="coerce").fillna(0)

        cols = {
            "bank_code": _find(df,BANK_CODE_CANDS),
            "bank_amt":  _find(df,BANK_AMT_CANDS),
            "books_amt": _find(df,BOOKS_AMT_CANDS),
            "ref":       _find(df,REF1_CANDS),
            "ref2":      _find(df,REF2_CANDS),
            "date":      _find(df,DATE_CANDS),
            "details":   _find(df,DETAILS_CANDS),
        }

        match, s1 = rule_1(df, cols, match)
        match, s2, vk = rule_2_collect(df, cols, match, ws.title)
        if not vk.empty: standing_all.append(vk)
        match, s3 = rule_3(df, cols, match)
        match, s4 = rule_4(df, cols, match)

        df_out = df.copy(); df_out[match_col]=match
        # מחיקה וכתיבה מחדש (שומר שם/סדר)
        for _ in range(ws.max_row-1): ws.delete_rows(2,1)
        for j,col in enumerate(df_out.columns, start=1): ws.cell(row=1,column=j).value = col
        for r in df_out.itertuples(index=False): ws.append(list(r))
        style_sheet(ws)

        summary.append({"גיליון":ws.title,
                        "זוגות 1":s1["pairs"], "מסומני 2":s2["flagged"],
                        "זוגות 3":s3["pairs"], "זוגות 4":s4["pairs"],
                        "עמודת התאמה":match_col})

    # גיליון „הוראת קבע ספקים”
    if standing_all:
        vk = pd.concat(standing_all, ignore_index=True)
        vk["מס' ספק"] = vk.apply(lambda r: map_supplier_by_rules(r.get("פרטים",""), r.get("סכום", np.nan)), axis=1)
        # עמודות עזר סכום חובה/זכות (לא חובה, אך שימושי)
        if "סכום" in vk.columns:
            vk["סכום חובה"] = vk["סכום"].apply(lambda x: abs(x) if pd.notna(x) and x<0 else 0.0)
            vk["סכום זכות"] = vk["סכום"].apply(lambda x: abs(x) if pd.notna(x) and x>0 else 0.0)
    else:
        vk = pd.DataFrame([{"גיליון מקור":"—","קוד פעולת בנק (מקורי)":"—","פרטים":"—","סכום":0.0,"מס' ספק":""}])

    if "הוראת קבע ספקים" in [s.title for s in wb.worksheets]:
        ws_vk = wb["הוראת קבע ספקים"]
    else:
        ws_vk = wb.create_sheet("הוראת קבע ספקים")
    for _ in range(ws_vk.max_row): ws_vk.delete_rows(1,1)
    for j,col in enumerate(vk.columns, start=1): ws_vk.cell(row=1,column=j).value = col
    for r in vk.itertuples(index=False): ws_vk.append(list(r))
    style_sheet(ws_vk)

    out = io.BytesIO(); wb.save(out)
    return out.getvalue(), pd.DataFrame(summary)

# ---------- UI – כללי VLOOKUP + עזר ----------
st.subheader("⚙️ כללי VLOOKUP (שם/פרטים ↔ מס' ספק, סכום ↔ מס' ספק)")
mode = st.radio("סוג עדכון", ["לפי פרטים (שם)", "לפי סכום"], horizontal=True)
ca, cb, cc = st.columns([2,1,1])

if mode == "לפי פרטים (שם)":
    key_name = ca.text_input("שם/פרטים (כמו שמופיע בעמודת 'פרטים')")
    val_supp = cb.text_input("מס' ספק")
    if cc.button("➕ הוסף/עדכן"):
        if key_name and val_supp:
            st.session_state.name_map[_s(key_name).strip()] = _s(val_supp).strip()
            save_rules_to_disk(); st.success("נשמר.")
else:
    key_amt = ca.text_input("סכום (יוזן כמספר או טקסט; נשמר כמחרוזת מעוגלת)")
    val_supp2 = cb.text_input("מס' ספק")
    if cc.button("➕ הוסף/עדכן"):
        if key_amt and val_supp2:
            try:
                k = str(round(abs(float(key_amt)),2))
            except Exception:
                k = _s(key_amt).strip()
            st.session_state.amount_map[k] = _s(val_supp2).strip()
            save_rules_to_disk(); st.success("נשמר.")

st.markdown("**(אופציונלי)** ייבוא כללים מקובץ עזר (XLSX) עם גיליונות 'שם→ספק' ו-'סכום→ספק'")
aux_rules_file = st.file_uploader("קובץ עזר לכללים (xlsx)", type=["xlsx"], key="aux_rules_file")
if st.button("📥 ייבוא כללים מקובץ עזר (xlsx)"):
    if aux_rules_file is None:
        st.warning("לא נבחר קובץ.")
    else:
        n = import_rules_from_excel_bytes(aux_rules_file.read())
        if n > 0: st.success(f"ייובאו {n} כללים מקובץ העזר ונשמרו.")
        else:     st.error("לא נמצאו גיליונות/עמודות מתאימים בקובץ העזר.")

c1,c2 = st.columns(2)
c1.download_button("⬇️ ייצוא JSON",
    data=json.dumps({"name_map":st.session_state.name_map,"amount_map":st.session_state.amount_map},
                    ensure_ascii=False, indent=2).encode("utf-8"),
    file_name="rules_store.json", mime="application/json")
if c2.button("איפוס כללים"):
    st.session_state.name_map = {}; st.session_state.amount_map = {}; save_rules_to_disk(); st.warning("נוקה.")

st.dataframe(pd.DataFrame({"שם/פרטים": list(st.session_state.name_map.keys()),
                           "מס' ספק": list(st.session_state.name_map.values())}),
             use_container_width=True, height=220)
st.dataframe(pd.DataFrame({"סכום": list(st.session_state.amount_map.keys()),
                           "מס' ספק": list(st.session_state.amount_map.values())}),
             use_container_width=True, height=220)

st.divider()

# ---------- UI – הרצה ----------
main = st.file_uploader("בחרי קובץ אקסל מקור (xlsx)", type=["xlsx"])
aux_rules_for_run = st.file_uploader("(אופציונלי) קובץ עזר לכללי VLOOKUP (xlsx) – טעינה אוטומטית לפני הריצה", type=["xlsx"], key="aux_rules_for_run")

if st.button("▶️ הרצה – 1+2+3+4 + גיליון הוראת קבע + VLOOKUP") and main is not None:
    # ייבוא כללים אוטומטי אם הועלה קובץ עזר
    if aux_rules_for_run is not None:
        n = import_rules_from_excel_bytes(aux_rules_for_run.read())
        if n > 0:
            st.info(f"ייובאו {n} כללים מקובץ העזר לפני העיבוד.")

    with st.spinner("מריץ את כל הכללים..."):
        out_bytes, summary_df = process_workbook(main.read())
    st.success("מוכן! אפשר להוריד.")
    st.dataframe(summary_df, use_container_width=True)
    st.download_button("⬇️ תוצאה סופית - 1+2+3+4 + הוראת קבע.xlsx",
        data=out_bytes,
        file_name="תוצאה סופית - 1+2+3+4 + הוראת קבע.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("העלאי קובץ מקור, ואם תרצי — הוסיפי קובץ עזר לכללים, ואז הרצה.")
