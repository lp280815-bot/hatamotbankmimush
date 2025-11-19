# -*- coding: utf-8 -*-
"""
התאמות בנק – 1 עד 12 + התאמות ספקים (גיול חובות)
- כלל 1: OV/RC 1:1 (תאריך+סכום)
- כלל 2: הוראות קבע (469/515) + 'הוראת קבע ספקים': כל השורות בחובה; שורת סיכום 20001 בזכות = סה״כ חובה של שורות עם מס' ספק. שורות בלי מס' ספק צבועות כתום.
- כלל 3: העברות (485, 'העב' במקבץ-נט') – מסמן רק אם קיים צד בנק וגם צד ספרים ושוויי־סכום (במונחי ערך מוחלט). אין דרישת התאמת תאריך. אם אין התאמה → גיליון 'פערי סכומים – כלל 3'.
- כלל 4: שיקים ספקים (493) עם טולרנס סכום על התאמת אסמכתאות (Ref1 בנק ↔ Ref2 ספרים).
- כללים 5–10: לפי הלוגיקה שאישרת.
- 11–12: placeholders.
- עיצוב: RTL, A4 לרוחב, Fit-to-width=1, שוליים נוחים.

התאמות ספקים (גיול חובות):
- כלל ראשון: 100% התאמה - חוב מצטבר בין -2 ל-2 ש"ח
- כלל שני: 80% התאמה - יש יתרה 0 בין השורות אבל חוב מצטבר סופי > 2
- כלל שלישי: שליחת מייל לספקים עם העברות חסרות חשבונית
"""

import io, os, re, json, smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font
from openpyxl.worksheet.page import PageMargins

# Import database module
import database as db

# Initialize database on app start
@st.cache_resource
def init_app_database():
    """Initialize database once per app session"""
    db.init_database()
    # Migrate from JSON if exists
    db.migrate_from_json()
    return True

init_app_database()

# ---------------- UI ----------------
st.set_page_config(page_title="התאמות בנק – 1 עד 12", page_icon="✅", layout="centered")
st.markdown("""
<style>
html, body, [class*="css"] { direction: rtl; text-align: right; }
.block-container { padding-top: 1rem; max-width: 1100px; }
</style>
""", unsafe_allow_html=True)
st.title("התאמות בנק – 1 עד 12")

# ---------------- Constants ----------------
STANDING_CODES = {469, 515}         # כלל 2
OVRC_CODES     = {120, 175}         # כלל 1
TRANSFER_CODE  = 485                # כלל 3
TRANSFER_PHRASE = "העב' במקבץ-נט"
RULE4_CODE     = 493                # כלל 4
RULE4_EPS      = 0.50               # טולרנס כלל 4

# כלל 3 – התאמת סכומים (0.00 = חייב זהות מוחלטת בערך מוחלט)
RULE3_AMOUNT_EPS = 0.00

# כללים 5–10
RULE5_CODES = {453, 472, 473, 124}  # עמלות – חיובי ועד 1000
RULE6_COMPANY = 'פאיימי בע"מ'       # קוד 175, שלילי, פרטים בדיוק
RULE7_CODE = 143; RULE7_PHRASE = "שיקים ממשמרת"
RULE8_CODE = 191; RULE8_PHRASE = "הפק' שיק-שידור"
RULE9_CODE = 205; RULE9_PHRASE = "הפק.שיק במכונה"
RULE10_CODES = {191, 132, 396}

# placeholders 11–12
def rule11_placeholder(df, match_col, code_col, bamt_col, details_col): return df[match_col]
def rule12_placeholder(df, match_col, code_col, bamt_col, details_col): return df[match_col]

# ---------------- Column maps ----------------
MATCH_COLS = ["מס.התאמה","מס. התאמה","מס התאמה","מספר התאמה","התאמה"]
BANK_CODES = ["קוד פעולת בנק","קוד פעולה","קוד פעולת","Bank Code"]
BANK_AMTS  = ["סכום בדף","סכום דף","סכום בבנק","סכום תנועת בנק","Bank Amount"]
BOOKS_AMTS = ["סכום בספרים","סכום בספר","סכום ספרים","Books Amount"]
REF1S      = ["אסמכתא 1","אסמכתא1","אסמכתא","אסמכתה","Ref1"]
REF2S      = ["אסמכתא 2","אסמכתא2","אסמכתא-2","אסמכתה 2","Ref2"]
DATES      = ["תאריך מאזן","תאריך ערך","תאריך","Date"]
DETAILS    = ["פרטים","תיאור","שם ספק","Details","תאור"]

# aux (עזר להעברות/כלל 3)
AUX_DATE_KEYS = ["תאריך פריקה","תאריך","פריקה"]   # כולל שעה/חותמת אירוע
AUX_AMT_KEYS  = ["אחרי ניכוי","אחרי","סכום"]
AUX_PAYNO_KEYS= ["מס' תשלום","מס תשלום","מספר תשלום"]

# ---------------- Helpers ----------------
def pick_col(df, names):
    for n in names:
        if n in df.columns: return n
    for n in names:
        for c in df.columns:
            if isinstance(c,str) and n in c: return c
    return None

def to_num(s):
    s = (s.astype(str)
         .str.replace(",","",regex=False)
         .str.replace("₪","",regex=False)
         .str.replace("\u200f","",regex=False)
         .str.replace("\u200e","",regex=False)
         .str.strip())
    return pd.to_numeric(s, errors="coerce")

def norm_date(series):
    def f(x):
        if pd.isna(x): return pd.NaT
        if isinstance(x,(pd.Timestamp,datetime)): return pd.Timestamp(x.date())
        return pd.to_datetime(x, dayfirst=True, errors="coerce").normalize()
    return series.apply(f)

def ws_to_df(ws):
    rows=list(ws.iter_rows(values_only=True))
    if not rows: return pd.DataFrame()
    header=[str(x) if x is not None else "" for x in rows[0]]
    data=[list(r[:len(header)]) for r in rows[1:]]
    return pd.DataFrame(data, columns=header)

def only_digits(s): return re.sub(r"\D","", str(s)).lstrip("0") or "0"

# ---------------- התאמות ספקים (גיול חובות) ----------------
def parse_supplier_aging(df: pd.DataFrame):
    """
    מנתח קובץ גיול חובות ומחלק לספקים בודדים.
    מזהה שורות של "חשבון: XXXX, תאור חשבון: שם ספק" ו"סה"כ לחשבון"

    Returns:
        list of dict: כל ספק עם הנתונים שלו
    """
    suppliers = []
    current_supplier = None
    current_rows = []

    # מציאת עמודות
    col_total_debt = None
    for col in df.columns:
        if 'חוב מצטבר' in str(col):
            col_total_debt = col
            break

    if col_total_debt is None:
        return []

    # מעבר על השורות
    for idx, row in df.iterrows():
        first_col = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""

        # זיהוי תחילת ספק
        if first_col.startswith("חשבון:"):
            # שמירת ספק קודם אם קיים
            if current_supplier is not None:
                suppliers.append({
                    'account_num': current_supplier['account_num'],
                    'account_name': current_supplier['account_name'],
                    'rows': current_rows.copy(),
                    'total_debt': current_supplier['total_debt']
                })

            # התחלת ספק חדש
            parts = first_col.split(',')
            account_num = parts[0].replace('חשבון:', '').strip()
            account_name = parts[1].replace('תאור חשבון:', '').strip() if len(parts) > 1 else ""

            current_supplier = {
                'account_num': account_num,
                'account_name': account_name,
                'total_debt': 0
            }
            current_rows = []

        # זיהוי סיום ספק
        elif first_col.startswith('סה"כ לחשבון:') or first_col.startswith('סה״כ לחשבון:'):
            if current_supplier is not None:
                # שליפת חוב מצטבר סופי
                total_debt_val = row[col_total_debt]
                if pd.notna(total_debt_val):
                    try:
                        total_debt_val = float(str(total_debt_val).replace(',', '').replace('₪', ''))
                    except:
                        total_debt_val = 0
                else:
                    total_debt_val = 0

                current_supplier['total_debt'] = total_debt_val
                suppliers.append({
                    'account_num': current_supplier['account_num'],
                    'account_name': current_supplier['account_name'],
                    'rows': current_rows.copy(),
                    'total_debt': total_debt_val
                })
                current_supplier = None
                current_rows = []

        # שורת תנועה רגילה
        elif current_supplier is not None and not first_col.startswith("תאריך"):
            # בדיקה שיש תוכן בשורה
            if pd.notna(row.iloc[0]) or any(pd.notna(row.iloc[i]) for i in range(len(row))):
                current_rows.append(row.to_dict())

    return suppliers

def classify_suppliers(suppliers: list):
    """
    מסווג ספקים לפי הכללים:
    - כלל 1 (100%): חוב מצטבר בין -2 ל-2
    - כלל 2 (80%): חוב מצטבר > 2 אבל יש שורה עם חוב מצטבר = 0
    - כלל 3: יש העברות ולא נכנס לכלל 1 או 2
    """
    rule1_suppliers = []  # 100% התאמה
    rule2_suppliers = []  # 80% התאמה
    rule3_suppliers = []  # העברות חסרות

    for supplier in suppliers:
        total_debt = supplier['total_debt']
        rows = supplier['rows']

        # כלל 1: חוב מצטבר בין -2 ל-2
        if -2 <= total_debt <= 2:
            rule1_suppliers.append(supplier)
        else:
            # בדיקה אם יש שורה עם חוב מצטבר = 0 (או קרוב ל-0)
            has_zero_debt_row = False
            for row in rows:
                row_debt = row.get('חוב מצטבר', None)
                if pd.notna(row_debt):
                    try:
                        debt_val = float(str(row_debt).replace(',', '').replace('₪', ''))
                        if -2 <= debt_val <= 2:
                            has_zero_debt_row = True
                            break
                    except:
                        pass

            if has_zero_debt_row:
                # כלל 2: יש יתרה 0 בין השורות
                rule2_suppliers.append(supplier)
            else:
                # בדיקה אם יש העברות
                has_transfer = False
                for row in rows:
                    movement_type = str(row.get('סוג תנועה', '')).strip()
                    if movement_type == 'העב':
                        has_transfer = True
                        break

                if has_transfer:
                    rule3_suppliers.append(supplier)

    return rule1_suppliers, rule2_suppliers, rule3_suppliers

def create_supplier_reconciliation_sheets(suppliers_list, rule_name):
    """
    יוצר DataFrame עבור גיליון התאמה
    """
    data = []
    for supplier in suppliers_list:
        data.append({
            'מס\' ספק': supplier['account_num'],
            'שם ספק': supplier['account_name'],
            'חוב מצטבר': supplier['total_debt']
        })

    return pd.DataFrame(data)

def identify_missing_transfers(suppliers: list, df_original: pd.DataFrame):
    """
    מזהה העברות שחסרות להן חשבוניות ומכין טיוטת מייל
    """
    missing_transfers = []

    for supplier in suppliers:
        for row in supplier['rows']:
            movement_type = str(row.get('סוג תנועה', '')).strip()
            if movement_type == 'העב':
                transfer_date = row.get('תאריך תשלום', '')
                transfer_amount = row.get('חוב לחשבונית', 0)

                try:
                    amount_val = float(str(transfer_amount).replace(',', '').replace('₪', ''))
                    amount_val = abs(amount_val)  # סכום בערך מוחלט
                except:
                    amount_val = 0

                missing_transfers.append({
                    'מס\' ספק': supplier['account_num'],
                    'שם ספק': supplier['account_name'],
                    'תאריך העברה': transfer_date,
                    'סכום העברה': amount_val,
                    'טיוטת מייל': f"""שלום,
אני מנהלת חשבונות של [שם לקוח].
חסרה לי חשבונית עבור העברה מתאריך {transfer_date} בסכום {amount_val:.2f} ₪.
אשמח לקבל חשבונית בהקדם.
תודה"""
                })

    return pd.DataFrame(missing_transfers)

def load_supplier_emails(email_file_bytes):
    """
    טוען קובץ עזר עם מיילים של ספקים
    מצפה למבנה: מס' ספק | שם ספק | מייל ספק
    """
    try:
        wb = load_workbook(io.BytesIO(email_file_bytes), data_only=True)
        ws = wb.worksheets[0]
        df = ws_to_df(ws)

        # מציאת עמודות
        email_map = {}
        for idx, row in df.iterrows():
            supplier_num = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            supplier_email = ""

            # חיפוש עמודת מייל
            for col_val in row:
                if pd.notna(col_val) and '@' in str(col_val):
                    supplier_email = str(col_val).strip()
                    break

            if supplier_num and supplier_email:
                email_map[supplier_num] = supplier_email

        return email_map
    except Exception as e:
        st.error(f"שגיאה בטעינת קובץ מיילים: {str(e)}")
        return {}

def send_email_smtp(smtp_server, smtp_port, sender_email, sender_password, recipient_email, subject, body, supplier_account=None):
    """
    שולח מייל דרך SMTP ומתעד במסד נתונים
    """
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'plain', 'utf-8'))

        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)

        # Log to database
        if supplier_account:
            db.log_email(supplier_account, recipient_email, subject, body, "success")

        return True, "נשלח בהצלחה"
    except Exception as e:
        # Log failure to database
        if supplier_account:
            db.log_email(supplier_account, recipient_email, subject, body, f"failed: {str(e)}")
        return False, str(e)

# ---------------- VLOOKUP store (using database) ----------------
def build_vlookup_sheet(datasheet_df: pd.DataFrame) -> pd.DataFrame:
    """
    כל שורה → 'סכום חובה' = |סכום|.
    שורת סיכום 20001 בזכות = סה״כ חובה של השורות שיש להן 'מס' ספק'.
    שורות בלי 'מס' ספק' – יצבעו בכתום בשלב העיצוב.
    """
    # Load mappings from database
    name_map = db.get_name_mappings()
    amount_map = db.get_amount_mappings()

    col_match = pick_col(datasheet_df, MATCH_COLS) or datasheet_df.columns[0]
    col_bamt  = pick_col(datasheet_df, BANK_AMTS)
    col_det   = pick_col(datasheet_df, DETAILS)

    match = pd.to_numeric(datasheet_df[col_match], errors="coerce").fillna(0).astype(int)
    bamt  = to_num(datasheet_df[col_bamt]) if col_bamt else pd.Series([np.nan]*len(datasheet_df))
    det   = datasheet_df[col_det].astype(str).fillna("") if col_det else pd.Series([""]*len(datasheet_df))

    # בדיקה שיש עמודות נדרשות
    if not col_det or not col_bamt:
        return pd.DataFrame(columns=["פרטים","סכום","מס' ספק","סכום חובה","סכום זכות"])

    vk = datasheet_df.loc[match==2, [col_det, col_bamt]].rename(columns={col_det:"פרטים", col_bamt:"סכום"}).copy()
    if vk.empty:
        return pd.DataFrame(columns=["פרטים","סכום","מס' ספק","סכום חובה","סכום זכות"])

    def pick_supplier(row):
        s = str(row["פרטים"])
        for k,v in name_map.items():
            if k and k in s: return v
        try:
            key = round(abs(float(row["סכום"])),2)
            return amount_map.get(key,"")
        except Exception:
            return ""

    vk["מס' ספק"]   = vk.apply(pick_supplier, axis=1)
    vk["סכום חובה"] = vk["סכום"].apply(lambda x: abs(x) if pd.notna(x) else 0.0)
    vk["סכום זכות"] = 0.0

    total_hova_with_supplier = vk.loc[vk["מס' ספק"].astype(str).str.len()>0, "סכום חובה"].sum()
    if total_hova_with_supplier and total_hova_with_supplier != 0:
        vk = pd.concat([vk, pd.DataFrame([{
            "פרטים": "סה\"כ זכות – עם מס' ספק",
            "סכום": 0.0,
            "מס' ספק": 20001,
            "סכום חובה": 0.0,
            "סכום זכות": round(float(total_hova_with_supplier), 2)
        }])], ignore_index=True)

    return vk

# ---------------- Rules 1–4 ----------------
def apply_rules_1_4(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    col_match = pick_col(out, MATCH_COLS) or out.columns[0]
    col_code  = pick_col(out, BANK_CODES)
    col_bamt  = pick_col(out, BANK_AMTS)
    col_aamt  = pick_col(out, BOOKS_AMTS)
    col_ref1  = pick_col(out, REF1S)
    col_ref2  = pick_col(out, REF2S)
    col_date  = pick_col(out, DATES)
    col_det   = pick_col(out, DETAILS)

    if col_match not in out.columns: out[col_match]=0

    match = pd.to_numeric(out[col_match], errors="coerce").fillna(0).astype(int)
    code  = to_num(out[col_code]) if col_code else pd.Series([np.nan]*len(out))
    bamt  = to_num(out[col_bamt]) if col_bamt else pd.Series([np.nan]*len(out))
    aamt  = to_num(out[col_aamt]) if col_aamt else pd.Series([np.nan]*len(out))
    datev = norm_date(pd.to_datetime(out[col_date], errors="coerce")) if col_date else pd.Series([pd.NaT]*len(out))
    det   = out[col_det].astype(str).fillna("") if col_det else pd.Series([""]*len(out))
    ref1  = out[col_ref1].astype(str).fillna("") if col_ref1 else pd.Series([""]*len(out))
    ref2  = out[col_ref2].astype(str).fillna("") if col_ref2 else pd.Series([""]*len(out))

    # 1: OV/RC 1:1
    bank_keys, books_keys = {}, {}
    for i in range(len(out)):
        if match.iat[i]!=0: continue
        if pd.notna(code.iat[i]) and int(code.iat[i]) in OVRC_CODES and pd.notna(bamt.iat[i]) and bamt.iat[i] < 0 and pd.notna(datev.iat[i]):
            k=(round(abs(float(bamt.iat[i])),2), datev.iat[i]); bank_keys.setdefault(k, []).append(i)
    for j in range(len(out)):
        if match.iat[j]!=0: continue
        if pd.notna(aamt.iat[j]) and aamt.iat[j]>0 and pd.notna(datev.iat[j]) and str(ref1.iat[j]).upper().startswith(("OV","RC")):
            k=(round(abs(float(aamt.iat[j])),2), datev.iat[j]); books_keys.setdefault(k, []).append(j)
    for k, bidx in bank_keys.items():
        if len(bidx)==1 and len(books_keys.get(k,[]))==1:
            i=bidx[0]; j=books_keys[k][0]
            if match.iat[i]==0 and match.iat[j]==0: match.iat[i]=1; match.iat[j]=1

    # 2: Standing orders (סימון בלבד)
    for i in range(len(out)):
        if match.iat[i]==0 and pd.notna(code.iat[i]) and int(code.iat[i]) in STANDING_CODES:
            match.iat[i]=2

    # 3: יסומן בשלב process_workbook (תלוי עזר; ללא בדיקת תאריך)
    # 4: שיקים ספקים (Ref1 בנק ↔ Ref2 ספרים) + טולרנס
    bank_idx = [i for i in range(len(out)) if match.iat[i]==0 and pd.notna(code.iat[i]) and int(code.iat[i])==RULE4_CODE and str(ref1.iat[i]).strip() and pd.notna(bamt.iat[i])]
    books_idx= [j for j in range(len(out)) if match.iat[j]==0 and str(ref1.iat[j]).upper().startswith("CH") and str(ref2.iat[j]).strip() and pd.notna(aamt.iat[j])]
    used=set()
    for i in bank_idx:
        ref_b = only_digits(ref1.iat[i]); ab = abs(float(bamt.iat[i]))
        for j in books_idx:
            if j in used or match.iat[j]!=0: continue
            if only_digits(ref2.iat[j]) != ref_b: continue
            aj = abs(float(aamt.iat[j]))
            if abs(aj - ab) <= RULE4_EPS:
                match.iat[i]=4; match.iat[j]=4; used.add(j); break

    out[col_match] = match
    return out

# ---------------- Rules 5–12 ----------------
def apply_rules_5_12(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    col_match = pick_col(out, MATCH_COLS) or out.columns[0]
    col_code  = pick_col(out, BANK_CODES)
    col_bamt  = pick_col(out, BANK_AMTS)
    col_det   = pick_col(out, DETAILS)

    if col_match not in out.columns: out[col_match]=0
    match = pd.to_numeric(out[col_match], errors="coerce").fillna(0).astype(int)
    code  = to_num(out[col_code]) if col_code else pd.Series([np.nan]*len(out))
    bamt  = to_num(out[col_bamt]) if col_bamt else pd.Series([np.nan]*len(out))
    det   = out[col_det].astype(str).fillna("") if col_det else pd.Series([""]*len(out))

    m5  = (match==0) & (code.isin(list(RULE5_CODES))) & (bamt>0) & (bamt<=1000); match.loc[m5]=5
    m6  = (match==0) & (code==175) & (bamt<0) & (det==RULE6_COMPANY); match.loc[m6]=6
    m7  = (match==0) & (code==RULE7_CODE) & (bamt<0) & (det==RULE7_PHRASE); match.loc[m7]=7
    m8  = (match==0) & (code==RULE8_CODE) & (bamt<0) & (det==RULE8_PHRASE); match.loc[m8]=8
    m9  = (match==0) & (code==RULE9_CODE) & (bamt<0) & (det==RULE9_PHRASE); match.loc[m9]=9
    m10 = (match==0) & (code.isin(list(RULE10_CODES))) & (bamt.notna()) & (bamt!=0); match.loc[m10]=10

    match = rule11_placeholder(out.assign(**{col_match:match}), col_match, pick_col(out,BANK_CODES), pick_col(out,BANK_AMTS), pick_col(out,DETAILS))
    match = rule12_placeholder(out.assign(**{col_match:match}), col_match, pick_col(out,BANK_CODES), pick_col(out,BANK_AMTS), pick_col(out,DETAILS))

    out[col_match]=match
    return out

# ---------------- Styling & print ----------------
def style_and_print(wb):
    for ws in wb.worksheets:
        ws.sheet_view.rightToLeft = True
        ws.page_setup.paperSize = 9      # A4
        ws.page_setup.orientation = 'landscape'
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_margins = PageMargins(left=0.4, right=0.4, top=0.6, bottom=0.6)

    # צביעה כתומה לשורות ללא 'מס' ספק' בגיליון הוראת קבע
    if "הוראת קבע ספקים" in wb.sheetnames:
        ws = wb["הוראת קבע ספקים"]
        header = {cell.value: idx for idx, cell in enumerate(ws[1], start=1)}
        col_supplier = header.get("מס' ספק")
        if col_supplier:
            orange = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            for r in range(2, ws.max_row+1):
                v = ws.cell(row=r, column=col_supplier).value
                if v in ("", None):
                    for c in range(1, ws.max_column+1):
                        ws.cell(row=r, column=c).fill = orange
        for cell in ws[ws.max_row]:
            cell.font = Font(bold=True)

# ---------------- Processing ----------------
def process_workbook(main_bytes: bytes, aux_bytes: bytes|None):
    # קריאה
    wb = load_workbook(io.BytesIO(main_bytes), data_only=True)
    ws = wb["DataSheet"] if "DataSheet" in wb.sheetnames else wb.worksheets[0]
    df = ws_to_df(ws)
    if df.empty: return None, None, None

    # 1–4
    df = apply_rules_1_4(df)

    # === כלל 3 – סכומים זהים בלבד (ללא דרישת תאריך) ===
    st.session_state["_rule3_mismatches_df"] = None
    if aux_bytes is not None:
        aux_wb = load_workbook(io.BytesIO(aux_bytes), data_only=True)
        a_ws   = aux_wb.worksheets[0]
        a_df   = ws_to_df(a_ws)

        c_dt   = pick_col(a_df, AUX_DATE_KEYS)   # תאריך/חותמת אירוע
        c_amt  = pick_col(a_df, AUX_AMT_KEYS)    # אחרי ניכוי
        c_pay  = pick_col(a_df, AUX_PAYNO_KEYS)  # מס' תשלום

        if c_dt and c_amt:
            a_dt  = pd.to_datetime(a_df[c_dt], errors="coerce")               # אירוע
            a_amt = pd.to_numeric(a_df[c_amt], errors="coerce").round(2)
            groups = (pd.DataFrame({"evt": a_dt, "amt": a_amt})
                        .dropna(subset=["evt"])
                        .groupby("evt")["amt"].sum().round(2).to_dict())

            pays_by_evt = {}
            if c_pay:
                pays_by_evt = (pd.DataFrame({"evt": a_dt, "pay": a_df[c_pay].astype(str).str.strip()})
                                 .groupby("evt")["pay"]
                                 .apply(lambda s: set(s.dropna().astype(str)))
                                 .to_dict())

            col_match = pick_col(df, MATCH_COLS) or df.columns[0]
            col_code  = pick_col(df, BANK_CODES)
            col_bamt  = pick_col(df, BANK_AMTS)
            col_det   = pick_col(df, DETAILS)
            col_ref1  = pick_col(df, REF1S)
            col_aamt  = pick_col(df, BOOKS_AMTS)

            match = pd.to_numeric(df[col_match], errors="coerce").fillna(0).astype(int)
            code  = to_num(df[col_code]) if col_code else pd.Series([np.nan]*len(df))
            bamt  = to_num(df[col_bamt]).round(2) if col_bamt else pd.Series([np.nan]*len(df))
            det   = df[col_det].astype(str).fillna("") if col_det else pd.Series([""]*len(df))
            ref1  = df[col_ref1].astype(str).str.strip() if col_ref1 else pd.Series([""]*len(df))
            aamt  = to_num(df[col_aamt]).round(2) if col_aamt else pd.Series([np.nan]*len(df))

            bank_mask = (match==0) & (code==TRANSFER_CODE) & (bamt>0) & (det.str.contains(TRANSFER_PHRASE, na=False))
            mismatches = []

            for evt, evt_sum in groups.items():
                # ספרים לפי payset של האירוע
                payset = pays_by_evt.get(evt, set())
                books_idx = []
                books_sum = 0.0
                if payset is not None and len(payset)>0 and col_ref1 and col_aamt:
                    books_idx = df.index[(match==0) & (ref1.astype(str).isin(payset))].tolist()
                    if books_idx:
                        books_sum = float(pd.to_numeric(aamt.iloc[books_idx], errors="coerce").fillna(0).sum().round(2))

                # בנק – כל השורות שסכומן |bamt| == |evt_sum|
                bank_idx = df.index[ bank_mask & (bamt.abs().sub(abs(evt_sum)).abs() <= RULE3_AMOUNT_EPS) ].tolist()

                if bank_idx and books_idx:
                    # התאמה חייבת להיות שוויון בערך מוחלט
                    if abs(abs(books_sum) - abs(evt_sum)) <= RULE3_AMOUNT_EPS:
                        for i in bank_idx:
                            if match.iat[i] in (0,2): match.iat[i]=3
                        for j in books_idx:
                            if match.iat[j] in (0,2): match.iat[j]=3
                    else:
                        mismatches.append({
                            "אירוע": str(evt),
                            "סכום בעזר (אחרי ניכוי)": float(evt_sum),
                            "סכום בספרים (סיכום)": float(books_sum),
                            "פער |ספרים|-|עזר|": float(round(abs(abs(books_sum)-abs(evt_sum)),2)),
                            "count_בנק": len(bank_idx),
                            "count_ספרים": len(books_idx)
                        })
                else:
                    mismatches.append({
                        "אירוע": str(evt),
                        "סכום בעזר (אחרי ניכוי)": float(evt_sum),
                        "סכום בספרים (סיכום)": float(books_sum) if books_idx else np.nan,
                        "פער |ספרים|-|עזר|": np.nan,
                        "count_בנק": len(bank_idx),
                        "count_ספרים": len(books_idx)
                    })

            df[col_match]=match
            if mismatches:
                st.session_state["_rule3_mismatches_df"] = pd.DataFrame(mismatches)

    # 5–12 (רק על 0)
    df = apply_rules_5_12(df)

    # גיליון הוראת קבע ספקים
    vk_df = build_vlookup_sheet(df)

    # יצוא עם עיצוב + גיליון בקרה לכלל 3 (אם יש)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as wr:
        df.to_excel(wr, index=False, sheet_name="DataSheet")
        counts = pd.to_numeric(df[pick_col(df, MATCH_COLS) or df.columns[0]],
                               errors="coerce").fillna(0).astype(int).value_counts().sort_index()
        pd.DataFrame({"מס":counts.index,"כמות":counts.values}).to_excel(wr, index=False, sheet_name="סיכום")
        vk_df.to_excel(wr, index=False, sheet_name="הוראת קבע ספקים")

        misdf = st.session_state.get("_rule3_mismatches_df", None)
        if misdf is not None and not misdf.empty:
            misdf.to_excel(wr, index=False, sheet_name="פערי סכומים – כלל 3")

    wb_out = load_workbook(io.BytesIO(buffer.getvalue()))
    style_and_print(wb_out)
    final = io.BytesIO(); wb_out.save(final)
    return df, vk_df, final.getvalue()

# ---------------- UI ----------------
c1, c2 = st.columns([2,2])
main_file = c1.file_uploader("בחרי קובץ מקור – DataSheet בלבד", type=["xlsx"])
aux_file  = c2.file_uploader("⬆️ קובץ עזר להעברות (לכלל 3)", type=["xlsx"],
                             help="קובץ Excel עם פרטי העברות מהבנק (תאריך, סכום, מס' תשלום) לצורך התאמת כלל 3")
st.caption("💡 קובץ עזר = קובץ ממערכת הבנק עם פירוט העברות (תאריך אירוע, אחרי ניכוי, מס' תשלום)")
st.caption("VLOOKUP שומר מפות במסד נתונים (שם/סכום → מס' ספק).")

if st.button("הרצה 1–12"):
    if not main_file:
        st.error("נא להעלות קובץ מקור.")
    else:
        with st.spinner("מעבד..."):
            df_out, vk_out, out_bytes = process_workbook(main_file.read(), aux_file.read() if aux_file else None)
        if df_out is None:
            st.error("לא נמצאו נתונים.")
        else:
            st.success("מוכן!")
            col_match = pick_col(df_out, MATCH_COLS) or df_out.columns[0]
            cnt = pd.to_numeric(df_out[col_match], errors="coerce").fillna(0).astype(int).value_counts().sort_index()
            st.dataframe(pd.DataFrame({"מס":cnt.index,"כמות":cnt.values}), use_container_width=True)
            st.download_button("📥 הורד קובץ מעודכן", data=out_bytes,
                               file_name="התאמות_1_עד_12.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ניהול מפות ל-VLOOKUP
st.divider()
st.subheader("🔎 VLOOKUP – הוראת קבע ספקים (עריכה ושמירה)")
with st.expander("מפות מיפוי (נשמר במסד נתונים)", expanded=False):
    t1, t2 = st.columns([2,1])
    nm = t1.text_input("מיפוי לפי 'פרטים' (contains)")
    sp = t2.text_input("מס' ספק")
    if st.button("➕ הוסף/עדכן לפי שם"):
        if nm and sp:
            db.save_name_mapping(nm, sp)
            st.success("נשמר לפי שם במסד הנתונים.")
    t3, t4 = st.columns([1,1])
    amt = t3.number_input("מיפוי לפי סכום (ערך מוחלט)", step=0.01, format="%.2f")
    sp2 = t4.text_input("מס' ספק", key="vk2")
    if st.button("➕ הוסף/עדכן לפי סכום"):
        try:
            db.save_amount_mapping(abs(float(amt)), sp2)
            st.success("נשמר לפי סכום במסד הנתונים.")
        except Exception as e:
            st.error(str(e))

# ---------------- התאמות ספקים (גיול חובות) ----------------
st.divider()
st.header("📊 התאמות ספקים - גיול חובות")

aging_file = st.file_uploader("העלאת קובץ גיול חובות ספקים", type=["xlsx"], key="aging")
client_name_input = st.text_input("שם הלקוח (לשליחת מיילים)", value="")

if st.button("🔍 נתח גיול חובות"):
    if not aging_file:
        st.error("נא להעלות קובץ גיול חובות")
    else:
        with st.spinner("מעבד גיול חובות..."):
            try:
                # קריאת קובץ
                aging_wb = load_workbook(io.BytesIO(aging_file.read()), data_only=True)
                aging_ws = aging_wb.worksheets[0]
                aging_df = ws_to_df(aging_ws)

                # ניתוח ספקים
                suppliers = parse_supplier_aging(aging_df)

                if not suppliers:
                    st.error("לא נמצאו ספקים בקובץ")
                else:
                    # סיווג ספקים
                    rule1, rule2, rule3 = classify_suppliers(suppliers)

                    # יצירת גיליונות
                    df_100 = create_supplier_reconciliation_sheets(rule1, "100% התאמה")
                    df_80 = create_supplier_reconciliation_sheets(rule2, "80% התאמה")
                    df_transfers = identify_missing_transfers(rule3, aging_df)

                    # שמירת שם לקוח במייל אם צוין
                    if client_name_input and not df_transfers.empty:
                        df_transfers['טיוטת מייל'] = df_transfers['טיוטת מייל'].str.replace(
                            '[שם לקוח]', client_name_input
                        )

                    # הצגת תוצאות
                    st.success(f"נותחו {len(suppliers)} ספקים")

                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("100% התאמה", len(rule1))
                    with col2:
                        st.metric("80% התאמה", len(rule2))
                    with col3:
                        st.metric("העברות חסרות", len(rule3))

                    # תצוגת טבלאות
                    if not df_100.empty:
                        st.subheader("✅ 100% התאמה")
                        st.dataframe(df_100, use_container_width=True)

                    if not df_80.empty:
                        st.subheader("⚠️ 80% התאמה")
                        st.dataframe(df_80, use_container_width=True)

                    if not df_transfers.empty:
                        st.subheader("📧 העברות חסרות חשבונית (טיוטות מייל)")
                        st.dataframe(df_transfers, use_container_width=True)

                    # יצוא לקובץ Excel
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
                        if not df_100.empty:
                            df_100.to_excel(writer, index=False, sheet_name="100% התאמה")
                        if not df_80.empty:
                            df_80.to_excel(writer, index=False, sheet_name="80% התאמה")
                        if not df_transfers.empty:
                            df_transfers.to_excel(writer, index=False, sheet_name="העברות חסרות")

                    wb_aging = load_workbook(io.BytesIO(buffer.getvalue()))
                    style_and_print(wb_aging)
                    final_aging = io.BytesIO()
                    wb_aging.save(final_aging)

                    st.download_button(
                        "📥 הורד דוח התאמות ספקים",
                        data=final_aging.getvalue(),
                        file_name="התאמות_ספקים.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                    # שמירת נתונים ב-session_state למייל
                    st.session_state['transfers_for_email'] = df_transfers

            except Exception as e:
                st.error(f"שגיאה בעיבוד: {str(e)}")
                import traceback
                st.text(traceback.format_exc())

# ---------------- שליחת מיילים אוטומטית ----------------
st.divider()
st.header("📧 שליחת מיילים אוטומטית לספקים")

# טעינת קובץ עזר מיילים
email_helper_file = st.file_uploader("העלאת קובץ עזר - מיילים של ספקים", type=["xlsx"], key="email_helper")

# טעינת מיילים אם הועלה קובץ
if email_helper_file:
    email_map = load_supplier_emails(email_helper_file.read())
    if email_map:
        st.success(f"נטענו מיילים של {len(email_map)} ספקים")
        st.session_state['supplier_emails'] = email_map

if 'transfers_for_email' in st.session_state and not st.session_state['transfers_for_email'].empty:
    df_emails = st.session_state['transfers_for_email']

    st.info(f"נמצאו {len(df_emails)} העברות שדורשות מייל")

    # הגדרות SMTP
    with st.expander("⚙️ הגדרות מייל (SMTP)", expanded=False):
        smtp_server = st.text_input("שרת SMTP", value="smtp.gmail.com")
        smtp_port = st.number_input("פורט SMTP", value=587)
        sender_email = st.text_input("כתובת מייל שולח")
        sender_password = st.text_input("סיסמה", type="password")
        email_subject = st.text_input("נושא המייל", value="בקשה לחשבונית - העברה")

        if st.button("📨 שלח מיילים לכל הספקים"):
            if not sender_email or not sender_password:
                st.error("נא למלא כתובת מייל וסיסמה")
            elif 'supplier_emails' not in st.session_state:
                st.error("נא להעלות קובץ עזר עם מיילי ספקים")
            else:
                email_map = st.session_state['supplier_emails']
                success_count = 0
                fail_count = 0

                progress_bar = st.progress(0)
                status_text = st.empty()

                for idx, row in df_emails.iterrows():
                    supplier_num = str(row['מס\' ספק'])
                    supplier_name = row['שם ספק']
                    email_body = row['טיוטת מייל']

                    # חיפוש מייל ספק
                    recipient = email_map.get(supplier_num)

                    if recipient:
                        status_text.text(f"שולח מייל ל-{supplier_name}...")
                        success, msg = send_email_smtp(
                            smtp_server, smtp_port,
                            sender_email, sender_password,
                            recipient, email_subject, email_body,
                            supplier_account=supplier_num
                        )

                        if success:
                            success_count += 1
                        else:
                            fail_count += 1
                            st.warning(f"נכשל לשלוח ל-{supplier_name}: {msg}")
                    else:
                        fail_count += 1
                        st.warning(f"לא נמצא מייל עבור ספק {supplier_num} - {supplier_name}")

                    progress_bar.progress((idx + 1) / len(df_emails))

                status_text.text("")
                st.success(f"✅ נשלחו {success_count} מיילים בהצלחה")
                if fail_count > 0:
                    st.error(f"❌ {fail_count} מיילים נכשלו")

        # אפשרות לשליחה ידנית
        st.subheader("שליחה ידנית למייל ספציפי")
        manual_supplier = st.selectbox(
            "בחר ספק",
            options=range(len(df_emails)),
            format_func=lambda x: "{} - {}".format(df_emails.iloc[x]['מס\' ספק'], df_emails.iloc[x]['שם ספק'])
        )

        if manual_supplier is not None:
            selected_row = df_emails.iloc[manual_supplier]
            manual_email = st.text_input("מייל נמען", value="")
            st.text_area("תוכן המייל", value=selected_row['טיוטת מייל'], height=200)

            if st.button("שלח מייל בודד"):
                if not manual_email:
                    st.error("נא להזין כתובת מייל")
                elif not sender_email or not sender_password:
                    st.error("נא להזין פרטי שולח")
                else:
                    success, msg = send_email_smtp(
                        smtp_server, smtp_port,
                        sender_email, sender_password,
                        manual_email, email_subject,
                        selected_row['טיוטת מייל'],
                        supplier_account=selected_row['מס\' ספק']
                    )
                    if success:
                        st.success("✅ המייל נשלח בהצלחה!")
                    else:
                        st.error(f"❌ שגיאה: {msg}")

else:
    st.info("יש להריץ תחילה ניתוח גיול חובות כדי לזהות העברות שדורשות מייל")
