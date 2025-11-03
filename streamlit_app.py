# ---------- helper: build mapping with explicit column names ----------
def build_amount_to_paynums_explicit(aux_df: pd.DataFrame,
                                     col_date: str, col_amount: str, col_paynum: str,
                                     col_time: str | None, ignore_time: bool):
    # ניקוי שדות
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
    return amount_to_paynums, sums.reset_index().rename(columns={"key":"קבוצה (תאריך+זמן)","amt":"סכום אחרי ניכוי"})

# ---------- Tab: התאמות ספקים – העברות (3) עם בחירת עמודות ----------
with st.expander("🏷️ התאמות ספקים – העברות (3)", expanded=True):
    c1, c2 = st.columns(2)
    main_xlsx = c1.file_uploader("קובץ מקור (Excel)", type=["xlsx"], key="m3x")
    aux_xlsx  = c2.file_uploader("קובץ עזר (Excel) — תאריך פריקה/זמן*/אחרי ניכוי/מס׳ תשלום", type=["xlsx"], key="a3x")

    # פרמטרים
    st.caption("פרמטרים")
    p1, p2, p3, p4 = st.columns([1,1,1,1.2])
    bank_code_val = p1.number_input("קוד פעולה", value=485, step=1)
    details_phrase = p2.text_input("ביטוי ב'פרטים'", value="העב' במקבץ-נט")
    amount_tol = p3.number_input("סבילות סכום (₪)", value=0.05, step=0.01, format="%.2f")
    ignore_time = p4.checkbox("להתעלם משדה הזמן בקובץ העזר (קיבוץ לפי תאריך בלבד)", value=False)

    if main_xlsx and aux_xlsx:
        # טוענים את שני הקבצים לגיליונות ראשונים ותופסים כותרות
        m_wb = load_workbook(main_xlsx, data_only=True, read_only=True)
        a_wb = load_workbook(aux_xlsx,  data_only=True, read_only=True)
        m_df = ws_to_df(m_wb.worksheets[0])
        a_df = ws_to_df(a_wb.worksheets[0])

        st.markdown("**בחירת עמודות – קובץ מקור**")
        mcols = list(m_df.columns)

        # הצעות אוטומטיות
        def first_match(cands, cols):
            for n in cands:
                if n in cols: return n
            for n in cands:
                for c in cols:
                    if isinstance(c,str) and n in c: return c
            return None

        m_match  = st.selectbox("עמודת 'מס. התאמה'", options=mcols, index=(mcols.index(first_match(MATCH_COL_CANDS, mcols)) if first_match(MATCH_COL_CANDS, mcols) in mcols else 0))
        m_code   = st.selectbox("עמודת 'קוד פעולת בנק'", options=mcols, index=(mcols.index(first_match(BANK_CODE_CANDS, mcols)) if first_match(BANK_CODE_CANDS, mcols) in mcols else 0))
        m_bamt   = st.selectbox("עמודת 'סכום בדף'", options=mcols, index=(mcols.index(first_match(BANK_AMT_CANDS, mcols)) if first_match(BANK_AMT_CANDS, mcols) in mcols else 0))
        m_ref    = st.selectbox("עמודת 'אסמכתא 1'", options=mcols, index=(mcols.index(first_match(REF_CANDS, mcols)) if first_match(REF_CANDS, mcols) in mcols else 0))
        m_det    = st.selectbox("עמודת 'פרטים'", options=mcols, index=(mcols.index(first_match(DETAILS_CANDS, mcols)) if first_match(DETAILS_CANDS, mcols) in mcols else 0))

        st.markdown("**בחירת עמודות – קובץ עזר**")
        acols = list(a_df.columns)
        a_date  = st.selectbox("תאריך פריקה", options=acols, index=(acols.index(first_match(AUX_DATE_CANDS, acols)) if first_match(AUX_DATE_CANDS, acols) in acols else 0))
        a_time  = st.selectbox("זמן (לא חובה)", options=["(ללא)"]+acols, index=0)
        a_time  = None if a_time == "(ללא)" else a_time
        a_amount = st.selectbox("אחרי ניכוי", options=acols, index=(acols.index(first_match(AUX_AMOUNT_CANDS, acols)) if first_match(AUX_AMOUNT_CANDS, acols) in acols else 0))
        a_pay    = st.selectbox("מס' תשלום", options=acols, index=(acols.index(first_match(AUX_PAYNUM_CANDS, acols)) if first_match(AUX_PAYNUM_CANDS, acols) in acols else 0))

        if st.button("סמן התאמות 3 והורד קובץ (עם דיאגנוסטיקה)"):
            with st.spinner("מסמן 3..."):
                # מיפוי סכומים -> מס' תשלום
                amount_to_paynums, aux_groups = build_amount_to_paynums_explicit(a_df, a_date, a_amount, a_pay, a_time, ignore_time)

                # הכנה לכתיבה
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine="xlsxwriter") as wr:
                    matches_log = []

                    # מעבד את אותו גיליון בלבד (ראשון) – כמו בקוד המקורי
                    df = m_df.copy()
                    # סדרות
                    s_match = df[m_match].copy()
                    s_code  = to_number(df[m_code])
                    s_bank  = to_number(df[m_bamt])
                    s_ref   = df[m_ref].astype(str)
                    s_det   = df[m_det].astype(str)

                    # פילטור: שורות בנק 485 + ביטוי + סכום חיובי
                    mask_bank = (s_code == float(bank_code_val)) & (s_bank > 0) & (s_det.str.contains(details_phrase, na=False))
                    bank_idx = list(df.index[mask_bank])

                    flagged = 0
                    for i in bank_idx:
                        amt = round(float(s_bank.iat[i]), 2)

                        # חיפוש עם סבילות
                        # מאתרים סכומים בקובץ עזר שקרובים בתוך הטולרנס
                        close_paynums = set()
                        for key_amt, pays in amount_to_paynums.items():
                            if abs(key_amt - amt) <= float(amount_tol):
                                close_paynums |= pays

                        if not close_paynums:
                            matches_log.append({"row": i+1, "סכום בנק": amt, "סט מס׳ תשלום": "—", "התאמה": "לא נמצאה בקובץ עזר"})
                            continue

                        # מסמן את הבנק (לא לדרוס 1/2)
                        if s_match.iat[i] not in (1,2):
                            s_match.iat[i] = 3
                            flagged += 1

                        # מסמן את הספרים לפי אסמכתא 1 (לא לדרוס 1/2)
                        mask_ref = s_ref.isin(close_paynums)
                        for j in df.index[mask_ref]:
                            if s_match.iat[j] not in (1,2):
                                s_match.iat[j] = 3
                                flagged += 1

                        matches_log.append({"row": i+1, "סכום בנק": amt, "סט מס׳ תשלום": ", ".join(sorted(close_paynums)) or "—", "התאמה": f"סומנו {flagged} שורות (מצטבר)"})

                    df[m_match] = s_match
                    df.to_excel(wr, index=False, sheet_name="DataSheet")

                    # דוח דיאגנוסטיקה
                    if len(aux_groups):
                        aux_groups.to_excel(wr, index=False, sheet_name="דוח_קבוץ_עזר")
                    pd.DataFrame(matches_log).to_excel(wr, index=False, sheet_name="לוג_התאמות_3")

                out.seek(0)

            st.success("התאמה 3 הושלמה. ראי דוחות 'דוח_קבוץ_עזר' ו-'לוג_התאמות_3'.")
            st.download_button("⬇️ הורד/י קובץ מעודכן",
                               data=out.getvalue(),
                               file_name="מקור_עם_התאמה_3.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.info("נא להעלות גם קובץ מקור וגם קובץ עזר כדי לסמן התאמות 3.")
