import streamlit as st
import pandas as pd
import io
from typing import Optional
from datetime import date as dt_date, datetime as dt_datetime

# --- Page setup ---
st.set_page_config(page_title="Incentive Automation System", layout="wide")
st.title("Incentive Automation System — RCIL")

# ---------- Choose Project ----------
project_choice = st.selectbox("Select Incentive Project", ["Monthly SD Incentive", "Spot Incentive"])

# ----------------- MONTHLY SD INCENTIVE LOGIC -----------------
if project_choice == "Monthly SD Incentive":
    st.markdown("""
    ### SD Incentive Automation System
    Upload OLD-OS, NEW-OS, and SD-INCENTIVE files to compute incentives for OLD and NEW customers.
    """)

    # ---------- Constants & helpers ----------
    COMMON_DEPOSIT_COLS = [
        "deposit amount", "depositamount", "deposit", "amount", "outstanding", "outstanding amount", "balance", "bal"
    ]
    COMMON_NEWACC_COLS = [
        "new account number", "newaccountnumber", "newaccno", "account number", "accountno", "accno", "new acc no", "newacc",
        "account"
    ]
    COMMON_SCHEME_COLS = ["scheme code", "schemecode", "scheme", "scheme_name", "schemename", "scheme code"]
    COMMON_BRANCH_COLS = ["branch name", "branch", "branchname"]
    COMMON_CUSTID_COLS = ["customer id", "customerid", "cust id", "custid", "customer_id"]
    COMMON_CUSTNAME_COLS = ["customer name", "customername", "cust name", "custname", "customer_name"]
    COMMON_CANVASS_COLS = ["canvassed by", "canvassedby", "canvasser", "employee", "collected by"]
    COMMON_OLD_INC_COLS = ["old_incentive", "oldincentive", "old_incent", "oldinc", "old incentive"]
    COMMON_REAL_DATE_COLS = [
        "realisation date", "realization date", "realiztation date", "realisationdate", "realisation", "realized date",
        "realisation date", "realisation_date", "realised date", "date of realisation", "realisationdate"
    ]

    def normalize_col_name(c: str) -> str:
        return ''.join(ch.lower() for ch in str(c) if ch.isalnum())

    def find_column(df: pd.DataFrame, candidates) -> Optional[str]:
        cols_norm = {normalize_col_name(c): c for c in df.columns}
        for cand in candidates:
            key = normalize_col_name(cand)
            if key in cols_norm:
                return cols_norm[key]
        for cand in candidates:
            key = normalize_col_name(cand)
            for k, orig in cols_norm.items():
                if key in k:
                    return orig
        return None

    def read_uploaded_file(uploaded) -> Optional[pd.DataFrame]:
        if uploaded is None:
            return None
        try:
            uploaded.seek(0)
        except Exception:
            pass
        name = getattr(uploaded, 'name', '') or ''
        lower_name = name.lower()
        if lower_name.endswith(('.xls', '.xlsx', '.xlsm', '.xlsb')):
            try:
                return pd.read_excel(uploaded, engine='openpyxl', dtype=str)
            except Exception:
                try:
                    uploaded.seek(0)
                    return pd.read_excel(uploaded)
                except Exception:
                    pass
        try:
            uploaded.seek(0)
            return pd.read_excel(uploaded, engine='openpyxl', dtype=str)
        except Exception:
            pass
        encodings_to_try = ["utf-8", "cp1252", "latin1", "iso-8859-1"]
        for enc in encodings_to_try:
            try:
                uploaded.seek(0)
                return pd.read_csv(uploaded, encoding=enc)
            except Exception:
                continue
        try:
            uploaded.seek(0)
            raw = uploaded.read()
            if isinstance(raw, bytes):
                for enc in encodings_to_try:
                    try:
                        text = raw.decode(enc)
                        return pd.read_csv(io.StringIO(text))
                    except Exception:
                        continue
        except Exception:
            pass
        st.error("Failed to read file. Please provide a valid Excel (.xls/.xlsx/.xlsm) or CSV file.")
        return None

    def parse_interest_value(x) -> float:
        try:
            if x is None or (isinstance(x, float) and pd.isna(x)):
                return 0.0
            if isinstance(x, str):
                s = x.strip()
                if s == '':
                    return 0.0
                if '%' in s:
                    s_clean = s.replace('%', '').replace(',', '')
                    return float(s_clean) / 100.0
                s_clean = s.replace(',', '')
                val = float(s_clean)
                if val > 1:
                    return val / 100.0
                return val
            val = float(x)
            if val > 1:
                return val / 100.0
            return val
        except Exception:
            return 0.0

    def load_incentive_map(df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        sc_col = find_column(df, COMMON_SCHEME_COLS)
        int_col = None
        for maybe in ['interest', 'rate', 'interest rate', 'value']:
            col = find_column(df, [maybe])
            if col:
                int_col = col
                break
        if sc_col is None:
            sc_col = df.columns[0]
        if int_col is None:
            if len(df.columns) > 1:
                int_col = df.columns[1]
            else:
                raise ValueError("Couldn't find interest column in SD-INCENTIVE file")
        df = df[[sc_col, int_col]].rename(columns={sc_col: 'SchemeCode', int_col: 'InterestRaw'})
        df['SchemeCode'] = df['SchemeCode'].astype(str).str.strip()
        df['Interest'] = df['InterestRaw'].apply(parse_interest_value)
        df['SchemeKey'] = df['SchemeCode'].apply(lambda x: normalize_col_name(x))
        df = df.drop_duplicates(subset=['SchemeKey'], keep='last').set_index('SchemeKey')
        return df

    def build_old_incentive_map(df: pd.DataFrame) -> pd.Series:
        df = df.copy()
        newacc_col = find_column(df, COMMON_NEWACC_COLS)
        oldinc_col = find_column(df, COMMON_OLD_INC_COLS)
        if newacc_col is None:
            newacc_col = df.columns[0]
        if oldinc_col is None:
            if len(df.columns) > 1:
                oldinc_col = df.columns[1]
        if newacc_col is None or oldinc_col is None:
            st.error("OLD-OS must contain: New Account Number and OLD_INCENTIVE (column names flexible).")
            return pd.Series(dtype=float)
        s = pd.Series(df[oldinc_col].values, index=df[newacc_col].astype(str).str.strip().apply(normalize_col_name))
        s = s.apply(parse_interest_value)
        s.name = 'OLD_INTEREST'
        return s

    def prepare_os_df(df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        df.columns = [str(c).strip() for c in df.columns]
        return df

    def to_excel_bytes(df: pd.DataFrame) -> bytes:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Incentive')
        return output.getvalue()

    # ---------- UI: File uploads ----------
    with st.sidebar:
        st.header("Upload files")
        old_file = st.file_uploader("OLD-OS (New Account Number + OLD_INCENTIVE)", type=["csv", "xlsx", "xls"], key="old")
        new_file = st.file_uploader("NEW-OS (full outstanding)", type=["csv", "xlsx", "xls"], key="new")
        sd_file = st.file_uploader("SD-INCENTIVE (Scheme Code -> Interest)", type=["csv", "xlsx", "xls"], key="sd")
        st.markdown("---")
        st.write("Date selection for incentive computation")
        calc_date = st.date_input("Calculate up to (inclusive)", value=dt_date.today())

    if isinstance(calc_date, dt_datetime):
        calc_date = calc_date.date()

    old_df = read_uploaded_file(old_file)
    new_df = read_uploaded_file(new_file)
    sd_df = read_uploaded_file(sd_file)

    old_map = pd.Series(dtype=float)
    sd_map = None

    if old_df is not None:
        try:
            old_map = build_old_incentive_map(old_df)
        except Exception as e:
            st.error(f"Failed to parse OLD-OS: {e}")
            old_map = pd.Series(dtype=float)

    if sd_df is not None:
        try:
            sd_map = load_incentive_map(sd_df)
        except Exception as e:
            st.error(f"Failed to parse SD-INCENTIVE: {e}")
            sd_map = None

    st.markdown("---")
    st.write("When ready, press **Compute Combined Incentive** to run the engine.")

    if st.button("COMPUTE COMBINED INCENTIVE"):
        if new_df is None:
            st.error("Please upload NEW-OS (full outstanding) first.")
        else:
            new = prepare_os_df(new_df)
            keys = {
                'deposit': find_column(new, COMMON_DEPOSIT_COLS),
                'newacc': find_column(new, COMMON_NEWACC_COLS),
                'scheme': find_column(new, COMMON_SCHEME_COLS),
                'branch': find_column(new, COMMON_BRANCH_COLS),
                'customer_id': find_column(new, COMMON_CUSTID_COLS),
                'customer_name': find_column(new, COMMON_CUSTNAME_COLS),
                'canvassed_by': find_column(new, COMMON_CANVASS_COLS),
                'realisation_date': find_column(new, COMMON_REAL_DATE_COLS)
            }

            if keys['newacc'] is None or keys['deposit'] is None:
                st.error("Required columns not found in NEW-OS.")
            else:
                new[keys['deposit']] = pd.to_numeric(new[keys['deposit']], errors='coerce').fillna(0)
                new['NewAccVal'] = new[keys['newacc']].astype(str).str.strip()
                real_col = keys.get('realisation_date')
                if real_col and real_col in new.columns:
                    new['_RealisationParsed'] = pd.to_datetime(new[real_col].astype(str).str.strip().replace('nan', ''), errors='coerce', dayfirst=True, infer_datetime_format=True)
                else:
                    new['_RealisationParsed'] = pd.NaT

                interests, remarks, days_list, incentives = [], [], [], []

                for idx, row in new.iterrows():
                    interest_val = 0.0
                    acct_key = normalize_col_name(str(row['NewAccVal']))
                    is_old = False
                    if not old_map.empty and acct_key in old_map.index:
                        interest_val = float(old_map.loc[acct_key])
                        is_old = True
                    else:
                        if keys['scheme'] and keys['scheme'] in row.index and sd_map is not None:
                            sk = normalize_col_name(str(row[keys['scheme']]))
                            if sk in sd_map.index:
                                interest_val = float(sd_map.loc[sk, 'Interest'])
                    remark = "Old Customer" if is_old else "New Customer"
                    deposit_val = float(row[keys['deposit']]) if keys['deposit'] in row.index and pd.notnull(row[keys['deposit']]) else 0.0
                    real_parsed = row.get('_RealisationParsed', pd.NaT)
                    no_of_days = None
                    if pd.notna(real_parsed):
                        try:
                            real_date_only = real_parsed.date()
                            no_of_days = abs((calc_date - real_date_only).days)
                        except Exception:
                            no_of_days = None
                    if is_old:
                        incentive_val = deposit_val * interest_val / 12.0
                    else:
                        if no_of_days is None:
                            incentive_val = pd.NA
                            remark = "New Customer - Invalid Realisation Date"
                        else:
                            incentive_val = deposit_val * interest_val * (no_of_days / 365.0)
                    interests.append(interest_val)
                    remarks.append(remark)
                    days_list.append(no_of_days if no_of_days is not None else pd.NA)
                    incentives.append(incentive_val if (incentive_val is not None) else pd.NA)

                new['_Interest'] = pd.Series(interests, index=new.index)
                new['_Remark'] = pd.Series(remarks, index=new.index)
                new['_No_of_Days'] = pd.Series(days_list, index=new.index)
                new['_Incentive'] = pd.Series(incentives, index=new.index)

                out = pd.DataFrame(index=new.index)
                out['Branch'] = new[keys['branch']] if keys['branch'] in new.columns else ''
                out['Customer ID'] = new[keys['customer_id']] if keys['customer_id'] in new.columns else ''
                out['New Account Number'] = new[keys['newacc']]
                out['Scheme Code'] = new[keys['scheme']] if keys['scheme'] in new.columns else ''
                out['Customer Name'] = new[keys['customer_name']] if keys['customer_name'] in new.columns else ''
                out['Deposit'] = new[keys['deposit']]
                out['Canvassed By'] = new[keys['canvassed_by']] if keys['canvassed_by'] in new.columns else ''
                out['Realisation Date'] = new[real_col] if real_col and real_col in new.columns else ''
                out['Interest'] = new['_Interest']
                out['No_of_Days'] = new['_No_of_Days']
                out['Incentive'] = new['_Incentive']
                out['Remark'] = new['_Remark']

                st.dataframe(out.head(200))
                excel_bytes = to_excel_bytes(out.reset_index(drop=True))
                st.download_button(label="Download Combined Incentive Excel", data=excel_bytes,
                                   file_name="combined_incentive.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ----------------- SPOT INCENTIVE LOGIC -----------------
elif project_choice == "Spot Incentive":
    st.title("SPOT INCENTIVE CALCULATOR")

    # ----------------- File Upload -----------------
    spot_file = st.file_uploader("Upload SD_SPOT_INCENTIVE File", type=["xlsx", "csv"], key="spot")
    issue_file = st.file_uploader("Upload ISSUE_REPORT File", type=["xlsx", "csv"], key="issue")

    # ----------------- Customer ID Threshold -----------------
    threshold_number = st.number_input("Enter Customer ID Number Threshold", value=0, key="threshold")

    # ----------------- Read File Function -----------------
    def read_file(file):
        if file is None:
            return None
        if file.name.endswith(".csv"):
            return pd.read_csv(file)
        else:
            return pd.read_excel(file)

    # ----------------- Process Button -----------------
    if st.button("PROCESS SPOT INCENTIVE"):
        if spot_file is None or issue_file is None:
            st.error("Please upload both files!")
        else:
            # Read files
            spot_df = read_file(spot_file)
            issue_df = read_file(issue_file)

            # Clean column names
            spot_df.columns = spot_df.columns.str.strip().str.replace(" ", "").str.lower()
            issue_df_clean = issue_df.copy()
            issue_df_clean.columns = issue_df_clean.columns.str.strip().str.replace(" ", "").str.lower()

            custid_col = 'customeridnumber'
            acct_col = 'accountnumber'

            # Remove duplicate AccountNumbers from ISSUE_REPORT
            issue_df_clean = issue_df_clean.drop_duplicates(subset=[acct_col])

            # Merge SPOT_INCENTIVE by schemecode
            spot_df_unique = spot_df.drop_duplicates(subset=['schemecode'])

            merged = issue_df_clean.merge(
                spot_df_unique[['schemecode', 'spotoldcus', 'spotnewcus']],
                on="schemecode",
                how="left"
            )

            # Convert numeric columns
            merged['amount'] = pd.to_numeric(merged['amount'], errors='coerce')
            merged['spotoldcus'] = pd.to_numeric(merged['spotoldcus'], errors='coerce')
            merged['spotnewcus'] = pd.to_numeric(merged['spotnewcus'], errors='coerce')
            merged[custid_col] = pd.to_numeric(merged[custid_col], errors='coerce')

            # Calculate Incentives
            def calculate_incentives(row):
                if row[custid_col] > threshold_number:
                    new_incentive = row['amount'] * (row['spotnewcus'] / 100)
                    old_incentive = 0
                else:
                    old_incentive = row['amount'] * (row['spotoldcus'] / 100)
                    new_incentive = 0
                return pd.Series([old_incentive, new_incentive])

            merged[['old_incentive', 'new_incentive']] = merged.apply(calculate_incentives, axis=1)

            # Filter unwanted branches
            merged = merged[~merged['branchname'].str.contains("RELIANT CREDITSFIN", case=False, na=False)]

            final = merged[['branchname', acct_col, custid_col, 'name', 'schemecode', 'amount',
                            'canvasserid', 'canvassername', 'spotoldcus', 'spotnewcus',
                            'old_incentive', 'new_incentive']]

            st.success("✓ Processing Complete")
            st.dataframe(final)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                final.to_excel(writer, index=False, sheet_name='Incentive')
            st.download_button(
                label="Download Excel File",
                data=output.getvalue(),
                file_name="spot_incentive_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
