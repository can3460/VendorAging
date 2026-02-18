import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import os
import time
import plotly.express as px

try:
    import yfinance as yf
except ImportError:
    pass

# ==========================================
# 1. MASTER CONFIGURATION & ADMIN SETUP
# ==========================================
st.set_page_config(page_title="AP Analyzing Suite | Opella", layout="wide", page_icon="🛡️")
VERSION_NO = "v37.0"
MASTER_ADMIN = "can.adiguzel@sanofi.com"
USER_DB = "users.xlsx"

if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if 'view_currency' not in st.session_state: st.session_state['view_currency'] = "Local"
if 'results' not in st.session_state: st.session_state['results'] = None
if 'analysis_run' not in st.session_state: st.session_state['analysis_run'] = False

def load_users():
    if not os.path.exists(USER_DB):
        df = pd.DataFrame([{"email": MASTER_ADMIN, "role": "admin"}])
        df.to_excel(USER_DB, index=False)
        return df
    df = pd.read_excel(USER_DB)
    df.columns = [str(c).lower().strip() for c in df.columns]
    return df

def add_user(new_email):
    users = load_users()
    new_email = new_email.strip().lower()
    if 'email' in users.columns and new_email not in users['email'].values:
        new_row = pd.DataFrame([{"email": new_email, "role": "user"}])
        users = pd.concat([users, new_row], ignore_index=True)
        users.to_excel(USER_DB, index=False)
        return True
    return False

def remove_user(email_to_remove):
    users = load_users()
    if email_to_remove != MASTER_ADMIN:
        users = users[users['email'] != email_to_remove]
        users.to_excel(USER_DB, index=False)
        return True
    return False

# ==========================================
# 2. CORE ENGINE & FILE READERS
# ==========================================
def get_live_rate(base_currency):
    if base_currency == "EUR": return 1.0
    try:
        ticker = yf.Ticker(f"EUR{base_currency}=X")
        history = ticker.history(period="1d")
        if not history.empty:
            return float(history['Close'].iloc[-1])
        return None
    except Exception: 
        return None

def smart_read(file):
    name = file.name.lower()
    if name.endswith((".xlsx", ".xls")): 
        return pd.read_excel(file)
    raw = file.getvalue()
    for enc in ["utf-8-sig", "utf-8", "iso-8859-9", "cp1254", "latin-1", "windows-1252"]:
        try:
            return pd.read_csv(io.BytesIO(raw), encoding=enc, sep=None, engine="python", on_bad_lines="skip")
        except: continue
    raise ValueError(f"Could not read '{file.name}'. Please ensure it is an Excel or UTF-8 CSV file.")

def smart_parse_tb(file):
    try:
        df_tb = smart_read(file)
        gl_name_map, gl_solar_map, gl_balance_map = {}, {}, {}
        
        acc_col = next((c for c in df_tb.columns if 'Account Number' in str(c) or 'G/L' in str(c) or 'Account' in str(c)), None)
        name_col = next((c for c in df_tb.columns if 'Text' in str(c) or 'Description' in str(c)), None)
        solar_col = next((c for c in df_tb.columns if any(x in str(c).lower() for x in ['financial', 'fs item', 'solar', 'group'])), None)
        amt_col = next((c for c in df_tb.columns if any(x in str(c).lower() for x in ['total', 'balance', 'reporting'])), None)

        if not acc_col or not amt_col: return {}, {}, {}

        for _, row in df_tb.iterrows():
            raw_val = str(row[acc_col]).strip()
            clean_acc = raw_val.split('.')[0]
            if clean_acc.isdigit() and len(clean_acc) >= 6:
                gl_name_map[clean_acc] = str(row[name_col]).strip() if name_col and not pd.isna(row[name_col]) else "-"
                gl_solar_map[clean_acc] = str(row[solar_col]).strip() if solar_col and not pd.isna(row[solar_col]) else "-"
                gl_balance_map[clean_acc] = row[amt_col] if not pd.isna(row[amt_col]) else 0
        return gl_name_map, gl_solar_map, gl_balance_map
    except Exception as e: 
        st.error(f"Error reading Trial Balance: {e}")
        return {}, {}, {}

def append_totals(df, numeric_cols, label_col='Vendor'):
    if df.empty: return df
    tot_dict = {c: df[c].sum() for c in numeric_cols if c in df.columns}
    tot_dict[label_col] = 'TOTAL'
    tot_df = pd.DataFrame([tot_dict])
    return pd.concat([df, tot_df], ignore_index=True)

def generate_html_report(dfs, titles, display_curr, rate):
    html = f"<html><head><style>body{{font-family:sans-serif;padding:20px;}}h2{{color:#064e3b;border-bottom:2px solid #064e3b;padding-bottom:5px;}}table{{border-collapse:collapse;width:100%;margin-bottom:30px;font-size:12px;}}th{{background:#f1f5f9;padding:10px;border:1px solid #cbd5e1;text-align:right;}}td{{padding:8px;border:1px solid #cbd5e1;text-align:right;}}td:first-child, th:first-child{{text-align:left;font-weight:bold;}}</style></head><body>"
    html += f"<h1>Opella AP Analyzing Suite</h1><p><b>Currency View:</b> {display_curr} | <b>Date:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')} | <b>EUR Rate:</b> {rate:,.4f}</p>"
    for df, title in zip(dfs, titles):
        if df is not None and not df.empty:
            html += f"<h2>{title}</h2>"
            df_fmt = df.copy()
            for col in df_fmt.select_dtypes(include=[np.number]).columns:
                df_fmt[col] = df_fmt[col].apply(lambda x: f"{x:,.0f}" if not pd.isna(x) else "")
            html += df_fmt.to_html(index=False)
    html += "</body></html>"
    return html

# Excel Formatlama Fonksiyonu - NaN Hatası Giderildi
def format_excel_sheet(writer, df, sheet_name):
    # NaN değerleri boş stringe çeviriyoruz ki xlsxwriter hata vermesin
    df = df.fillna("")
    
    df.to_excel(writer, sheet_name=sheet_name, index=False)
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    header_format = workbook.add_format({
        'bold': True, 'font_color': '#fde68a', 'bg_color': '#064e3b', 
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    cell_format = workbook.add_format({'border': 1})
    num_format = workbook.add_format({'border': 1, 'num_format': '#,##0'})
    
    total_row_format = workbook.add_format({
        'bold': True, 'font_color': '#fde68a', 'bg_color': '#064e3b', 
        'border': 1, 'num_format': '#,##0'
    })
    
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
        column_len = max(df[value].astype(str).map(len).max(), len(str(value))) + 2
        
        # Sütunun geneli numeric ise
        is_numeric = pd.to_numeric(df[value].replace('', np.nan), errors='coerce').notna().any()
        if is_numeric:
            worksheet.set_column(col_num, col_num, min(column_len, 25), num_format)
        else:
            worksheet.set_column(col_num, col_num, min(column_len, 45), cell_format)
            
    if not df.empty and ('TOTAL' in df.iloc[-1].values):
        last_row_idx = len(df)
        for col_num, value in enumerate(df.iloc[-1].values):
            val_to_write = "" if pd.isna(value) or value == "" else value
            worksheet.write(last_row_idx, col_num, val_to_write, total_row_format)

# ==========================================
# 3. UI & LOGIN SYSTEM (OPELLA CSS)
# ==========================================
st.markdown("""
<style>
/* OPELLA DARK GREEN SIDEBAR */
[data-testid="stSidebar"] { background: linear-gradient(160deg, #022c22 0%, #064e3b 100%); }

/* Sidebar Base Text Colors */
[data-testid="stSidebar"] label, 
[data-testid="stSidebar"] p, 
[data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] p { 
    color: #F8FAFC !important; 
    font-weight: 500; 
}

/* FIX FOR DROPDOWNS & INPUTS (Bembeyaz Arka Plan, Simsiyah Metin) */
[data-testid="stSidebar"] input { 
    color: #000000 !important; 
    background-color: #ffffff !important; 
    -webkit-text-fill-color: #000000 !important;
    font-weight: 800 !important;
}
[data-testid="stSidebar"] div[data-baseweb="select"] > div { 
    background-color: #ffffff !important; 
    cursor: pointer; 
}
[data-testid="stSidebar"] div[data-baseweb="select"] span { 
    color: #000000 !important; 
    -webkit-text-fill-color: #000000 !important;
    font-weight: 800 !important; 
}
/* Dropdown List Items */
ul[role="listbox"] { background-color: #ffffff !important; }
ul[role="listbox"] li { color: #000000 !important; font-weight: bold !important; }

/* FIX FOR FILE UPLOADER (Bembeyaz Arka Plan, Simsiyah Metin) */
[data-testid="stFileUploadDropzone"] { 
    background-color: #ffffff !important; 
    border: 2px dashed #94A3B8 !important;
}
[data-testid="stFileUploadDropzone"] * {
    color: #000000 !important;
    -webkit-text-fill-color: #000000 !important;
    font-weight: 800 !important;
}

/* Sidebar Buttons */
[data-testid="stSidebar"] button { 
    background-color: #475569 !important; 
    border: 1px solid #94A3B8 !important; 
    color: #FFFFFF !important; 
    font-weight: bold !important; 
    border-radius: 6px; 
}
[data-testid="stSidebar"] button:hover { background-color: #64748B !important; border-color: #CBD5E1 !important; }

/* Main Area Download Buttons */
[data-testid="stDownloadButton"] button {
    background-color: #064e3b !important;
    color: #fde68a !important; 
    border: 2px solid #022c22 !important;
    font-weight: 800 !important;
    border-radius: 8px !important;
}
[data-testid="stDownloadButton"] button p { color: #fde68a !important; font-weight: 800 !important; }
[data-testid="stDownloadButton"] button:hover {
    background-color: #022c22 !important; color: #ffffff !important; border-color: #fde68a !important;
}
[data-testid="stDownloadButton"] button:hover p { color: #ffffff !important; }

/* Dashboard Cards */
.kpi-row { display:flex; gap:12px; margin-bottom:18px; flex-wrap:wrap; }
.kpi-card { flex:1; min-width:180px; background:#fff; border-radius:10px; padding:15px; box-shadow:0 1px 3px rgba(0,0,0,.1); border-top:4px solid #E5E7EB; }
.kpi-card.blue { border-top-color:#1A56DB; }
.kpi-card.purple { border-top-color:#7E3AF2; }
.kpi-card.green { border-top-color:#057A55; }
.kpi-card.amber { border-top-color:#D97706; }
.kpi-card.red { border-top-color:#E02424; }
.kpi-label { font-size:0.75rem; text-transform:uppercase; color:#64748B; font-weight:700; margin-bottom:5px; }
.kpi-value { font-size:1.6rem; font-weight:800; color:#0F172A; }
.kpi-sub { font-size:0.7rem; color:#64748B; margin-top:5px; }

/* Dataframes Full Width */
[data-testid="stDataFrame"] { width: 100% !important; }
</style>
""", unsafe_allow_html=True)

if not st.session_state['logged_in']:
    st.markdown(f"""<div style="position: fixed; top: 15px; right: 20px; background: #e0e7ff; color: #3730a3; padding: 5px 15px; border-radius: 20px; font-size: 13px; font-weight: bold; border: 1px solid #c7d2fe; z-index: 9999;">{VERSION_NO}</div>""", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        if os.path.exists("logo.png"):
            st.image("logo.png", width=250)
        else:
            st.markdown("<h1 style='text-align: center; color:#064e3b;'>Opella</h1>", unsafe_allow_html=True)
            
        st.markdown("<h2 style='text-align: center; color:#333;'>AP Analyzing Suite</h2>", unsafe_allow_html=True)
        with st.form("login_form"):
            email_input = st.text_input("Corporate Email", placeholder="name.surname@sanofi.com").strip().lower()
            if st.form_submit_button("Secure Login", use_container_width=True):
                users = load_users()
                if 'email' in users.columns and email_input in users['email'].values:
                    st.session_state.update({'logged_in': True, 'user_name': email_input.split('@')[0].replace('.',' ').title(), 'user_email': email_input})
                    st.rerun()
                else: st.error("Unauthorized access. Please contact System Administrator.")
    st.stop()

# ==========================================
# 4. NAVIGATION & SIDEBAR
# ==========================================
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    
    st.markdown(f"👤 **{st.session_state['user_name']}**")
    st.divider()
    
    menu_options = ["🏠 Analysis Home"]
    if st.session_state['user_email'] == MASTER_ADMIN:
        menu_options.append("🛠️ Manage Users")
    
    page = st.radio("Navigation Menu", menu_options)
    st.divider()

    if page == "🏠 Analysis Home":
        uploaded_file = st.file_uploader("1. FBL1N Report (Mandatory)", type=["xlsx", "xls", "csv"])
        tb_file = st.file_uploader("2. Trial Balance F.01 (Mandatory)", type=["xlsx", "xls", "csv"])
        
        currency = st.selectbox("Base Currency", ["TRY", "EUR", "USD", "GBP", "EGP"], index=0)
        
        if st.button("🌐 Sync Online EUR Rate"):
            live = get_live_rate(currency)
            if live: 
                st.session_state['cur_val'] = live
                st.success(f"Live Rate Synced: {live:.4f}")
            else:
                st.error("Failed to fetch live rate. Please enter manually.")
        
        eur_rate = st.number_input(f"1 EUR = ? {currency}", value=st.session_state.get('cur_val', 35.00), format="%.4f")
        
        display_unit = f"k{currency}" if st.session_state['view_currency'] == "Local" else "kEUR"
        scalar = 1000 if st.session_state['view_currency'] == "Local" else (1000 * eur_rate)

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Logout"): st.session_state.clear(); st.rerun()

# ==========================================
# 5. ADMIN PANEL PAGE
# ==========================================
if page == "🛠️ Manage Users":
    st.title("🛠️ System Administration")
    c1, c2 = st.columns(2)
    with c1:
        with st.form("add_user_form"):
            new_u = st.text_input("New Colleague Email")
            if st.form_submit_button("Grant Access"):
                if add_user(new_u): st.success(f"Access granted to {new_u}"); st.rerun()
                else: st.warning("User already exists in the system.")
    with c2:
        all_users = load_users()
        st.write("Authorized Personnel:")
        st.dataframe(all_users, use_container_width=True)
        user_to_del = st.selectbox("Select user to revoke access", all_users[all_users['email'] != MASTER_ADMIN]['email'].values)
        if st.button("Revoke Access"):
            if remove_user(user_to_del): st.success("Access successfully revoked."); st.rerun()

# ==========================================
# 6. ANALYSIS HOME PAGE (THE CORE)
# ==========================================
elif page == "🏠 Analysis Home":
    st.markdown(f"""<div style="display: flex; justify-content: space-between; align-items: center;"><div><h1 style="margin:0; color:#064e3b;">📊 AP Analyzing Suite</h1><p style="color:#64748b; margin:0;">Operational Audit & Intelligence Dashboard</p></div><div style="text-align: right; color: #94a3b8; font-size: 11px;">Developed by <b>Can Adiguzel</b><br>{VERSION_NO} | {display_unit} View</div></div>""", unsafe_allow_html=True)
    st.markdown("<hr style='margin-top:10px; margin-bottom:20px;'>", unsafe_allow_html=True)

    col_t1, col_t2 = st.columns([8, 2])
    with col_t2:
        toggle_label = "Switch to kEUR View" if st.session_state['view_currency'] == "Local" else f"Switch to k{currency} View"
        if st.button(f"🔄 {toggle_label}", use_container_width=True):
            st.session_state['view_currency'] = "EUR" if st.session_state['view_currency'] == "Local" else "Local"
            st.rerun()

    if uploaded_file and tb_file:
        btn_text = "🔄 Refresh Data & Re-Run Analysis" if st.session_state['analysis_run'] else "🚀 Run Analysis Engine"
        btn_type = "secondary" if st.session_state['analysis_run'] else "primary"
        
        if st.button(btn_text, type=btn_type, use_container_width=True):
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            status_text.text("Reading Trial Balance (F.01) and mapping SOLAR codes...")
            name_map, solar_map, tb_bal_map = smart_parse_tb(tb_file)
            progress_bar.progress(30)
            time.sleep(0.5)
            
            status_text.text("Processing FBL1N open items...")
            df_raw = smart_read(uploaded_file)
            df = df_raw.dropna(subset=['Document Number']) if 'Document Number' in df_raw.columns else df_raw
            
            df['Amount'] = pd.to_numeric(df.get('Amount in local currency', df.get('Amount', 0)), errors='coerce').fillna(0)
            df['GL'] = df.get('G/L Account', df.get('GL Account', '')).astype(str).str.split('.').str[0]
            df['Vendor'] = df.get('Vendor name', df.get('Supplier', pd.Series(dtype=str))).astype(str)
            
            progress_bar.progress(50)
            status_text.text("Executing Segmentation and Aging calculations...")
            
            df['SOLAR Code'] = df['GL'].map(solar_map).fillna("Unknown")
            
            def get_segment(solar):
                s = str(solar).strip()
                if s == "42905": return "ICO"
                if s == "42006": return "Employee"
                if s == "24018": return "Prepayment"
                if s == "40000": return "3rd Party"
                return "Other"
                
            df['Segment'] = df['SOLAR Code'].apply(get_segment)
            
            report_date = pd.to_datetime(df['Posting Date']).max() if 'Posting Date' in df.columns else pd.Timestamp(datetime.now())
            buckets = ["Not Due", "1-30 Days", "31-90 Days", "91-360 Days", "360+ Days"]
            
            def calc_bucket(pay_date):
                if pd.isna(pay_date): return "Not Due"
                days = (report_date - pay_date).days
                if days < 0: return "Not Due"
                elif days <= 30: return "1-30 Days"
                elif days <= 90: return "31-90 Days"
                elif days <= 360: return "91-360 Days"
                else: return "360+ Days"
                
            date_col = 'Payment date' if 'Payment date' in df.columns else ('Due Date' if 'Due Date' in df.columns else 'Document Date')
            df['Bucket'] = pd.to_datetime(df[date_col], errors='coerce').apply(calc_bucket)
            
            progress_bar.progress(75)
            status_text.text("Building Audit & Reconciliation matrices...")
            
            vendor_net = df[df['Segment'].isin(['3rd Party', 'ICO', 'Employee'])].groupby('Vendor')['Amount'].sum().reset_index()
            debit_vendors = vendor_net[vendor_net['Amount'] > 0]['Vendor'].tolist()
            
            gl_fbl1n_sums = df.groupby('GL')['Amount'].sum().to_dict()
            rec_list, gap_list = [], []
            payable_solar_groups = ['40000', '42905', '42006', '24018']
            
            fbl1n_unique_gls = set(df['GL'].unique())
            for gl in fbl1n_unique_gls:
                f_val = gl_fbl1n_sums.get(gl, 0)
                tb_val = tb_bal_map.get(gl, 0)
                diff = tb_val - f_val
                status = "✅ Matched" if abs(diff) < 1.0 else "⚠️ Mismatch"
                rec_list.append({
                    "GL Account": gl, "Description": name_map.get(gl, "-"), 
                    "SOLAR Group": str(solar_map.get(gl, "")).strip(), 
                    "F.01 Balance": tb_val, "FBL1N Balance": f_val, 
                    "Difference": diff, "Status": status
                })
                
            for gl, tb_val in tb_bal_map.items():
                s_code = str(solar_map.get(gl, "")).strip()
                if s_code in payable_solar_groups and gl not in fbl1n_unique_gls and abs(tb_val) >= 1:
                    gap_list.append({
                        "GL Account": gl, "Description": name_map.get(gl, "-"), 
                        "SOLAR Group": s_code, "F.01 Balance": tb_val
                    })
            
            rec_df = pd.DataFrame(rec_list)
            gap_df = pd.DataFrame(gap_list)
            
            progress_bar.progress(100)
            status_text.text("Analysis Engine execution completed successfully!")
            time.sleep(1)
            progress_bar.empty()
            status_text.empty()
            
            st.session_state['analysis_run'] = True
            st.session_state['results'] = {
                'raw_data': df,
                'rec_df': rec_df,
                'gap_df': gap_df,
                'buckets': buckets,
                'debit_vendors': debit_vendors
            }

    # ==========================================
    # DASHBOARD RESULTS RENDER
    # ==========================================
    if st.session_state['results']:
        res = st.session_state['results']
        df = res['raw_data']
        buckets = res['buckets']
        debit_vendors = res['debit_vendors']
        
        def sc(val): return val / scalar

        payables_df = df[df['Segment'].isin(['3rd Party', 'ICO', 'Employee'])]
        prep_df = df[df['Segment'] == 'Prepayment']
        debit_df = payables_df[payables_df['Vendor'].isin(debit_vendors)]

        tot_debt = payables_df['Amount'].sum()
        prep_total = prep_df['Amount'].sum()
        debit_total = debit_df['Amount'].sum() if not debit_df.empty else 0
        ico_total = df[df['Segment'] == 'ICO']['Amount'].sum()
        emp_total = df[df['Segment'] == 'Employee']['Amount'].sum()

        st.markdown(f"""
        <div class="kpi-row">
          <div class="kpi-card blue">
            <div class="kpi-label">Total Trade Debt (40000)</div>
            <div class="kpi-value">{display_unit.replace('k','')} {sc(abs(tot_debt)):,.0f}</div>
            <div class="kpi-sub">Open Items: {len(payables_df):,} | Vendors: {payables_df['Vendor'].nunique():,}</div>
          </div>
          <div class="kpi-card amber">
            <div class="kpi-label">Total Prepayments (24018)</div>
            <div class="kpi-value">{display_unit.replace('k','')} {sc(abs(prep_total)):,.0f}</div>
            <div class="kpi-sub">Vendors with advances: {prep_df['Vendor'].nunique():,}</div>
          </div>
          <div class="kpi-card red">
            <div class="kpi-label">Debit Balances (Net Positive)</div>
            <div class="kpi-value">{display_unit.replace('k','')} {sc(abs(debit_total)):,.0f}</div>
            <div class="kpi-sub">Debit Vendors: {len(debit_vendors):,}</div>
          </div>
          <div class="kpi-card purple">
            <div class="kpi-label">ICO Balance (42905)</div>
            <div class="kpi-value">{display_unit.replace('k','')} {sc(abs(ico_total)):,.0f}</div>
            <div class="kpi-sub">Intercompany Group Code</div>
          </div>
          <div class="kpi-card green">
            <div class="kpi-label">Employee Balance (42006)</div>
            <div class="kpi-value">{display_unit.replace('k','')} {sc(abs(emp_total)):,.0f}</div>
            <div class="kpi-sub">Staff Payables Group Code</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["📊 Aging Analytics", "💳 Prepayments", "⚖️ Debit Balances", "🏦 GL Breakdown", "🔄 F.01 Reconciliation", "📤 Export Center"])

        with tab1:
            st.markdown(f"### Payables Aging Summary ({display_unit})")
            aging_summary = payables_df.groupby('Bucket')['Amount'].sum().abs().reset_index()
            aging_summary['Bucket'] = pd.Categorical(aging_summary['Bucket'], categories=buckets, ordered=True)
            aging_summary = aging_summary.sort_values('Bucket')
            
            c1, c2 = st.columns([1, 2])
            with c1:
                disp_aging = aging_summary.copy()
                disp_aging['Amount Scaled'] = disp_aging['Amount'].apply(lambda x: sc(x))
                disp_aging = append_totals(disp_aging, ['Amount Scaled'], label_col='Bucket')
                st.dataframe(disp_aging[['Bucket', 'Amount Scaled']].style.format({'Amount Scaled': "{:,.0f}"}), use_container_width=True, hide_index=True)
            with c2:
                fig = px.bar(aging_summary, x='Bucket', y=aging_summary['Amount']/scalar, text_auto=',.0f', title=f"Aging Distribution ({display_unit})", color='Bucket', color_discrete_sequence=px.colors.qualitative.Pastel)
                fig.update_traces(textposition='outside')
                fig.update_layout(showlegend=False, xaxis_title="", yaxis_title=display_unit, margin=dict(t=40, b=0, l=0, r=0))
                st.plotly_chart(fig, use_container_width=True)

            st.divider()
            st.markdown(f"### Top 10 Vendor Aging ({display_unit})")
            ap_full_df = payables_df.pivot_table(index='Vendor', columns='Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
            ap_full_df['Total'] = ap_full_df.sum(axis=1)
            
            top10_disp = ap_full_df.sort_values('Total', key=abs, ascending=False).head(10).reset_index()
            for col in buckets + ['Total']: top10_disp[col] = top10_disp[col].apply(lambda x: sc(x))
            top10_disp = append_totals(top10_disp, buckets + ['Total'], label_col='Vendor')
            st.dataframe(top10_disp.style.format({c: "{:,.0f}" for c in buckets+['Total']}), use_container_width=True, hide_index=True)

        with tab2:
            st.markdown(f"### Prepayments Detail - SOLAR 24018 ({display_unit})")
            if prep_df.empty:
                st.info("No prepayment line items detected matching SOLAR code 24018.")
                prep_full_df = pd.DataFrame()
            else:
                prep_full_df = prep_df.pivot_table(index='Vendor', columns='Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
                prep_full_df['Total'] = prep_full_df.sum(axis=1)
                
                prep_disp = prep_full_df.sort_values('Total', key=abs, ascending=False).reset_index()
                for col in buckets + ['Total']: prep_disp[col] = prep_disp[col].apply(lambda x: sc(x))
                prep_disp = append_totals(prep_disp, buckets + ['Total'], label_col='Vendor')
                st.dataframe(prep_disp.style.format({c: "{:,.0f}" for c in buckets+['Total']}), use_container_width=True, hide_index=True)

        with tab3:
            st.markdown(f"### Debit Balance Details (Net Positive Vendors) - ({display_unit})")
            if debit_df.empty:
                st.success("✅ No vendors with net debit balances found.")
                debit_full_df = pd.DataFrame()
            else:
                debit_full_df = debit_df.pivot_table(index='Vendor', columns='Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
                debit_full_df['Total'] = debit_full_df.sum(axis=1)
                
                debit_disp = debit_full_df.sort_values('Total', key=abs, ascending=False).head(10).reset_index()
                for col in buckets + ['Total']: debit_disp[col] = debit_disp[col].apply(lambda x: sc(x))
                debit_disp = append_totals(debit_disp, buckets + ['Total'], label_col='Vendor')
                
                st.markdown(f"**Top 10 Vendors in Debit Position**")
                st.dataframe(debit_disp.style.format({c: "{:,.0f}" for c in buckets+['Total']}), use_container_width=True, hide_index=True)

        with tab4:
            st.markdown(f"### Detailed GL Breakdown ({display_unit})")
            gl_pivot = df.pivot_table(index=['SOLAR Code', 'GL'], columns='Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
            gl_pivot['Total Balance'] = gl_pivot.sum(axis=1)
            gl_disp = gl_pivot.reset_index().sort_values('Total Balance', key=abs, ascending=False)
            
            for col in buckets + ['Total Balance']: gl_disp[col] = gl_disp[col].apply(lambda x: sc(x))
            gl_disp = append_totals(gl_disp, buckets + ['Total Balance'], label_col='SOLAR Code')
            gl_disp.loc[gl_disp['SOLAR Code'] == 'TOTAL', 'GL'] = '' 
            st.dataframe(gl_disp.style.format({c: "{:,.0f}" for c in buckets+['Total Balance']}), use_container_width=True, hide_index=True)

        with tab5:
            st.markdown(f"### AP Sub-Ledger GLs vs Trial Balance ({display_unit})")
            rec_df = res['rec_df'].copy()
            if rec_df.empty:
                st.warning("No Trial Balance data uploaded, or no overlapping accounts found.")
            else:
                disp_rec = rec_df.sort_values('Difference', key=abs, ascending=False)
                for c in ['F.01 Balance', 'FBL1N Balance', 'Difference']: disp_rec[c] = disp_rec[c].apply(lambda x: sc(x))
                disp_rec = append_totals(disp_rec, ['F.01 Balance', 'FBL1N Balance', 'Difference'], label_col='GL Account')
                disp_rec.loc[disp_rec['GL Account'] == 'TOTAL', ['Description', 'SOLAR Group', 'Status']] = ''
                
                st.dataframe(disp_rec.style.format({c: "{:,.0f}" for c in ['F.01 Balance', 'FBL1N Balance', 'Difference']})
                             .applymap(lambda x: 'background-color: #dcfce7; color: #065f46; font-weight: bold;' if x == '✅ Matched' else ('background-color: #fee2e2; color: #991b1b; font-weight: bold;' if x == '⚠️ Mismatch' else ''), subset=['Status']), 
                             use_container_width=True, hide_index=True)
                             
                gap_df = res['gap_df']
                if not gap_df.empty:
                    st.divider()
                    st.markdown("#### 🚨 Advisory Note for Head of Finance Operations (HFO)")
                    st.error("**ATTENTION:** The following GL accounts carry balances under Payable SOLAR groups (40000, 24018, 42905, 42006) in the Trial Balance (F.01), but are completely absent from the Sub-Ledger (FBL1N) open items report. Immediate investigation is recommended.")
                    
                    disp_gap = gap_df.sort_values('F.01 Balance', key=abs, ascending=False)
                    disp_gap['F.01 Balance'] = disp_gap['F.01 Balance'].apply(lambda x: sc(x))
                    disp_gap = append_totals(disp_gap, ['F.01 Balance'], label_col='GL Account')
                    disp_gap.loc[disp_gap['GL Account'] == 'TOTAL', ['Description', 'SOLAR Group']] = ''
                    st.dataframe(disp_gap.style.format({'F.01 Balance': "{:,.0f}"}), use_container_width=True, hide_index=True)

        with tab6:
            st.markdown("### 📥 Report Export Hub")
            col_ex1, col_ex2 = st.columns(2)
            
            with col_ex1:
                st.markdown("<div style='background:#fff;border:1px solid #E5E7EB;border-radius:10px;padding:16px;min-height:110px;'><b>🌐 HTML Dashboard</b><br/><span style='font-size:.75rem;color:#6B7280;'>Printable web layout of the analysis.</span></div>", unsafe_allow_html=True)
                
                ex_ap_html = ap_full_df.sort_values('Total', key=abs, ascending=False).reset_index()
                for col in buckets + ['Total']: ex_ap_html[col] = ex_ap_html[col].apply(lambda x: sc(x))
                ex_ap_html = append_totals(ex_ap_html, buckets + ['Total'], label_col='Vendor')
                
                html_out = generate_html_report([ex_ap_html, gl_disp, top10_disp, disp_rec, disp_gap if not gap_df.empty else None], ["1. Full AP Aging Summary", "2. GL Breakdown", "3. Top 10 Payables", "4. Reconciliation Status", "5. Missing FBL1N Items (HFO Advisory)"], display_unit, eur_rate)
                st.download_button("📄 Download HTML Report", html_out, f"Opella_Dashboard_{display_unit}.html", "text/html", use_container_width=True)
                
            with col_ex2:
                st.markdown("<div style='background:#fff;border:1px solid #E5E7EB;border-radius:10px;padding:16px;min-height:110px;'><b>📊 Detailed Excel Report</b><br/><span style='font-size:.75rem;color:#6B7280;'>Full Excel pack with Opella formatting.</span></div>", unsafe_allow_html=True)
                
                output_full = io.BytesIO()
                # nan_inf_to_errors eklendi
                with pd.ExcelWriter(output_full, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
                    def clean_and_total(df_in, numeric_cols, label_col='Vendor'):
                        if df_in.empty: return df_in
                        d = df_in.copy().reset_index() if df_in.index.name else df_in.copy()
                        sort_c = next((c for c in ['Total', 'Total Balance', 'F.01 Balance', 'Difference'] if c in d.columns), None)
                        if sort_c: d = d.sort_values(sort_c, key=abs, ascending=False)
                        
                        for c in numeric_cols:
                            if c in d.columns: d[c] = (d[c]/scalar).round(0) if d[c].dtype != 'object' else d[c]
                        return append_totals(d, numeric_cols, label_col)
                    
                    ex_ap   = clean_and_total(ap_full_df, buckets + ['Total'], 'Vendor')
                    ex_prep = clean_and_total(prep_full_df, buckets + ['Total'], 'Vendor') if not prep_df.empty else pd.DataFrame()
                    ex_deb  = clean_and_total(debit_full_df, buckets + ['Total'], 'Vendor') if not debit_df.empty else pd.DataFrame()
                    ex_gl   = clean_and_total(gl_pivot.reset_index(), buckets + ['Total Balance'], 'SOLAR Code')
                    ex_rec  = clean_and_total(rec_df, ['F.01 Balance', 'FBL1N Balance', 'Difference'], 'GL Account')
                    ex_gap  = clean_and_total(gap_df, ['F.01 Balance'], 'GL Account')
                    
                    format_excel_sheet(writer, ex_ap, 'AP Aging Summary')
                    if not ex_prep.empty: format_excel_sheet(writer, ex_prep, 'Prepayment Detail')
                    if not ex_deb.empty: format_excel_sheet(writer, ex_deb, 'Debit Balance Detail')
                    format_excel_sheet(writer, ex_gl, 'GL Breakdown')
                    format_excel_sheet(writer, ex_rec, 'Reconciliation Audit')
                    if not ex_gap.empty: format_excel_sheet(writer, ex_gap, 'Recon - Missing FBL1N')
                    format_excel_sheet(writer, df.head(5000), 'Raw Classified Sample')
                    
                st.download_button(
                    "📥 Download Excel Pack", 
                    output_full.getvalue(), 
                    f"Opella_AP_Dashboard_{display_unit}_{datetime.now().strftime('%Y%m%d')}.xlsx", 
                    use_container_width=True, 
                    type="primary"
                )

    elif not uploaded_file or not tb_file:
        st.info("👆 Please upload the required FBL1N and F.01 Trial Balance reports from the sidebar to begin.")
