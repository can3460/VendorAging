"""
╔══════════════════════════════════════════════════════════════════════════════╗
║                    Opella AP Analyzing Suite | VendorFace UI                 ║
║                    Perfect Engine + Stunning UX | Merged Edition             ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""
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
# 1. PAGE CONFIG & MASTER SETUP
# ==========================================
st.set_page_config(page_title="AP Analyzing Suite | Opella", layout="wide", page_icon="🛡️")
VERSION_NO = "v47.5 ULTIMATE"
MASTER_ADMIN = "can.adiguzel@sanofi.com"
USER_DB = "users.xlsx"

# Hide Streamlit Cloud Elements
st.markdown("""
<style>
    header {visibility: hidden !important;}
    .stApp {margin-top: -50px !important;}
    #MainMenu {visibility: hidden !important;}
    footer {visibility: hidden !important;}
    .stDeployButton {display: none !important;}
    .viewerBadge_container__1QSob {display: none !important;}
    .viewerBadge_link__1S137 {display: none !important;}
</style>
""", unsafe_allow_html=True)

# State Management
if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if 'view_currency' not in st.session_state: st.session_state['view_currency'] = "Local"
if 'results' not in st.session_state: st.session_state['results'] = None
if 'analysis_run' not in st.session_state: st.session_state['analysis_run'] = False
if 'theme' not in st.session_state: st.session_state['theme'] = 'light'

def load_users():
    if not os.path.exists(USER_DB):
        default_users = pd.DataFrame([
            {"email": MASTER_ADMIN, "role": "admin"},
            {"email": "can.adiguzel@opella.com", "role": "admin"},
            {"email": "admin@opella.com", "role": "admin"}
        ])
        default_users.to_excel(USER_DB, index=False)
        return default_users
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
# 2. THEME SYSTEM & UI CSS (BUG FIXES)
# ==========================================
def get_theme():
    if st.session_state.theme == 'dark':
        return {'bg': '#0F172A', 'card_bg': '#1E293B', 'text': '#F1F5F9', 'text_sec': '#94A3B8', 'border': '#334155', 'accent': '#3B82F6', 'btn_text': '#FFFFFF'}
    return {'bg': '#F8FAFC', 'card_bg': '#FFFFFF', 'text': '#0F172A', 'text_sec': '#64748B', 'border': '#E2E8F0', 'accent': '#064e3b', 'btn_text': '#000000'}

theme = get_theme()

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
* {{ font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif; }}

html, body, [data-testid="stAppViewContainer"] {{
    background: {theme['bg']}; color: {theme['text']};
}}

[data-testid="stSidebar"] {{ background: linear-gradient(160deg, #022c22 0%, #064e3b 100%); }}
[data-testid="stSidebar"] * {{ color: #F8FAFC !important; }}

/* GLASSMORPHISM CARDS */
.glass-card {{
    background: {theme['card_bg']}; border: 1px solid {theme['border']}; border-radius: 12px;
    padding: 24px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05); transition: all 0.3s ease; margin-bottom: 20px;
}}
.glass-card:hover {{ transform: translateY(-2px); box-shadow: 0 10px 15px rgba(0, 0, 0, 0.1); }}

/* KPI GRID */
.kpi-row {{ display:flex; gap:16px; margin-bottom:24px; flex-wrap:wrap; }}
.kpi-card {{ flex:1; min-width:180px; background:{theme['card_bg']}; border:1px solid {theme['border']}; border-radius:12px; padding:20px; border-top:4px solid var(--accent); }}
.kpi-label {{ font-size:0.75rem; text-transform:uppercase; color:{theme['text_sec']}; font-weight:700; margin-bottom:8px; }}
.kpi-value {{ font-size:1.8rem; font-weight:800; color:{theme['text']}; line-height:1.2; }}
.kpi-sub {{ font-size:0.8rem; color:{theme['text_sec']}; margin-top:4px; }}

/* TABS STYLING */
div[data-testid="stTabs"] button {{ background-color: transparent; border-radius: 8px !important; padding: 10px 24px; font-weight: 600; color: {theme['text_sec']}; border: 1px solid transparent; }}
div[data-testid="stTabs"] button[aria-selected="true"] {{ background-color: {theme['accent']} !important; color: #ffffff !important; border-color: {theme['accent']}; }}

/* 🔥 THE FIX: INPUTS & BUTTONS (DUAL MODE VISIBILITY) 🔥 */
div[data-baseweb="input"] > div {{ background-color: #ffffff !important; border: 1px solid #cbd5e1 !important; }}
div[data-baseweb="input"] input {{ color: #000000 !important; -webkit-text-fill-color: #000000 !important; font-weight: 600 !important; }}

button[kind="primary"] {{ background-color: {theme['accent']} !important; color: #ffffff !important; border: none !important; }}
button[kind="primary"] p {{ color: #ffffff !important; font-weight: bold !important; }}

/* DataFrame / Table styling fix for Dark Mode */
[data-testid="stDataFrame"] {{ background-color: {theme['card_bg']}; }}

.sec-hdr {{ font-size: 1.2rem; font-weight: 700; color: {theme['text']}; margin: 24px 0 12px; border-bottom: 2px solid {theme['accent']}; padding-bottom: 8px; }}
</style>
""", unsafe_allow_html=True)

# ==========================================
# 3. GLOBAL THEME TOGGLE (TOP RIGHT)
# ==========================================
top_col1, top_col2 = st.columns([10, 1])
with top_col2:
    if st.button("🌙 Dark" if st.session_state.theme == 'light' else "☀️ Light", key="global_theme_toggle", use_container_width=True):
        st.session_state.theme = 'dark' if st.session_state.theme == 'light' else 'light'
        st.rerun()

# ==========================================
# 4. LOGIN SCREEN
# ==========================================
if not st.session_state['logged_in']:
    c1, c2, c3 = st.columns([1, 1.2, 1])
    with c2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        if os.path.exists("logo.png"):
            st.image("logo.png", width=250)
        else:
            st.markdown(f"<h1 style='text-align: center; color:{theme['text']};'>Opella</h1>", unsafe_allow_html=True)
            
        st.markdown(f"<h3 style='text-align: center; color:{theme['text_sec']}; font-weight:400; margin-bottom: 30px;'>AP Analyzing Suite {VERSION_NO}</h3>", unsafe_allow_html=True)
        
        with st.form("login_form"):
            email_input = st.text_input("Corporate Email", placeholder="name.surname@opella.com").strip().lower()
            if st.form_submit_button("Secure Access", use_container_width=True):
                if not (email_input.endswith("@sanofi.com") or email_input.endswith("@opella.com")):
                    st.error("🔒 Security Policy: Only @sanofi.com or @opella.com domains are allowed.")
                else:
                    users = load_users()
                    if 'email' in users.columns and email_input in users['email'].values:
                        st.session_state.update({'logged_in': True, 'user_name': email_input.split('@')[0].replace('.',' ').title(), 'user_email': email_input})
                        st.rerun()
                    else:
                        st.warning("⚠️ Your email is not registered in the authorized users list.")
    st.stop()

# ==========================================
# 5. CORE ENGINE & FILE READERS (THE BRAINS)
# ==========================================
def get_live_rate(base_currency):
    if base_currency == "EUR": return 1.0
    try:
        ticker = yf.Ticker(f"EUR{base_currency}=X")
        history = ticker.history(period="1d")
        if not history.empty: return float(history['Close'].iloc[-1])
        return None
    except Exception: return None

def smart_read(file):
    name = file.name.lower()
    if name.endswith((".xlsx", ".xls")): return pd.read_excel(file)
    raw = file.getvalue()
    for enc in ["utf-8-sig", "utf-8", "iso-8859-9", "cp1254", "latin-1", "windows-1252"]:
        try: return pd.read_csv(io.BytesIO(raw), encoding=enc, sep=None, engine="python", on_bad_lines="skip")
        except: continue
    raise ValueError(f"Could not read '{file.name}'.")

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
    if df is None or df.empty: return df
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

def format_excel_sheet(writer, df, sheet_name):
    if df is None or df.empty:
        df = pd.DataFrame({'Data': ['No data available in this category.']})
    df = df.copy()
    for col in df.select_dtypes(include=['datetimetz']).columns:
        df[col] = df[col].dt.tz_localize(None)
    df = df.replace([np.inf, -np.inf], np.nan)
    df.to_excel(writer, sheet_name=sheet_name, index=False, na_rep="")
    
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    header_format = workbook.add_format({'bold': True, 'font_color': '#ffffff', 'bg_color': '#064e3b', 'border': 1})
    cell_format = workbook.add_format({'border': 1})
    num_format = workbook.add_format({'border': 1, 'num_format': '#,##0'})
    total_row_format = workbook.add_format({'bold': True, 'font_color': '#ffffff', 'bg_color': '#064e3b', 'border': 1, 'num_format': '#,##0'})
    date_format = workbook.add_format({'border': 1, 'num_format': 'yyyy-mm-dd'})
    
    worksheet.set_column(0, max(len(df.columns) - 1, 0), 20)
    for col_num, col_name in enumerate(df.columns):
        worksheet.write(0, col_num, str(col_name), header_format)
    for row_num in range(len(df)):
        is_total_row = str(df.iloc[row_num, 0]).strip() == 'TOTAL'
        for col_num, col_name in enumerate(df.columns):
            val = df.iloc[row_num, col_num]
            if pd.isna(val):
                worksheet.write(row_num + 1, col_num, "", cell_format)
                continue
            if is_total_row: worksheet.write(row_num + 1, col_num, val, total_row_format)
            elif pd.api.types.is_numeric_dtype(df[col_name]) and isinstance(val, (int, float)): worksheet.write(row_num + 1, col_num, val, num_format)
            elif pd.api.types.is_datetime64_any_dtype(df[col_name]): worksheet.write(row_num + 1, col_num, val, date_format)
            else: worksheet.write(row_num + 1, col_num, str(val), cell_format)

def toggle_currency_view():
    st.session_state['view_currency'] = "EUR" if st.session_state['view_currency'] == "Local" else "Local"

def prepare_for_display(df, numeric_cols):
    if df is None or df.empty: return df
    df = df.copy()
    df.columns.name = None; df.index.name = None
    for col in numeric_cols:
        if col in df.columns: df[col] = df[col].fillna(0)
    return df

# ==========================================
# 6. NAVIGATION & SIDEBAR
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
        uploaded_file = st.file_uploader("1. FBL1N Report", type=["xlsx", "xls", "csv"])
        tb_file = st.file_uploader("2. Trial Balance F.01", type=["xlsx", "xls", "csv"])
        
        currency_list = ["TRY", "EUR", "USD", "GBP", "EGP", "AUD", "JPY", "VND", "MYR", "SGD", "KRW", "TND", "CNY", "INR", "THB"]
        currency = st.selectbox("Base Currency", currency_list, index=0)
        
        if st.button("🌐 Sync Online EUR Rate"):
            live = get_live_rate(currency)
            if live: 
                st.session_state['cur_val'] = live
                st.success(f"Live Rate Synced: {live:.4f}")
            else:
                st.error("Failed to fetch live rate. Enter manually.")
        
        eur_rate = st.number_input(f"1 EUR = ? {currency}", value=st.session_state.get('cur_val', 35.00), format="%.4f")
        
        display_unit = f"k{currency}" if st.session_state['view_currency'] == "Local" else "kEUR"
        scalar = 1000 if st.session_state['view_currency'] == "Local" else (1000 * eur_rate)

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Logout", use_container_width=True): st.session_state.clear(); st.rerun()

# ==========================================
# 7. ADMIN PANEL PAGE
# ==========================================
if page == "🛠️ Manage Users":
    st.markdown("<div class='sec-hdr'>System Administration</div>", unsafe_allow_html=True)
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
# 8. ANALYSIS HOME PAGE (THE CORE ENGINE)
# ==========================================
elif page == "🏠 Analysis Home":
    st.markdown(f"""
    <div style="background:linear-gradient(135deg, #022c22 0%, #064e3b 100%); padding:24px; border-radius:16px; color:white; margin-bottom:20px; box-shadow:0 10px 30px rgba(0,0,0,0.1);">
        <div style="display: flex; justify-content: space-between; align-items: center;">
            <div>
                <h1 style="margin:0; color:#ffffff; font-size: 2.2rem;">📊 AP Analyzing Suite</h1>
                <p style="color:#cbd5e1; margin:0; font-weight:500;">Operational Intelligence Dashboard | Opella Finance</p>
            </div>
            <div style="text-align: right; color: #f8fafc; font-size: 15px;">
                Developed by <b>Can Adiguzel</b><br>
                {VERSION_NO} | <b>{display_unit} View</b>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    col_t1, col_t2 = st.columns([8, 2])
    with col_t2:
        st.button(f"🔄 Switch to {'kEUR' if st.session_state['view_currency'] == 'Local' else f'k{currency}'} View", key="toggle_curr_btn", on_click=toggle_currency_view, use_container_width=True)

    if uploaded_file and tb_file:
        btn_text = "🔄 Refresh Data & Re-Run Analysis" if st.session_state['analysis_run'] else "🚀 Run Analysis Engine"
        btn_type = "secondary" if st.session_state['analysis_run'] else "primary"
        
        if st.button(btn_text, type=btn_type, use_container_width=True):
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            status_text.text("Reading Trial Balance (F.01) and mapping SOLAR codes...")
            name_map, solar_map, tb_bal_map = smart_parse_tb(tb_file)
            progress_bar.progress(30)
            
            status_text.text("Processing FBL1N open items and fixing Vendor NaNs...")
            df_raw = smart_read(uploaded_file)
            df = df_raw.dropna(subset=['Document Number']) if 'Document Number' in df_raw.columns else df_raw
            
            df['Amount'] = pd.to_numeric(df.get('Amount in local currency', df.get('Amount', 0)), errors='coerce').fillna(0)
            df['GL'] = df.get('G/L Account', df.get('GL Account', '')).astype(str).str.split('.').str[0]
            
            # 🔥 THE FIX: PIVOT DROP NAN BUG 🔥
            v_name_col = df.get('Vendor name', df.get('Supplier', pd.Series(dtype=str))).astype(str)
            df['Vendor name'] = v_name_col.replace(['nan', 'None', ''], np.nan)
            
            v_code_col = df.get('Vendor', pd.Series(dtype=str)).astype(str)
            df['Vendor'] = v_code_col.replace(['nan', 'None', ''], np.nan).fillna('Unknown')
            
            df['Vendor name'] = df['Vendor name'].fillna(df['Vendor'])
            
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
            buckets = ["Not Due (1-90 Days)", "91-180 Days", "181-360 Days", "360+ Days"]
            
            def calc_bucket(pay_date):
                if pd.isna(pay_date): return "Not Due (1-90 Days)"
                days = (report_date - pay_date).days
                if days <= 90: return "Not Due (1-90 Days)"
                elif days <= 180: return "91-180 Days"
                elif days <= 360: return "181-360 Days"
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
                rec_list.append({"GL Account": gl, "Description": name_map.get(gl, "-"), "SOLAR Group": str(solar_map.get(gl, "")).strip(), "F.01 Balance": tb_val, "FBL1N Balance": f_val, "Difference": diff, "Status": status})
                
            for gl, tb_val in tb_bal_map.items():
                s_code = str(solar_map.get(gl, "")).strip()
                if s_code in payable_solar_groups and gl not in fbl1n_unique_gls and abs(tb_val) >= 1:
                    gap_list.append({"GL Account": gl, "Description": name_map.get(gl, "-"), "SOLAR Group": s_code, "F.01 Balance": tb_val})
            
            rec_df = pd.DataFrame(rec_list)
            gap_df = pd.DataFrame(gap_list)
            
            progress_bar.progress(100)
            status_text.text("Analysis Engine execution completed successfully!")
            time.sleep(1)
            progress_bar.empty()
            status_text.empty()
            
            st.session_state['analysis_run'] = True
            st.session_state['results'] = {'raw_data': df, 'rec_df': rec_df, 'gap_df': gap_df, 'buckets': buckets, 'debit_vendors': debit_vendors}

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
          <div class="kpi-card" style="--accent: #3B82F6;">
            <div class="kpi-label">Total Trade Debt (40000)</div>
            <div class="kpi-value">{display_unit} {sc(abs(tot_debt)):,.0f}</div>
            <div class="kpi-sub">Open Items: {len(payables_df):,} | Vendors: {payables_df['Vendor'].nunique():,}</div>
          </div>
          <div class="kpi-card" style="--accent: #F59E0B;">
            <div class="kpi-label">Total Prepayments (24018)</div>
            <div class="kpi-value">{display_unit} {sc(abs(prep_total)):,.0f}</div>
            <div class="kpi-sub">Vendors with advances: {prep_df['Vendor'].nunique():,}</div>
          </div>
          <div class="kpi-card" style="--accent: #EF4444;">
            <div class="kpi-label">Debit Balances (Net Positive)</div>
            <div class="kpi-value">{display_unit} {sc(abs(debit_total)):,.0f}</div>
            <div class="kpi-sub">Debit Vendors: {len(debit_vendors):,}</div>
          </div>
          <div class="kpi-card" style="--accent: #8B5CF6;">
            <div class="kpi-label">ICO Balance (42905)</div>
            <div class="kpi-value">{display_unit} {sc(abs(ico_total)):,.0f}</div>
            <div class="kpi-sub">Intercompany Group Code</div>
          </div>
          <div class="kpi-card" style="--accent: #10B981;">
            <div class="kpi-label">Employee Balance (42006)</div>
            <div class="kpi-value">{display_unit} {sc(abs(emp_total)):,.0f}</div>
            <div class="kpi-sub">Staff Payables Group Code</div>
          </div>
        </div>
        """, unsafe_allow_html=True)

        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["📊 Aging Analytics", "💳 Prepayments", "⚖️ Debit Balances", "🏦 GL Breakdown", "🔄 F.01 Reconciliation", "📤 Export Center"])

        def build_full_aging(segment_df):
            if segment_df.empty: return pd.DataFrame()
            full_df = segment_df.pivot_table(index=['Vendor', 'Vendor name'], columns='Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
            full_df['Total'] = full_df.sum(axis=1)
            return full_df.reset_index()

        df_3rd = payables_df[payables_df['Segment'] == '3rd Party']
        df_ico = payables_df[payables_df['Segment'] == 'ICO']
        df_emp = payables_df[payables_df['Segment'] == 'Employee']

        ap_full_3rd = build_full_aging(df_3rd)
        ap_full_ico = build_full_aging(df_ico)
        ap_full_emp = build_full_aging(df_emp)

        def get_top10(full_df):
            if full_df.empty: return pd.DataFrame()
            t10 = full_df.sort_values('Total', key=abs, ascending=False).head(10).reset_index(drop=True)
            for c in buckets + ['Total']: t10[c] = t10[c].apply(lambda x: sc(x))
            t10 = append_totals(t10, buckets + ['Total'], label_col='Vendor')
            return prepare_for_display(t10, buckets + ['Total'])

        top10_3rd = get_top10(ap_full_3rd)
        top10_ico = get_top10(ap_full_ico)
        top10_emp = get_top10(ap_full_emp)

        with tab1:
            st.markdown(f"<div class='sec-hdr'>Payables Aging Summary ({display_unit})</div>", unsafe_allow_html=True)
            aging_summary = payables_df.groupby('Bucket')['Amount'].sum().abs().reset_index()
            aging_summary['Bucket'] = pd.Categorical(aging_summary['Bucket'], categories=buckets, ordered=True)
            aging_summary = aging_summary.sort_values('Bucket')
            
            c1, c2 = st.columns([1, 2])
            with c1:
                disp_aging = aging_summary.copy()
                disp_aging['Amount Scaled'] = disp_aging['Amount'].apply(lambda x: sc(x))
                disp_aging = append_totals(disp_aging, ['Amount Scaled'], label_col='Bucket')
                disp_aging = prepare_for_display(disp_aging, ['Amount Scaled'])
                st.dataframe(disp_aging[['Bucket', 'Amount Scaled']].style.format({'Amount Scaled': "{:,.0f}"}), use_container_width=True, hide_index=True)
            with c2:
                fig = px.bar(aging_summary, x='Bucket', y=aging_summary['Amount']/scalar, text_auto=',.0f', color='Bucket', color_discrete_sequence=px.colors.qualitative.Pastel)
                fig.update_traces(textposition='outside')
                fig.update_layout(showlegend=False, xaxis_title="", yaxis_title=display_unit, margin=dict(t=20, b=0, l=0, r=0), plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)", font=dict(color=theme['text']))
                fig.update_yaxes(gridcolor=theme['border'])
                st.plotly_chart(fig, use_container_width=True)

            st.markdown(f"<div class='sec-hdr'>Top 10 Vendor Aging by Segment ({display_unit})</div>", unsafe_allow_html=True)
            t_3rd, t_ico, t_emp = st.tabs(["🏭 3rd Party (40000)", "🔗 Intercompany ICO (42905)", "👤 Employee (42006)"])
            with t_3rd:
                if not top10_3rd.empty: st.dataframe(top10_3rd.style.format({c: "{:,.0f}" for c in buckets+['Total']}), use_container_width=True, hide_index=True)
            with t_ico:
                if not top10_ico.empty: st.dataframe(top10_ico.style.format({c: "{:,.0f}" for c in buckets+['Total']}), use_container_width=True, hide_index=True)
            with t_emp:
                if not top10_emp.empty: st.dataframe(top10_emp.style.format({c: "{:,.0f}" for c in buckets+['Total']}), use_container_width=True, hide_index=True)

        with tab2:
            st.markdown(f"<div class='sec-hdr'>Prepayments Detail - SOLAR 24018 ({display_unit})</div>", unsafe_allow_html=True)
            if prep_df.empty: st.info("No prepayment line items detected matching SOLAR code 24018.")
            else:
                prep_full_df = prep_df.pivot_table(index=['Vendor', 'Vendor name'], columns='Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
                prep_full_df['Total'] = prep_full_df.sum(axis=1)
                prep_disp = prep_full_df.reset_index().sort_values('Total', key=abs, ascending=False)
                for col in buckets + ['Total']: prep_disp[col] = prep_disp[col].apply(lambda x: sc(x))
                prep_disp = append_totals(prep_disp, buckets + ['Total'], label_col='Vendor')
                prep_disp = prepare_for_display(prep_disp, buckets + ['Total'])
                st.dataframe(prep_disp.style.format({c: "{:,.0f}" for c in buckets+['Total']}), use_container_width=True, hide_index=True)

        with tab3:
            st.markdown(f"<div class='sec-hdr'>Debit Balance Details (Net Positive Vendors) - ({display_unit})</div>", unsafe_allow_html=True)
            if debit_df.empty: st.success("✅ No vendors with net debit balances found.")
            else:
                debit_full_df = debit_df.pivot_table(index=['Vendor', 'Vendor name'], columns='Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
                debit_full_df['Total'] = debit_full_df.sum(axis=1)
                debit_disp = debit_full_df.reset_index().sort_values('Total', key=abs, ascending=False).head(10)
                for col in buckets + ['Total']: debit_disp[col] = debit_disp[col].apply(lambda x: sc(x))
                debit_disp = append_totals(debit_disp, buckets + ['Total'], label_col='Vendor')
                debit_disp = prepare_for_display(debit_disp, buckets + ['Total'])
                st.dataframe(debit_disp.style.format({c: "{:,.0f}" for c in buckets+['Total']}), use_container_width=True, hide_index=True)

        with tab4:
            st.markdown(f"<div class='sec-hdr'>Detailed GL Breakdown ({display_unit})</div>", unsafe_allow_html=True)
            gl_pivot = df.pivot_table(index=['SOLAR Code', 'GL'], columns='Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
            gl_pivot['Total Balance'] = gl_pivot.sum(axis=1)
            gl_disp = gl_pivot.reset_index().sort_values('Total Balance', key=abs, ascending=False)
            for col in buckets + ['Total Balance']: gl_disp[col] = gl_disp[col].apply(lambda x: sc(x))
            gl_disp = append_totals(gl_disp, buckets + ['Total Balance'], label_col='SOLAR Code')
            gl_disp.loc[gl_disp['SOLAR Code'] == 'TOTAL', 'GL'] = '' 
            gl_disp = prepare_for_display(gl_disp, buckets + ['Total Balance'])
            st.dataframe(gl_disp.style.format({c: "{:,.0f}" for c in buckets+['Total Balance']}), use_container_width=True, hide_index=True)

        with tab5:
            st.markdown(f"<div class='sec-hdr'>AP Sub-Ledger GLs vs Trial Balance ({display_unit})</div>", unsafe_allow_html=True)
            rec_df, gap_df = res['rec_df'].copy(), res['gap_df'].copy()
            if not gap_df.empty:
                gap_disp = gap_df.copy()
                gap_disp['FBL1N Balance'] = 0; gap_disp['Difference'] = gap_disp['F.01 Balance']; gap_disp['Status'] = '🚨 Missing in FBL1N'
                full_rec_df = pd.concat([rec_df, gap_disp], ignore_index=True) if not rec_df.empty else gap_disp
            else: full_rec_df = rec_df.copy()
                
            if full_rec_df.empty: st.warning("No Trial Balance data uploaded, or no overlapping accounts found.")
            else:
                disp_rec = full_rec_df.sort_values('Difference', key=abs, ascending=False)
                for c in ['F.01 Balance', 'FBL1N Balance', 'Difference']: disp_rec[c] = disp_rec[c].apply(lambda x: sc(x))
                disp_rec = append_totals(disp_rec, ['F.01 Balance', 'FBL1N Balance', 'Difference'], label_col='GL Account')
                disp_rec.loc[disp_rec['GL Account'] == 'TOTAL', ['Description', 'SOLAR Group', 'Status']] = ''
                disp_rec = prepare_for_display(disp_rec, ['F.01 Balance', 'FBL1N Balance', 'Difference'])
                st.dataframe(disp_rec.style.format({c: "{:,.0f}" for c in ['F.01 Balance', 'FBL1N Balance', 'Difference']})
                             .applymap(lambda x: 'background-color: #dcfce7; color: #065f46; font-weight: bold;' if x == '✅ Matched' else ('background-color: #fee2e2; color: #991b1b; font-weight: bold;' if x == '⚠️ Mismatch' else ('background-color: #fef08a; color: #854d0e; font-weight: bold;' if 'Missing' in str(x) else '')), subset=['Status']), 
                             use_container_width=True, hide_index=True)

        with tab6:
            st.markdown("<div class='sec-hdr'>Report Export Hub</div>", unsafe_allow_html=True)
            col_ex1, col_ex2 = st.columns(2)
            with col_ex1:
                st.markdown(f"<div class='glass-card' style='text-align:center;'><h3>🌐 HTML Dashboard</h3><p style='color:{theme['text_sec']};'>Printable web layout of the analysis.</p></div>", unsafe_allow_html=True)
                html_out = generate_html_report([top10_3rd, top10_ico, top10_emp, gl_disp, disp_rec], ["1. Top 10 3rd Party Aging", "2. Top 10 ICO Aging", "3. Top 10 Employee Aging", "4. GL Breakdown", "5. Reconciliation Status (Inc. Missing FBL1N)"], display_unit, eur_rate)
                st.download_button("📄 Download HTML Report", html_out, f"Opella_Dashboard_{display_unit}.html", "text/html", use_container_width=True)
            
            with col_ex2:
                st.markdown(f"<div class='glass-card' style='text-align:center;'><h3>📊 Detailed Excel Report</h3><p style='color:{theme['text_sec']};'>Full Excel pack with Opella formatting.</p></div>", unsafe_allow_html=True)
                try:
                    # 🔥 THE BUG FIX FOR ROUNDING LOSS: Aggregation FIRST, then division and rounding 🔥
                    def clean_and_total_full(df_in, numeric_cols, label_col='Vendor'):
                        if df_in is None or df_in.empty: return pd.DataFrame()
                        d = df_in.copy().reset_index() if df_in.index.name else df_in.copy()
                        sort_c = next((c for c in ['Total', 'Total Balance', 'F.01 Balance', 'Difference'] if c in d.columns), None)
                        if sort_c: d = d.sort_values(sort_c, key=abs, ascending=False)
                        
                        # 1. Total computation on raw float values
                        d = append_totals(d, numeric_cols, label_col)
                        
                        # 2. Scale and round the whole set including the exact sum
                        for c in numeric_cols:
                            if c in d.columns: d[c] = (pd.to_numeric(d[c], errors='coerce') / scalar).round(0)
                        return d

                    ex_3rd  = clean_and_total_full(ap_full_3rd, buckets + ['Total'], 'Vendor')
                    ex_ico  = clean_and_total_full(ap_full_ico, buckets + ['Total'], 'Vendor')
                    ex_emp  = clean_and_total_full(ap_full_emp, buckets + ['Total'], 'Vendor')
                    ex_prep = clean_and_total_full(prep_full_df.reset_index(), buckets + ['Total'], 'Vendor') if not prep_df.empty else pd.DataFrame()
                    ex_deb  = clean_and_total_full(debit_full_df.reset_index(), buckets + ['Total'], 'Vendor') if not debit_df.empty else pd.DataFrame()
                    ex_gl   = clean_and_total_full(gl_pivot.reset_index(), buckets + ['Total Balance'], 'SOLAR Code')
                    ex_rec  = clean_and_total_full(full_rec_df, ['F.01 Balance', 'FBL1N Balance', 'Difference'], 'GL Account')
                    
                    output_full = io.BytesIO()
                    with pd.ExcelWriter(output_full, engine='xlsxwriter') as writer:
                        format_excel_sheet(writer, ex_3rd, 'FULL 3rd Party Aging')
                        format_excel_sheet(writer, ex_ico, 'FULL ICO Aging')
                        format_excel_sheet(writer, ex_emp, 'FULL Employee Aging')
                        format_excel_sheet(writer, ex_prep, 'FULL Prepayment Detail')
                        format_excel_sheet(writer, ex_deb, 'FULL Debit Balance')
                        format_excel_sheet(writer, ex_gl, 'GL Breakdown')
                        format_excel_sheet(writer, ex_rec, 'Reconciliation Audit')
                        format_excel_sheet(writer, df.head(5000), 'Raw Classified Sample')
                    
                    excel_data = output_full.getvalue()
                    st.download_button(label="📥 Download Full Excel Data", data=excel_data, file_name=f"Opella_AP_FULL_Data_{display_unit}_{datetime.now().strftime('%Y%m%d')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True, type="primary")
                except Exception as e:
                    st.error(f"🚨 Excel Creation Engine Error: {str(e)}")

    elif not uploaded_file or not tb_file:
        st.info("👆 Please upload the required FBL1N and F.01 Trial Balance reports from the sidebar to begin.")
