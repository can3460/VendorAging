import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import os
import time

# --- GÜVENLİ IMPORT ---
try:
    import yfinance as yf
    YFINANCE_AVAILABLE = True
except ImportError:
    YFINANCE_AVAILABLE = False

# ==========================================
# 1. CONFIGURATION & SETUP
# ==========================================
st.set_page_config(page_title="AP Analyzing Suite | Opella Finance", layout="wide", page_icon="🛡️")
USER_DB_FILE = "users.xlsx"
ADMIN_EMAIL = "can.adiguzel@sanofi.com"
VERSION_NO = "v15.0" 

# ==========================================
# 2. AUTHENTICATION SYSTEM
# ==========================================
def load_user_db():
    if not os.path.exists(USER_DB_FILE):
        initial_users = [
            {"Email": ADMIN_EMAIL, "Name": "Can Adiguzel", "Role": "Admin"},
            {"Email": "AyseDeniz.Sen@sanofi.com", "Name": "AyseDeniz Sen", "Role": "User"},
            {"Email": "Hassan.Sadek@sanofi.com", "Name": "Hassan Sadek", "Role": "User"},
            {"Email": "Omar.Kordy@sanofi.com", "Name": "Omar Kordy", "Role": "User"},
            {"Email": "Rishabh.Tiwari@sanofi.com", "Name": "Rishabh Tiwari", "Role": "User"},
            {"Email": "Molka.Mathlouthi@sanofi.com", "Name": "Molka Mathlouthi", "Role": "User"},
            {"Email": "Shweta.Sharma3@sanofi.com", "Name": "Shweta Sharma", "Role": "User"},
            {"Email": "Prachi.Shukla@sanofi.com", "Name": "Prachi Shukla", "Role": "User"},
            {"Email": "Cedric.Fallu@sanofi.com", "Name": "Cedric Fallu", "Role": "User"}
        ]
        df = pd.DataFrame(initial_users)
        df['Email'] = df['Email'].str.lower().str.strip()
        df['Added_Date'] = datetime.now().strftime("%Y-%m-%d")
        df.to_excel(USER_DB_FILE, index=False)
        return df
    else:
        return pd.read_excel(USER_DB_FILE)

if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False

# --- LOGIN SCREEN ---
if not st.session_state['logged_in']:
    st.markdown(f"""<div style="position: fixed; top: 15px; right: 80px; background: #dbeafe; color: #1e40af; padding: 5px 15px; border-radius: 20px; font-size: 14px; font-weight: bold; font-family: monospace; border: 1px solid #bfdbfe; z-index: 9999;">{VERSION_NO}</div>""", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        if os.path.exists("logo.png"): st.image("logo.png", width=200)
        st.markdown("""<div style='text-align: center;'><h2 style='color:#1e293b; margin-bottom: 0px;'>AP Analyzing Suite</h2></div>""", unsafe_allow_html=True)
        st.write("")
        with st.form("login_form"):
            email_input = st.text_input("Email Address").strip().lower()
            submit_button = st.form_submit_button("Secure Login", type="primary", use_container_width=True)
        if submit_button:
            users_df = load_user_db()
            user_record = users_df[users_df['Email'] == email_input]
            if not user_record.empty:
                st.session_state.update({'logged_in': True, 'user_email': email_input, 'user_name': user_record.iloc[0]['Name'], 'user_role': user_record.iloc[0]['Role']})
                st.rerun()
            else: st.warning("⚠️ Access Denied.")
    st.stop() 

# ==========================================
# 3. HELPER FUNCTIONS
# ==========================================

def get_live_rate(base_currency):
    if not YFINANCE_AVAILABLE: return None
    try:
        ticker = yf.Ticker(f"EUR{base_currency}=X")
        history = ticker.history(period="1d")
        return history['Close'].iloc[-1] if not history.empty else None
    except: return None

def clean_sap_data(df):
    return df.dropna(subset=['Document Number']) if 'Document Number' in df.columns else df

def process_tb_file(file, fbl1n_gl_summary):
    """Mizan dosyasından SOLAR Code (FS Item) ve Name bilgilerini çeker"""
    try:
        df_tb = pd.read_excel(file)
        acc_col = next((c for c in df_tb.columns if 'Account' in str(c) and 'Number' in str(c)), None)
        amt_col = next((c for c in df_tb.columns if 'Total' in str(c) and 'reporting' in str(c)), None)
        solar_col = next((c for c in df_tb.columns if 'Financial' in str(c) and 'Item' in str(c)), None)
        name_col = next((c for c in df_tb.columns if ('Text' in str(c) or 'Description' in str(c)) and 'B/S' in str(c)), None)
        if not name_col: name_col = next((c for c in df_tb.columns if 'Text' in str(c) and 'Account' not in str(c)), None)

        if not acc_col or not amt_col: return None, {}, {}, "Missing columns"

        df_tb = df_tb.dropna(subset=[acc_col])
        df_tb['GL_Account'] = df_tb[acc_col].astype(str).str.strip()
        df_tb = df_tb[df_tb['GL_Account'].str.match(r'^\d+$')]

        gl_name_map = df_tb.set_index('GL_Account')[name_col].to_dict() if name_col else {}
        gl_solar_map = df_tb.set_index('GL_Account')[solar_col].to_dict() if solar_col else {}

        # Reconciliation
        tb_summary = df_tb.groupby('GL_Account')[amt_col].sum().reset_index()
        fbl1n_check = fbl1n_gl_summary.reset_index()[['G/L Account', 'Total Balance']].rename(columns={'G/L Account': 'GL_Account', 'Total Balance': 'FBL1n_Sum'})
        merged = pd.merge(tb_summary, fbl1n_check, on='GL_Account', how='left').fillna(0)
        merged['Difference'] = merged[amt_col] - merged['FBL1n_Sum']
        
        # SOLAR ve Name ekleme
        merged['SOLAR Code'] = merged['GL_Account'].map(gl_solar_map).fillna("-")
        merged['GL Name'] = merged['GL_Account'].map(gl_name_map).fillna("-")
        
        return merged.rename(columns={amt_col: 'TB_Balance'}), gl_name_map, gl_solar_map, "Success"
    except Exception as e: return None, {}, {}, str(e)

def generate_html_report(dfs, titles, currency, rate):
    """Outlook uyumlu HTML raporu (Kıyaslama tablosu dahil)"""
    html = f"""<html><head><style>
    body {{ font-family: sans-serif; }}
    h2 {{ color: #1e40af; border-bottom: 2px solid #e2e8f0; }}
    table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; font-size: 12px; }}
    th {{ background: #f8fafc; padding: 8px; border: 1px solid #cbd5e1; text-align: left; }}
    td {{ padding: 8px; border: 1px solid #cbd5e1; text-align: right; }}
    td:first-child {{ text-align: left; font-weight: bold; }}
    .diff-red {{ color: red; font-weight: bold; }}
    </style></head><body>
    <h1>📊 AP Analyzing Suite - Smart Report</h1>
    <p><b>Date:</b> {datetime.now().strftime('%Y-%m-%d')} | <b>Currency:</b> {currency} | <b>Rate:</b> {rate}</p>"""
    
    for df, title in zip(dfs, titles):
        if df is not None and not df.empty:
            html += f"<h2>{title}</h2>"
            df_html = df.copy()
            # Sayısal formatlama
            for col in df_html.select_dtypes(include=[np.number]).columns:
                df_html[col] = df_html[col].apply(lambda x: f"{x:,.1f}")
            html += df_html.to_html(index=False)
    
    html += "</body></html>"
    return html

# ==========================================
# 4. SIDEBAR
# ==========================================
with st.sidebar:
    if os.path.exists("logo.png"): st.image("logo.png", use_container_width=True)
    st.markdown(f"👤 **{st.session_state['user_name']}**")
    uploaded_file = st.file_uploader("1. FBL1N Report (Required)", type=["xlsx", "xls"])
    tb_file = st.file_uploader("2. Trial Balance F.01 (Optional)", type=["xlsx", "xls"])
    selected_currency = st.selectbox("Local Currency", ["EGP", "TRY", "EUR", "USD", "TND", "AED"], index=0)
    if 'current_eur_rate' not in st.session_state: st.session_state['current_eur_rate'] = 52.50
    eur_rate = st.number_input(f"EUR / {selected_currency}", value=st.session_state['current_eur_rate'], format="%.4f")
    if st.button("🔒 Logout"): st.session_state['logged_in'] = False; st.rerun()

# ==========================================
# 5. DASHBOARD LOGIC
# ==========================================
if st.session_state['logged_in']:
    st.markdown(f"""<div style="display: flex; justify-content: space-between;"><div><h1>📊 AP Analyzing Suite</h1><p>HFO Audit & Payables Intelligence</p></div><div style="text-align: right; color: #94a3b8; font-size: 12px;">Developed by <b>Can Adiguzel</b><br>with Gemini AI technologies</div></div>""", unsafe_allow_html=True)

    if uploaded_file:
        if st.button("🚀 Start Analysis", type="primary"):
            with st.status("🔄 Processing..."):
                df_raw = pd.read_excel(uploaded_file)
                df = clean_sap_data(df_raw)
                df['Amount'] = pd.to_numeric(df['Amount in local currency'], errors='coerce').fillna(0)
                df['G/L Account'] = df['G/L Account'].astype(str).str.split('.').str[0]
                df['Vendor name'] = df['Vendor name'].fillna(df['Supplier'].astype(str))
                report_date = pd.to_datetime(df['Posting Date']).max()
                buckets = ["Not Due", "1-30 Days", "31-60 Days", "61-90 Days", "90+ Days"]
                df['Aging Bucket'] = pd.to_datetime(df['Payment date']).apply(lambda x: "Not Due" if pd.isna(x) or (report_date - x).days < 0 else ("1-30 Days" if (report_date - x).days <= 30 else ("31-60 Days" if (report_date - x).days <= 60 else ("61-90 Days" if (report_date - x).days <= 90 else "90+ Days"))))

                # PIVOTS
                gl_pivot = df.pivot_table(index='G/L Account', columns='Aging Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
                gl_pivot['Total Balance'] = gl_pivot.sum(axis=1)

                df_rec, gl_name_map, gl_solar_map, rec_msg = (None, {}, {}, "")
                if tb_file:
                    df_rec, gl_name_map, gl_solar_map, rec_msg = process_tb_file(tb_file, gl_pivot)

                # GL SUMMARY FINAL
                gl_final = gl_pivot.reset_index()
                gl_final['GL Name'] = gl_final['G/L Account'].map(gl_name_map).fillna("-")
                gl_final['SOLAR Code'] = gl_final['G/L Account'].map(gl_solar_map).fillna("-")
                gl_final = gl_final[['G/L Account', 'GL Name', 'SOLAR Code'] + buckets + ['Total Balance']].sort_values('Total Balance', key=abs, ascending=False)

                # Payables & Debit Balances (k view)
                v_raw = df.pivot_table(index='Vendor name', columns='Aging Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
                v_raw['Total Balance'] = v_raw.sum(axis=1)
                v_ap = v_raw[v_raw['Total Balance'] < 0].sort_values('Total Balance').head(20) / 1000
                v_db = v_raw[v_raw['Total Balance'] > 0].sort_values('Total Balance', ascending=False).head(20) / 1000

                # HTML Report Data
                titles = ["0. GL Reconciliation (TB vs FBL1n)", "1. GL Aging Summary (k)", "2. Top Payables (k)", "3. Debit Balances (k)"]
                dfs = [df_rec, gl_final, v_ap.reset_index(), v_db.reset_index()]
                html_report = generate_html_report(dfs, titles, selected_currency, eur_rate)

            # --- DISPLAY ---
            col_d1, col_d2 = st.columns(2)
            with col_d1: st.download_button("📄 Download Smart Report (HTML/Outlook)", html_report, f"Smart_Report_{datetime.now().strftime('%Y%m%d')}.html", "text/html", use_container_width=True)
            
            # 0. RECONCILIATION
            if df_rec is not None:
                st.markdown("### 0. GL Reconciliation (TB vs FBL1n)")
                st.dataframe(df_rec.style.format("{:,.2f}", subset=['TB_Balance', 'FBL1n_Sum', 'Difference']).applymap(lambda v: 'color: red;' if abs(v) > 1 else 'color: green;', subset=['Difference']), use_container_width=True)

            # 1. GL Summary
            st.markdown("### 1. GL & SOLAR Aging Summary (k)")
            gl_disp = gl_final.copy()
            gl_disp[buckets + ['Total Balance']] = gl_disp[buckets + ['Total Balance']] / 1000
            st.dataframe(gl_disp.style.format("{:,.0f}", subset=buckets+['Total Balance']), use_container_width=True)

            # 2 & 3
            c1, c2 = st.columns(2)
            with c1: st.markdown("### 2. Top Payables (k)"); st.dataframe(v_ap.style.format("{:,.0f}"), use_container_width=True)
            with c2: st.markdown("### 3. Top Debit Balances (k)"); st.dataframe(v_db.style.format("{:,.0f}"), use_container_width=True)

            # 4. HFO Audit Check
            if df_rec is not None:
                st.divider()
                st.markdown("### 🛡️ HFO Audit Check (SOLAR 40000 - Not in FBL1n)")
                hfo_check = df_rec[(df_rec['SOLAR Code'] == '40000') & (df_rec['FBL1n_Sum'] == 0)].copy()
                if not hfo_check.empty:
                    st.dataframe(hfo_check[['GL_Account', 'GL Name', 'TB_Balance']].style.format("{:,.2f}"), use_container_width=True)
                else: st.success("Tüm SOLAR 40000 kalemleri FBL1n ile uyumlu.")

            st.markdown(f"""<div style="position: fixed; bottom: 0; left: 0; width: 100%; background: #f8fafc; text-align: center; padding: 5px; font-size: 11px; color: #94a3b8; border-top: 1px solid #e2e8f0; z-index: 1000;">AP Analyzing Suite {VERSION_NO} | Developed by Can Adiguzel with Gemini AI</div>""", unsafe_allow_html=True)
