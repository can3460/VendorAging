import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import os
import time

# --- GÜVENLİ IMPORT (Yahoo Finance) ---
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
VERSION_NO = "v17.0" 

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
    return pd.read_excel(USER_DB_FILE)

if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False

# --- LOGIN SCREEN ---
if not st.session_state['logged_in']:
    st.markdown(f"""<div style="position: fixed; top: 15px; right: 80px; background: #dbeafe; color: #1e40af; padding: 5px 15px; border-radius: 20px; font-size: 14px; font-weight: bold; border: 1px solid #bfdbfe; z-index: 9999;">{VERSION_NO}</div>""", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("""<div style='text-align: center;'><h2 style='color:#1e293b; margin-bottom: 0px;'>AP Analyzing Suite</h2></div>""", unsafe_allow_html=True)
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
    """Yahoo Finance EUR Kur Çekici"""
    if base_currency == "EUR": return 1.0
    try:
        ticker = yf.Ticker(f"EUR{base_currency}=X")
        history = ticker.history(period="1d")
        if not history.empty:
            return history['Close'].iloc[-1]
    except:
        return None
    return None

def clean_sap_data(df):
    return df.dropna(subset=['Document Number']) if 'Document Number' in df.columns else df

def process_tb_file(file, fbl1n_gl_summary):
    try:
        df_tb = pd.read_excel(file)
        # Sütun Yakalama
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

        tb_summary = df_tb.groupby('GL_Account')[amt_col].sum().reset_index()
        fbl1n_check = fbl1n_gl_summary.reset_index()[['G/L Account', 'Total Balance']].rename(columns={'G/L Account': 'GL_Account', 'Total Balance': 'FBL1n_Sum'})
        
        # Sadece FBL1n'deki hesapları TB ile mutabakat et
        merged = pd.merge(fbl1n_check, tb_summary, on='GL_Account', how='left').fillna(0)
        merged['Difference'] = merged['TB_Balance'] - merged['FBL1n_Sum']
        merged['SOLAR Code'] = merged['GL_Account'].map(gl_solar_map).fillna("-")
        merged['GL Name'] = merged['GL_Account'].map(gl_name_map).fillna("-")
        
        return merged, gl_name_map, gl_solar_map, "Success"
    except Exception as e: return None, {}, {}, str(e)

def generate_html_report(dfs, titles, currency, rate):
    html = f"""<html><head><style>
    body {{ font-family: sans-serif; padding: 20px; }}
    h2 {{ color: #1e40af; border-bottom: 2px solid #cbd5e1; }}
    table {{ border-collapse: collapse; width: 100%; margin-bottom: 25px; font-size: 11px; }}
    th {{ background: #f8fafc; padding: 10px; border: 1px solid #cbd5e1; text-align: left; }}
    td {{ padding: 8px; border: 1px solid #cbd5e1; text-align: right; }}
    td:first-child {{ text-align: left; }}
    </style></head><body>
    <h1>AP Analyzing Suite - HFO Smart Report</h1>
    <p>Date: {datetime.now().strftime('%Y-%m-%d')} | Currency: {currency} | Rate: {rate}</p>"""
    for df, title in zip(dfs, titles):
        if df is not None and not df.empty:
            html += f"<h2>{title}</h2>"
            df_html = df.copy()
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
    uploaded_file = st.file_uploader("1. FBL1N Report", type=["xlsx", "xls"])
    tb_file = st.file_uploader("2. Trial Balance F.01", type=["xlsx", "xls"])
    selected_currency = st.selectbox("Currency", ["EGP", "TRY", "EUR", "USD", "TND"], index=0)
    
    # --- ONLINE RATE BUTTON ---
    if 'cur_val' not in st.session_state: st.session_state['cur_val'] = 52.50
    
    col_r1, col_r2 = st.columns([3, 2])
    with col_r2:
        st.write("") # Spacer
        if st.button("🌐 Get Rate", help="Fetch live rate from Yahoo Finance"):
            with st.spinner("..."):
                new_rate = get_live_rate(selected_currency)
                if new_rate:
                    st.session_state['cur_val'] = new_rate
                    st.toast(f"Rate updated: {new_rate:,.2f}")
                else: st.error("Error")
    
    with col_r1:
        eur_rate = st.number_input(f"EUR/{selected_currency}", value=st.session_state['cur_val'], format="%.4f")
    
    st.markdown("---")
    if st.button("🔒 Logout"): st.session_state['logged_in'] = False; st.rerun()

# ==========================================
# 5. DASHBOARD LOGIC
# ==========================================
if st.session_state['logged_in']:
    st.markdown(f"""<div style="display: flex; justify-content: space-between;"><div><h1>📊 AP Analyzing Suite</h1><p>HFO Audit & Intelligence</p></div><div style="text-align: right; color: #94a3b8; font-size: 12px;">Developed by <b>Can Adiguzel</b><br>with Gemini AI technologies</div></div>""", unsafe_allow_html=True)

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

                # GL FINAL
                gl_final = gl_pivot.reset_index()
                gl_final['GL Name'] = gl_final['G/L Account'].map(gl_name_map).fillna("-")
                gl_final['SOLAR Code'] = gl_final['G/L Account'].map(gl_solar_map).fillna("-")
                gl_final = gl_final[['G/L Account', 'GL Name', 'SOLAR Code'] + buckets + ['Total Balance']].sort_values('Total Balance', key=abs, ascending=False)

                # Top Payables / Debit (k)
                v_raw = df.pivot_table(index='Vendor name', columns='Aging Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
                v_raw['Total Balance'] = v_raw.sum(axis=1)
                v_ap = v_raw[v_raw['Total Balance'] < 0].sort_values('Total Balance').head(20) / 1000
                v_db = v_raw[v_raw['Total Balance'] > 0].sort_values('Total Balance', ascending=False).head(20) / 1000

                # Audit Gap Analysis
                hfo_gap = pd.DataFrame()
                if tb_file:
                    full_tb = pd.read_excel(tb_file)
                    # Sadece SOLAR 40000 olan ama FBL1n'de hiç olmayan GL'leri bul
                    audit_gls = full_tb[(full_tb.iloc[:, 1].astype(str) == '40000')]
                    gap_list = []
                    for _, row in audit_gls.iterrows():
                        gl = str(row.iloc[3]).strip().split('.')[0]
                        if gl.isdigit() and gl not in gl_pivot.index:
                            gap_list.append({"GL": gl, "Name": gl_name_map.get(gl, "-"), "TB Balance": row.iloc[7]})
                    hfo_gap = pd.DataFrame(gap_list).sort_values('TB Balance', ascending=True)

                titles = ["0. Reconciliation (Matched)", "1. GL Aging (k)", "2. Top Payables (k)", "3. Debit Balances (k)", "🛡️ Audit: SOLAR 40000 Gap"]
                dfs = [df_rec, gl_final, v_ap.reset_index(), v_db.reset_index(), hfo_gap]
                html_report = generate_html_report(dfs, titles, selected_currency, eur_rate)

            # --- DISPLAY ---
            st.download_button("📄 Download Smart Report (HTML)", html_report, f"Report_{datetime.now().strftime('%Y%m%d')}.html", "text/html")
            
            if df_rec is not None:
                st.markdown("### 0. GL Reconciliation (Matched Items)")
                num_cols = ['TB_Balance', 'FBL1n_Sum', 'Difference']
                st.dataframe(df_rec[['GL_Account', 'GL Name', 'SOLAR Code'] + num_cols].style.format("{:,.0f}", subset=num_cols), use_container_width=True)

            st.markdown("### 1. GL & SOLAR Aging Summary (k)")
            st.dataframe(gl_final.style.format("{:,.0f}", subset=buckets + ['Total Balance']), use_container_width=True)

            c1, c2 = st.columns(2)
            with c1: st.markdown("### 2. Top Payables (k)"); st.dataframe(v_ap.style.format("{:,.0f}"), use_container_width=True)
            with c2: st.markdown("### 3. Top Debit Balances (k)"); st.dataframe(v_db.style.format("{:,.0f}"), use_container_width=True)

            if not hfo_gap.empty:
                st.divider()
                st.markdown("### 🛡️ HFO Audit: Accounts NOT in FBL1n")
                st.dataframe(hfo_gap.style.format("{:,.0f}", subset=['TB Balance']), use_container_width=True)

            st.markdown(f"""<div style="position: fixed; bottom: 0; left: 0; width: 100%; background: #f8fafc; text-align: center; padding: 5px; font-size: 11px; color: #94a3b8; border-top: 1px solid #e2e8f0; z-index: 1000;">AP Analyzing Suite {VERSION_NO} | Developed by Can Adiguzel with Gemini AI</div>""", unsafe_allow_html=True)
