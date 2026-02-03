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
VERSION_NO = "v18.0" 

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
    if base_currency == "EUR": return 1.0
    try:
        ticker = yf.Ticker(f"EUR{base_currency}=X")
        history = ticker.history(period="1d")
        return history['Close'].iloc[-1] if not history.empty else None
    except: return None

def process_tb_file(file, fbl1n_gl_summary):
    try:
        df_tb = pd.read_excel(file)
        # Akıllı Sütun Yakalama
        acc_col = next((c for c in df_tb.columns if 'Account' in str(c) and 'Number' in str(c)), None)
        amt_col = next((c for c in df_tb.columns if 'Total' in str(c) and 'reporting' in str(c)), None)
        solar_col = next((c for c in df_tb.columns if 'Financial' in str(c) and 'Item' in str(c)), None)
        name_col = next((c for c in df_tb.columns if 'Text' in str(c)), None)

        if not acc_col or not amt_col: return None, {}, {}, "Missing columns"

        df_tb = df_tb.dropna(subset=[acc_col])
        df_tb['GL_Account'] = df_tb[acc_col].astype(str).str.strip().str.split('.').str[0]
        
        # Mapping dictionaries
        gl_name_map = df_tb.set_index('GL_Account')[name_col].to_dict() if name_col else {}
        gl_solar_map = df_tb.set_index('GL_Account')[solar_col].to_dict() if solar_col else {}

        # Mutabakat (Reconciliation)
        tb_summary = df_tb.groupby('GL_Account')[amt_col].sum().reset_index()
        fbl1n_check = fbl1n_gl_summary.reset_index()[['G/L Account', 'Total Balance']].rename(columns={'G/L Account': 'GL_Account', 'Total Balance': 'FBL1n_Sum'})
        
        merged = pd.merge(fbl1n_check, tb_summary, on='GL_Account', how='left').fillna(0)
        merged['Difference'] = merged['TB_Balance'] - merged['FBL1n_Sum']
        merged['SOLAR Code'] = merged['GL_Account'].map(gl_solar_map).fillna("-")
        merged['GL Name'] = merged['GL_Account'].map(gl_name_map).fillna("-")
        
        return merged, gl_name_map, gl_solar_map, "Success"
    except Exception as e: return None, {}, {}, str(e)

def generate_html_report(dfs, titles, currency, rate):
    html = f"<html><head><style>body {{ font-family: sans-serif; }} table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; font-size: 11px; }} th {{ background: #f1f5f9; padding: 8px; border: 1px solid #cbd5e1; }} td {{ padding: 6px; border: 1px solid #cbd5e1; text-align: right; }} td:first-child {{ text-align: left; }}</style></head><body>"
    html += f"<h1>AP Analyzing Suite Report</h1><p>Date: {datetime.now().strftime('%Y-%m-%d')} | Currency: {currency} | Rate: {rate}</p>"
    for df, title in zip(dfs, titles):
        if df is not None and not df.empty:
            html += f"<h3>{title}</h3>"
            df_fmt = df.copy()
            for col in df_fmt.select_dtypes(include=[np.number]).columns:
                df_fmt[col] = df_fmt[col].apply(lambda x: f"{x:,.0f}")
            html += df_fmt.to_html(index=False)
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
    
    if 'cur_val' not in st.session_state: st.session_state['cur_val'] = 52.50
    if st.button("🌐 Get Live Rate"):
        live = get_live_rate(selected_currency)
        if live: st.session_state['cur_val'] = live; st.toast("Updated!")
    
    eur_rate = st.number_input(f"EUR/{selected_currency}", value=st.session_state['cur_val'], format="%.4f")
    if st.button("🔒 Logout"): st.session_state['logged_in'] = False; st.rerun()

# ==========================================
# 5. MAIN DASHBOARD
# ==========================================
if st.session_state['logged_in']:
    st.markdown(f"""<div style="display: flex; justify-content: space-between;"><div><h1>📊 AP Analyzing Suite</h1><p>HFO Audit & Intelligence</p></div><div style="text-align: right; color: #94a3b8; font-size: 12px;">Developed by <b>Can Adiguzel</b><br>with Gemini AI technologies</div></div>""", unsafe_allow_html=True)

    if uploaded_file:
        if st.button("🚀 Run Analysis", type="primary"):
            with st.status("🔄 Processing..."):
                df_raw = pd.read_excel(uploaded_file)
                df = df_raw.dropna(subset=['Document Number']) if 'Document Number' in df_raw.columns else df_raw
                df['Amount'] = pd.to_numeric(df['Amount in local currency'], errors='coerce').fillna(0)
                df['GL'] = df['G/L Account'].astype(str).str.split('.').str[0]
                df['Vendor'] = df['Vendor name'].fillna(df['Supplier'].astype(str))
                
                # Aging
                report_date = pd.to_datetime(df['Posting Date']).max()
                buckets = ["Not Due", "1-30 Days", "31-60 Days", "61-90 Days", "90+ Days"]
                df['Bucket'] = pd.to_datetime(df['Payment date']).apply(lambda x: "Not Due" if pd.isna(x) or (report_date - x).days < 0 else ("1-30 Days" if (report_date - x).days <= 30 else ("31-60 Days" if (report_date - x).days <= 60 else ("61-90 Days" if (report_date - x).days <= 90 else "90+ Days"))))

                # PIVOTS
                gl_pivot = df.pivot_table(index='GL', columns='Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
                gl_pivot['Total Balance'] = gl_pivot.sum(axis=1)

                # TB Processing
                df_rec, gl_name_map, gl_solar_map, rec_msg = (None, {}, {}, "")
                if tb_file:
                    df_rec, gl_name_map, gl_solar_map, rec_msg = process_tb_file(tb_file, gl_pivot)

                # GL SUMMARY + MAIN DRIVER
                def get_top_driver(sub_df):
                    if sub_df.empty: return "-"
                    return sub_df.groupby('Vendor')['Amount'].sum().abs().idxmax()
                
                drivers = df.groupby('GL').apply(get_top_driver).to_dict()
                
                gl_final = gl_pivot.reset_index()
                gl_final['GL Name'] = gl_final['GL'].map(gl_name_map).fillna("-")
                gl_final['SOLAR Code'] = gl_final['GL'].map(gl_solar_map).fillna("-")
                gl_final['Main Driver Vendor'] = gl_final['GL'].map(drivers).fillna("-")
                
                gl_final = gl_final[['GL', 'GL Name', 'SOLAR Code', 'Main Driver Vendor'] + buckets + ['Total Balance']].sort_values('Total Balance', key=abs, ascending=False)

                # Audit Gap (TB but not FBL1n)
                hfo_gap = pd.DataFrame()
                if tb_file:
                    full_tb = pd.read_excel(tb_file)
                    acc_col_idx = 3 # Standart SAP TB structure
                    solar_col_idx = 1
                    amt_col_idx = 7
                    
                    gap_data = []
                    # Sadece SOLAR 40000 olanları tara
                    for _, row in full_tb.iterrows():
                        if str(row.iloc[solar_col_idx]) == '40000':
                            gl_acc = str(row.iloc[acc_col_idx]).strip().split('.')[0]
                            if gl_acc.isdigit() and gl_acc not in gl_pivot.index:
                                gap_data.append({
                                    "GL": gl_acc,
                                    "GL Name": gl_name_map.get(gl_acc, "-"),
                                    "SOLAR Code": "40000",
                                    "TB Balance": row.iloc[amt_col_idx]
                                })
                    hfo_gap = pd.DataFrame(gap_data).sort_values('TB Balance', ascending=True)

                # Prepare HTML Report
                titles = ["0. Reconciliation (FBL1n vs TB)", "1. GL & SOLAR Aging Summary", "🛡️ TB Check : Other Payables (Not in FBL1n)"]
                html_dfs = [df_rec, gl_final, hfo_gap]
                html_report = generate_html_report(html_dfs, titles, selected_currency, eur_rate)

            # --- RENDER ---
            st.download_button("📄 Download Smart Report (Outlook/PDF)", html_report, f"Report_{datetime.now().strftime('%Y%m%d')}.html", "text/html")

            if df_rec is not None:
                st.markdown("### 0. Reconciliation (FBL1n vs TB Match)")
                st.dataframe(df_rec[['GL_Account', 'GL Name', 'SOLAR Code', 'TB_Balance', 'FBL1n_Sum', 'Difference']].style.format("{:,.0f}", subset=['TB_Balance', 'FBL1n_Sum', 'Difference']), use_container_width=True)

            st.markdown("### 1. GL & SOLAR Aging Summary (k)")
            gl_disp = gl_final.copy()
            gl_disp[buckets + ['Total Balance']] = gl_disp[buckets + ['Total Balance']] / 1000
            st.dataframe(gl_disp.style.format("{:,.0f}", subset=buckets + ['Total Balance']), use_container_width=True)

            if not hfo_gap.empty:
                st.divider()
                st.markdown("### 🛡️ TB Check : Other Payables (Not Reported in FBL1n Transaction)")
                st.dataframe(hfo_gap.style.format("{:,.0f}", subset=['TB Balance']), use_container_width=True)

            st.markdown(f"""<div style="position: fixed; bottom: 0; left: 0; width: 100%; background: #f8fafc; text-align: center; padding: 5px; font-size: 11px; color: #94a3b8; border-top: 1px solid #e2e8f0; z-index: 1000;">AP Analyzing Suite {VERSION_NO} | Developed by Can Adiguzel with Gemini AI</div>""", unsafe_allow_html=True)
