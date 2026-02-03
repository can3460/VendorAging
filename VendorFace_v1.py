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
VERSION_NO = "v23.0" 

if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if 'view_currency' not in st.session_state: st.session_state['view_currency'] = "Local"

# ==========================================
# 2. DATA PROCESSING HELPERS
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
        # Sütunları daha geniş kriterlerle tespit et
        acc_col = next((c for c in df_tb.columns if any(x in str(c).lower() for x in ['account', 'g/l', 'acc.no'])), None)
        amt_col = next((c for c in df_tb.columns if any(x in str(c).lower() for x in ['total', 'balance', 'reporting'])), None)
        solar_col = next((c for c in df_tb.columns if any(x in str(c).lower() for x in ['financial', 'fs item', 'solar', 'item'])), None)
        # GL Name için 'Text' içeren ama 'Account' içermeyen ilk sütunu ara
        name_col = next((c for c in df_tb.columns if any(x in str(c).lower() for x in ['text', 'description', 'name']) and 'account' not in str(c).lower()), None)

        if not acc_col or not amt_col: return None, {}, {}, "Required columns missing in TB."

        df_tb = df_tb.dropna(subset=[acc_col])
        df_tb['GL_Key'] = df_tb[acc_col].astype(str).str.strip().str.split('.').str[0]
        
        gl_name_map = df_tb.set_index('GL_Key')[name_col].to_dict() if name_col else {}
        gl_solar_map = df_tb.set_index('GL_Key')[solar_col].to_dict() if solar_col else {}

        tb_summary = df_tb.groupby('GL_Key')[amt_col].sum().reset_index().rename(columns={amt_col: 'TB_Balance'})
        fbl1n_check = fbl1n_gl_summary.reset_index().rename(columns={'G/L Account': 'GL_Key', 'Total Balance': 'FBL1n_Sum'})
        
        merged = pd.merge(fbl1n_check[['GL_Key', 'FBL1n_Sum']], tb_summary, on='GL_Key', how='left').fillna(0)
        merged['Difference'] = merged['TB_Balance'] - merged['FBL1n_Sum']
        merged['SOLAR Code'] = merged['GL_Key'].map(gl_solar_map).fillna("-")
        merged['GL Name'] = merged['GL_Key'].map(gl_name_map).fillna("-")
        
        return merged, gl_name_map, gl_solar_map, "Success"
    except Exception as e: return None, {}, {}, str(e)

def generate_html_report(dfs, titles, display_curr, rate):
    html = f"<html><head><style>body{{font-family:Segoe UI,sans-serif;padding:20px;}}h2{{color:#1e40af;border-bottom:2px solid #5b21b6;}}table{{border-collapse:collapse;width:100%;margin-bottom:20px;font-size:11px;}}th{{background:#f1f5f9;padding:8px;border:1px solid #cbd5e1;}}td{{padding:6px;border:1px solid #cbd5e1;text-align:right;}}td:first-child{{text-align:left;font-weight:bold;}}</style></head><body>"
    html += f"<h1>AP Analyzing Suite Report ({display_curr})</h1><p>Date: {datetime.now().strftime('%Y-%m-%d')} | FX Rate: {rate:,.4f}</p>"
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
# 3. LOGIN & SIDEBAR
# ==========================================
if not st.session_state['logged_in']:
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        st.markdown("<br><br><h2 style='text-align: center;'>AP Analyzing Suite</h2>", unsafe_allow_html=True)
        with st.form("login"):
            email = st.text_input("Email").strip().lower()
            if st.form_submit_button("Login", use_container_width=True):
                st.session_state.update({'logged_in': True, 'user_name': email.split('@')[0].replace('.',' ').title()})
                st.rerun()
    st.stop()

with st.sidebar:
    if os.path.exists("logo.png"): st.image("logo.png", use_container_width=True)
    st.markdown(f"👤 **{st.session_state['user_name']}**")
    st.divider()
    uploaded_file = st.file_uploader("1. FBL1N Report", type=["xlsx", "xls"])
    tb_file = st.file_uploader("2. Trial Balance F.01", type=["xlsx", "xls"])
    currency = st.selectbox("Local Currency", ["EGP", "TRY", "USD", "TND"], index=0)
    if 'cur_val' not in st.session_state: st.session_state['cur_val'] = 52.50
    if st.button("🌐 Sync Online Rate"):
        live = get_live_rate(currency)
        if live: st.session_state['cur_val'] = live; st.toast("Rate Updated!")
    eur_rate = st.number_input(f"EUR/{currency}", value=st.session_state['cur_val'], format="%.4f")
    if st.button("🔒 Logout"): st.session_state['logged_in'] = False; st.rerun()

# ==========================================
# 4. DASHBOARD LOGIC
# ==========================================
st.markdown(f"""<div style="display: flex; justify-content: space-between; align-items: center;"><div><h1 style="margin:0;">📊 AP Analyzing Suite</h1><p style="color:#64748b; margin:0;">HFO Operational Audit Dashboard</p></div><div style="text-align: right; color: #94a3b8; font-size: 11px;">Dev by <b>Can Adiguzel</b><br>{VERSION_NO} | Gemini AI</div></div>""", unsafe_allow_html=True)

col_t1, col_t2 = st.columns([8, 2])
with col_t2:
    toggle_label = "Switch to kEUR" if st.session_state['view_currency'] == "Local" else f"Switch to k{currency}"
    if st.button(f"🔄 {toggle_label}", use_container_width=True):
        st.session_state['view_currency'] = "EUR" if st.session_state['view_currency'] == "Local" else "Local"
        st.rerun()

display_unit = f"k{currency}" if st.session_state['view_currency'] == "Local" else "kEUR"
scalar = 1000 if st.session_state['view_currency'] == "Local" else (1000 * eur_rate)

if uploaded_file:
    if st.button("🚀 Execute Audit", type="primary"):
        with st.status("🔄 Processing..."):
            df_raw = pd.read_excel(uploaded_file)
            df = df_raw.dropna(subset=['Document Number']) if 'Document Number' in df_raw.columns else df_raw
            df['Amount'] = pd.to_numeric(df['Amount in local currency'], errors='coerce').fillna(0)
            df['GL'] = df['G/L Account'].astype(str).str.split('.').str[0]
            df['Vendor'] = df['Vendor name'].fillna(df['Supplier'].astype(str))
            
            report_date = pd.to_datetime(df['Posting Date']).max()
            buckets = ["Not Due", "1-30 Days", "31-60 Days", "61-90 Days", "90+ Days"]
            df['Bucket'] = pd.to_datetime(df['Payment date']).apply(lambda x: "Not Due" if pd.isna(x) or (report_date - x).days < 0 else ("1-30 Days" if (report_date - x).days <= 30 else ("31-60 Days" if (report_date - x).days <= 60 else ("61-90 Days" if (report_date - x).days <= 90 else "90+ Days"))))

            gl_pivot = df.pivot_table(index='GL', columns='Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
            gl_pivot['Total Balance'] = gl_pivot.sum(axis=1)

            df_rec, gl_name_map, gl_solar_map, rec_msg = process_tb_file(tb_file, gl_pivot) if tb_file else (None, {}, {}, "")

            def get_main_driver(sub_df):
                if sub_df.empty: return "-"
                return sub_df.groupby('Vendor')['Amount'].sum().abs().idxmax()
            drivers_map = df.groupby('GL').apply(get_main_driver).to_dict()
            
            gl_final = gl_pivot.reset_index()
            gl_final['GL Name'] = gl_final['GL'].map(gl_name_map).fillna("-")
            gl_final['SOLAR Code'] = gl_final['GL'].map(gl_solar_map).fillna("-")
            gl_final['Main Driver Vendor'] = gl_final['GL'].map(drivers_map).fillna("-")
            gl_final = gl_final[['GL', 'GL Name', 'SOLAR Code', 'Main Driver Vendor'] + buckets + ['Total Balance']].sort_values('Total Balance', key=abs, ascending=False)

            v_raw = df.pivot_table(index='Vendor', columns='Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
            v_raw['Total Balance'] = v_raw.sum(axis=1)
            v_ap = v_raw[v_raw['Total Balance'] < 0].sort_values('Total Balance').head(20)
            v_db = v_raw[v_raw['Total Balance'] > 0].sort_values('Total Balance', ascending=False).head(20)

            # TB Gap Analysis
            hfo_gap = pd.DataFrame()
            if tb_file:
                full_tb = pd.read_excel(tb_file)
                acc_idx = next((i for i,c in enumerate(full_tb.columns) if 'account' in str(c).lower() and 'number' in str(c).lower()), 3)
                solar_idx = next((i for i,c in enumerate(full_tb.columns) if 'financial' in str(c).lower()), 1)
                amt_idx = next((i for i,c in enumerate(full_tb.columns) if 'total' in str(c).lower() and 'reporting' in str(c).lower()), 7)
                
                gap_list = []
                for _, row in full_tb.iterrows():
                    if str(row.iloc[solar_idx]) == '40000':
                        g_acc = str(row.iloc[acc_idx]).strip().split('.')[0]
                        if g_acc.isdigit() and g_acc not in gl_pivot.index:
                            gap_list.append({"GL": g_acc, "GL Name": gl_name_map.get(g_acc, "-"), "SOLAR Code": "40000", "TB Balance": row.iloc[amt_idx]})
                if gap_list: hfo_gap = pd.DataFrame(gap_list).sort_values('TB Balance', ascending=True)

            # Scaled Display Data
            def scale_it(df_in, cols):
                if df_in is None or df_in.empty: return pd.DataFrame()
                df_out = df_in.copy()
                df_out[cols] = df_out[cols] / scalar
                return df_out

            gl_disp = scale_it(gl_final, buckets + ['Total Balance'])
            ap_disp = scale_it(v_ap.reset_index(), buckets + ['Total Balance'])
            db_disp = scale_it(v_db.reset_index(), buckets + ['Total Balance'])
            gap_disp = scale_it(hfo_gap, ['TB Balance'])
            rec_disp = scale_it(df_rec, ['TB_Balance', 'FBL1n_Sum', 'Difference']) if df_rec is not None else None

            html_report = generate_html_report([rec_disp, gl_disp, ap_disp, db_disp, gap_disp], [f"{t} ({display_unit})" for t in ["0. Reconciliation", "1. GL Aging", "2. Top Payables", "3. Top Debit Balances", "🛡️ TB Check"]], display_unit, eur_rate)

        st.download_button(f"📄 Download Report ({display_unit})", html_report, f"Report_{display_unit}.html", "text/html", use_container_width=True)

        if rec_disp is not None and not rec_disp.empty:
            st.markdown(f"### 0. GL Reconciliation ({display_unit})")
            st.dataframe(rec_disp.style.format("{:,.0f}").applymap(lambda v: 'color:red;' if abs(v)>1 else 'color:green;', subset=['Difference']), use_container_width=True)

        st.markdown(f"### 1. GL & SOLAR Aging Summary ({display_unit})")
        st.dataframe(gl_disp.style.format("{:,.0f}", subset=buckets + ['Total Balance']), use_container_width=True)

        c1, c2 = st.columns(2)
        with c1: 
            st.markdown(f"### 2. Top Payables ({display_unit})")
            if not ap_disp.empty: st.dataframe(ap_disp.style.format("{:,.0f}"), use_container_width=True)
            else: st.info("No payables data found.")
        with c2: 
            st.markdown(f"### 3. Top Debit Balances ({display_unit})")
            if not db_disp.empty: st.dataframe(db_disp.style.format("{:,.0f}"), use_container_width=True)
            else: st.info("No debit balance data found.")

        if not gap_disp.empty:
            st.divider()
            st.markdown(f"### 🛡️ TB Check : Other Payables (Not Reported in FBL1n) ({display_unit})")
            st.dataframe(gap_disp.style.format("{:,.0f}"), use_container_width=True)

        st.markdown(f"""<div style="position:fixed;bottom:0;left:0;width:100%;background:#f8fafc;text-align:center;padding:5px;font-size:11px;color:#94a3b8;border-top:1px solid #e2e8f0;z-index:1000;">AP Analyzing Suite | {display_unit} View | {st.session_state['user_name']}</div>""", unsafe_allow_html=True)
