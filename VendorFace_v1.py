import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import os

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
VERSION_NO = "v27.0" 

if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if 'view_currency' not in st.session_state: st.session_state['view_currency'] = "Local"
if 'last_analysis' not in st.session_state: st.session_state['last_analysis'] = None

# ==========================================
# 2. DATA ENGINE (PRECISION MAPPING)
# ==========================================

def get_live_rate(base_currency):
    if base_currency == "EUR": return 1.0
    try:
        ticker = yf.Ticker(f"EUR{base_currency}=X")
        history = ticker.history(period="1d")
        return history['Close'].iloc[-1] if not history.empty else None
    except: return None

def smart_parse_tb(file):
    """Account Number bazlı kesin eşleşme ve Text for B/S P&L item çekimi"""
    try:
        df_tb = pd.read_excel(file)
        gl_name_map, gl_solar_map, gl_balance_map = {}, {}, {}

        # Sütun Yakalama
        acc_col = next((c for c in df_tb.columns if 'Account Number' in str(c)), None)
        name_col = next((c for c in df_tb.columns if 'Text for B/S P&L item' in str(c)), None)
        solar_col = next((c for c in df_tb.columns if any(x in str(c).lower() for x in ['financial', 'fs item', 'solar'])), None)
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
    except: return {}, {}, {}

def safe_format_df(styler, numeric_cols):
    if styler.data.empty: return styler.data
    existing = [c for c in numeric_cols if c in styler.data.columns]
    return styler.format("{:,.0f}", subset=existing) if existing else styler.data

# ==========================================
# 3. AUTH & SIDEBAR
# ==========================================
if not st.session_state['logged_in']:
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        st.markdown("<br><br><h2 style='text-align: center; color:#1e3a8a;'>AP Analyzing Suite</h2>", unsafe_allow_html=True)
        with st.form("login"):
            email = st.text_input("Corporate Email").strip().lower()
            if st.form_submit_button("Login", use_container_width=True):
                st.session_state.update({'logged_in': True, 'user_name': email.split('@')[0].replace('.',' ').title()})
                st.rerun()
    st.stop()

with st.sidebar:
    st.markdown(f"👤 **{st.session_state['user_name']}**")
    st.divider()
    uploaded_file = st.file_uploader("1. FBL1N Report", type=["xlsx", "xls"])
    tb_file = st.file_uploader("2. Trial Balance F.01", type=["xlsx", "xls"])
    currency = st.selectbox("Base Currency", ["EGP", "TRY", "USD", "TND"], index=0)
    if 'cur_val' not in st.session_state: st.session_state['cur_val'] = 52.50
    if st.button("🌐 Sync EUR Rate"):
        live = get_live_rate(currency)
        if live: st.session_state['cur_val'] = live; st.toast("Synced!")
    eur_rate = st.number_input(f"EUR/{currency}", value=st.session_state['cur_val'], format="%.4f")
    if st.button("🔒 Logout"): st.session_state['logged_in'] = False; st.rerun()

# ==========================================
# 4. DASHBOARD ENGINE
# ==========================================
st.markdown(f"""<div style="display: flex; justify-content: space-between; align-items: center;"><div><h1 style="margin:0; color:#1e3a8a;">📊 AP Analyzing Suite</h1><p style="color:#64748b; margin:0;">HFO Operational Audit Dashboard</p></div><div style="text-align: right; color: #94a3b8; font-size: 11px;">Dev by <b>Can Adiguzel</b><br>{VERSION_NO} | Gemini AI</div></div>""", unsafe_allow_html=True)

# Toggle logic
col_t1, col_t2 = st.columns([8, 2])
with col_t2:
    toggle_label = "View kEUR" if st.session_state['view_currency'] == "Local" else f"View k{currency}"
    if st.button(f"🔄 {toggle_label}", use_container_width=True):
        st.session_state['view_currency'] = "EUR" if st.session_state['view_currency'] == "Local" else "Local"
        st.rerun()

display_unit = f"k{currency}" if st.session_state['view_currency'] == "Local" else "kEUR"
scalar = 1000 if st.session_state['view_currency'] == "Local" else (1000 * eur_rate)

# AUTO-ANALYSIS TRIGGER
if uploaded_file:
    # Dosya yüklendiğinde veya Currency değiştiğinde otomatik çalışır
    with st.spinner("Processing..."):
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

        name_map, solar_map, tb_bal_map = smart_parse_tb(tb_file) if tb_file else ({}, {}, {})
        drivers = df.groupby('GL').apply(lambda x: x.groupby('Vendor')['Amount'].sum().abs().idxmax() if not x.empty else "-").to_dict()

        # Aging Table
        gl_final = gl_pivot.reset_index()
        gl_final['GL Name'] = gl_final['GL'].map(name_map).fillna("-")
        gl_final['SOLAR Code'] = gl_final['GL'].map(solar_map).fillna("-")
        gl_final['Main Driver'] = gl_final['GL'].map(drivers).fillna("-")
        gl_final = gl_final[['GL', 'GL Name', 'SOLAR Code', 'Main Driver'] + buckets + ['Total Balance']].sort_values('Total Balance', key=abs, ascending=False)

        # Vendors
        v_raw = df.pivot_table(index='Vendor', columns='Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
        v_raw['Total Balance'] = v_raw.sum(axis=1)
        v_ap = v_raw[v_raw['Total Balance'] < 0].sort_values('Total Balance').head(20).reset_index()
        v_db = v_raw[v_raw['Total Balance'] > 0].sort_values('Total Balance', ascending=False).head(20).reset_index()

        # Recon & Gap
        rec_list, gap_list = [], []
        if tb_file:
            for gl in gl_pivot.index:
                tb_v, fbl_v = tb_bal_map.get(gl, 0), gl_pivot.loc[gl, 'Total Balance']
                rec_list.append({"GL": gl, "GL Name": name_map.get(gl, "-"), "SOLAR Code": solar_map.get(gl, "-"), "TB_Balance": tb_v, "FBL1n_Sum": fbl_v, "Difference": tb_v - fbl_v})
            for gl, s_c in solar_map.items():
                if str(s_c).strip() == '40000' and gl not in gl_pivot.index:
                    gap_list.append({"GL": gl, "GL Name": name_map.get(gl, "-"), "SOLAR Code": "40000", "TB Balance": tb_bal_map.get(gl, 0)})

        # Scaled Views
        def scale(df_in, cols):
            if df_in is None or (isinstance(df_in, pd.DataFrame) and df_in.empty): return pd.DataFrame()
            df_out = df_in.copy()
            df_out[cols] = df_out[cols] / scalar
            return df_out

        gl_disp = scale(gl_final, buckets + ['Total Balance'])
        ap_disp = scale(v_ap, buckets + ['Total Balance'])
        db_disp = scale(v_db, buckets + ['Total Balance'])
        rec_disp = scale(pd.DataFrame(rec_list), ['TB_Balance', 'FBL1n_Sum', 'Difference']) if rec_list else None
        gap_disp = scale(pd.DataFrame(gap_list), ['TB Balance']) if gap_list else pd.DataFrame()

    # --- UI RENDER ---
    if rec_disp is not None:
        st.markdown(f"### 0. Reconciliation: FBL1n vs TB ({display_unit})")
        st.dataframe(safe_format_df(rec_disp.style, ['TB_Balance', 'FBL1n_Sum', 'Difference']), use_container_width=True)

    st.markdown(f"### 1. GL & SOLAR Aging Summary ({display_unit})")
    st.dataframe(safe_format_df(gl_disp.style, buckets + ['Total Balance']), use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"### 2. Top Payables ({display_unit})")
        if not ap_disp.empty: st.dataframe(safe_format_df(ap_disp.style, buckets + ['Total Balance']), use_container_width=True)
    with c2:
        st.markdown(f"### 3. Top Debit Balances ({display_unit})")
        if not db_disp.empty: st.dataframe(safe_format_df(db_disp.style, buckets + ['Total Balance']), use_container_width=True)

    if not gap_disp.empty:
        st.divider()
        st.markdown(f"### 🛡️ TB Check : Other Payables (Not Reported in FBL1n) ({display_unit})")
        st.dataframe(safe_format_df(gap_disp.style, ['TB Balance']), use_container_width=True)

    st.markdown(f"""<div style="position:fixed;bottom:0;left:0;width:100%;background:#f8fafc;text-align:center;padding:5px;font-size:11px;color:#94a3b8;border-top:1px solid #e2e8f0;z-index:1000;">AP Analyzing Suite | {display_unit} View | {datetime.now().strftime('%H:%M')}</div>""", unsafe_allow_html=True)
else:
    st.info("👋 Welcome! Please upload FBL1N report to start.")
