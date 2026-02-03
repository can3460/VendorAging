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
VERSION_NO = "v24.0" 

if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if 'view_currency' not in st.session_state: st.session_state['view_currency'] = "Local"

# ==========================================
# 2. DATA ENGINE (SMART PARSING)
# ==========================================

def get_live_rate(base_currency):
    if base_currency == "EUR": return 1.0
    try:
        ticker = yf.Ticker(f"EUR{base_currency}=X")
        history = ticker.history(period="1d")
        return history['Close'].iloc[-1] if not history.empty else None
    except: return None

def smart_parse_tb(file):
    """Mizandaki G/L isimlerini ve SOLAR kodlarını satır bazlı tarayarak yakalar"""
    try:
        df_tb = pd.read_excel(file)
        gl_name_map = {}
        gl_solar_map = {}
        gl_balance_map = {}

        # Sütunları tahmin et
        acc_col_idx = next((i for i, c in enumerate(df_tb.columns) if any(x in str(c).lower() for x in ['account', 'g/l', 'acc.no'])), 3)
        solar_col_idx = next((i for i, c in enumerate(df_tb.columns) if any(x in str(c).lower() for x in ['financial', 'fs item', 'solar'])), 1)
        amt_col_idx = next((i for i, c in enumerate(df_tb.columns) if any(x in str(c).lower() for x in ['total', 'balance', 'reporting'])), 7)
        # Name kolonu genelde Account Number'ın sağındaki ilk metin sütunudur
        name_col_idx = next((i for i, c in enumerate(df_tb.columns) if any(x in str(c).lower() for x in ['text', 'description', 'name']) and i != acc_col_idx), 5)

        for _, row in df_tb.iterrows():
            raw_acc = str(row.iloc[acc_col_idx]).strip().split('.')[0]
            if raw_acc.isdigit() and len(raw_acc) >= 6:
                gl_name_map[raw_acc] = str(row.iloc[name_col_idx]) if not pd.isna(row.iloc[name_col_idx]) else "-"
                gl_solar_map[raw_acc] = str(row.iloc[solar_col_idx]) if not pd.isna(row.iloc[solar_col_idx]) else "-"
                gl_balance_map[raw_acc] = row.iloc[amt_col_idx] if not pd.isna(row.iloc[amt_col_idx]) else 0

        return gl_name_map, gl_solar_map, gl_balance_map
    except:
        return {}, {}, {}

def safe_format_df(styler, numeric_cols):
    """ValueError çökmesini engelleyen güvenli formatlayıcı"""
    if styler.data.empty:
        return styler.data
    try:
        # Sadece mevcut olan sayısal kolonları formatla
        existing_num_cols = [c for c in numeric_cols if c in styler.data.columns]
        return styler.format("{:,.0f}", subset=existing_num_cols)
    except:
        return styler.data

# ==========================================
# 3. INTERFACE (LOGIN & SIDEBAR)
# ==========================================
if not st.session_state['logged_in']:
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        st.markdown("<br><br><h2 style='text-align: center;'>AP Analyzing Suite</h2>", unsafe_allow_html=True)
        with st.form("login"):
            email = st.text_input("Email").strip().lower()
            if st.form_submit_button("Secure Login", use_container_width=True):
                st.session_state.update({'logged_in': True, 'user_name': email.split('@')[0].replace('.',' ').title()})
                st.rerun()
    st.stop()

with st.sidebar:
    st.markdown(f"👤 **{st.session_state['user_name']}**")
    st.divider()
    uploaded_file = st.file_uploader("1. FBL1N Report", type=["xlsx", "xls"])
    tb_file = st.file_uploader("2. Trial Balance F.01", type=["xlsx", "xls"])
    currency = st.selectbox("Company Local Currency", ["EGP", "TRY", "USD", "TND"], index=0)
    if 'cur_val' not in st.session_state: st.session_state['cur_val'] = 52.50
    if st.button("🌐 Sync Online EUR Rate"):
        live = get_live_rate(currency)
        if live: st.session_state['cur_val'] = live; st.toast("Synced!")
    eur_rate = st.number_input(f"EUR/{currency} Rate", value=st.session_state['cur_val'], format="%.4f")
    if st.button("🔒 Logout"): st.session_state['logged_in'] = False; st.rerun()

# ==========================================
# 4. DASHBOARD LOGIC
# ==========================================
st.markdown(f"""<div style="display: flex; justify-content: space-between; align-items: center;"><div><h1 style="margin:0;">📊 AP Analyzing Suite</h1><p style="color:#64748b; margin:0;">HFO Operational Audit Dashboard</p></div><div style="text-align: right; color: #94a3b8; font-size: 11px;">Dev by <b>Can Adiguzel</b><br>{VERSION_NO} | Gemini AI</div></div>""", unsafe_allow_html=True)

# Currency Toggle
col_t1, col_t2 = st.columns([8, 2])
with col_t2:
    toggle_label = "Switch to kEUR" if st.session_state['view_currency'] == "Local" else f"Switch to k{currency}"
    if st.button(f"🔄 {toggle_label}", use_container_width=True):
        st.session_state['view_currency'] = "EUR" if st.session_state['view_currency'] == "Local" else "Local"
        st.rerun()

display_unit = f"k{currency}" if st.session_state['view_currency'] == "Local" else "kEUR"
scalar = 1000 if st.session_state['view_currency'] == "Local" else (1000 * eur_rate)

if uploaded_file:
    if st.button("🚀 Execute Comprehensive Analysis", type="primary"):
        with st.status("🔄 Harmonizing Data..."):
            # Load and Clean FBL1N
            df_raw = pd.read_excel(uploaded_file)
            df = df_raw.dropna(subset=['Document Number']) if 'Document Number' in df_raw.columns else df_raw
            df['Amount'] = pd.to_numeric(df['Amount in local currency'], errors='coerce').fillna(0)
            df['GL'] = df['G/L Account'].astype(str).str.split('.').str[0]
            df['Vendor'] = df['Vendor name'].fillna(df['Supplier'].astype(str))
            
            # Aging
            report_date = pd.to_datetime(df['Posting Date']).max()
            buckets = ["Not Due", "1-30 Days", "31-60 Days", "61-90 Days", "90+ Days"]
            df['Bucket'] = pd.to_datetime(df['Payment date']).apply(lambda x: "Not Due" if pd.isna(x) or (report_date - x).days < 0 else ("1-30 Days" if (report_date - x).days <= 30 else ("31-60 Days" if (report_date - x).days <= 60 else ("61-90 Days" if (report_date - x).days <= 90 else "90+ Days"))))

            gl_pivot = df.pivot_table(index='GL', columns='Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
            gl_pivot['Total Balance'] = gl_pivot.sum(axis=1)

            # Mizan İşleme (Smart Parsing)
            name_map, solar_map, tb_bal_map = smart_parse_tb(tb_file) if tb_file else ({}, {}, {})

            # Main Driver Vendor
            drivers = df.groupby('GL').apply(lambda x: x.groupby('Vendor')['Amount'].sum().abs().idxmax() if not x.empty else "-").to_dict()

            # 1. GL Summary Table
            gl_final = gl_pivot.reset_index()
            gl_final['GL Name'] = gl_final['GL'].map(name_map).fillna("-")
            gl_final['SOLAR Code'] = gl_final['GL'].map(solar_map).fillna("-")
            gl_final['Main Driver'] = gl_final['GL'].map(drivers).fillna("-")
            gl_final = gl_final[['GL', 'GL Name', 'SOLAR Code', 'Main Driver'] + buckets + ['Total Balance']].sort_values('Total Balance', key=abs, ascending=False)

            # 2 & 3. Vendor Tables
            v_raw = df.pivot_table(index='Vendor', columns='Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
            v_raw['Total Balance'] = v_raw.sum(axis=1)
            v_ap = v_raw[v_raw['Total Balance'] < 0].sort_values('Total Balance').head(20).reset_index()
            v_db = v_raw[v_raw['Total Balance'] > 0].sort_values('Total Balance', ascending=False).head(20).reset_index()

            # 4. Reconciliation Table
            rec_df = pd.DataFrame()
            if tb_file:
                rec_list = []
                for gl in gl_pivot.index:
                    tb_val = tb_bal_map.get(gl, 0)
                    fbl_val = gl_pivot.loc[gl, 'Total Balance']
                    rec_list.append({
                        "GL": gl, "GL Name": name_map.get(gl, "-"), "SOLAR Code": solar_map.get(gl, "-"),
                        "TB_Balance": tb_val, "FBL1n_Sum": fbl_val, "Difference": tb_val - fbl_val
                    })
                rec_df = pd.DataFrame(rec_list).sort_values('Difference', key=abs, ascending=False)

            # 5. TB Gap Analysis (Other Payables)
            gap_df = pd.DataFrame()
            if tb_file:
                gap_list = []
                for gl, solar in solar_map.items():
                    if solar == '40000' and gl not in gl_pivot.index:
                        gap_list.append({"GL": gl, "GL Name": name_map.get(gl, "-"), "SOLAR Code": "40000", "TB Balance": tb_bal_map.get(gl, 0)})
                if gap_list: gap_df = pd.DataFrame(gap_list).sort_values('TB Balance', ascending=True)

        # --- RENDERING ---
        
        # Reconciliation
        if not rec_df.empty:
            st.markdown(f"### 0. Reconciliation: FBL1n vs TB ({display_unit})")
            disp_rec = rec_df.copy()
            disp_rec[['TB_Balance', 'FBL1n_Sum', 'Difference']] /= scalar
            st.dataframe(safe_format_df(disp_rec.style, ['TB_Balance', 'FBL1n_Sum', 'Difference']), use_container_width=True)

        # Main Table
        st.markdown(f"### 1. GL & SOLAR Aging Summary ({display_unit})")
        disp_gl = gl_final.copy()
        disp_gl[buckets + ['Total Balance']] /= scalar
        st.dataframe(safe_format_df(disp_gl.style, buckets + ['Total Balance']), use_container_width=True)

        # Side Tables
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f"### 2. Top Payables ({display_unit})")
            if not v_ap.empty:
                disp_ap = v_ap.copy()
                disp_ap[buckets + ['Total Balance']] /= scalar
                st.dataframe(safe_format_df(disp_ap.style, buckets + ['Total Balance']), use_container_width=True)
            else: st.info("No payables found.")
        with c2:
            st.markdown(f"### 3. Top Debit Balances ({display_unit})")
            if not v_db.empty:
                disp_db = v_db.copy()
                disp_db[buckets + ['Total Balance']] /= scalar
                st.dataframe(safe_format_df(disp_db.style, buckets + ['Total Balance']), use_container_width=True)
            else: st.info("No debit balances found.")

        # Gap Analysis
        if not gap_df.empty:
            st.divider()
            st.markdown(f"### 🛡️ TB Check : Other Payables (Not Reported in FBL1n) ({display_unit})")
            disp_gap = gap_df.copy()
            disp_gap['TB Balance'] /= scalar
            st.dataframe(safe_format_df(disp_gap.style, ['TB Balance']), use_container_width=True)

        st.markdown(f"""<div style="position:fixed;bottom:0;left:0;width:100%;background:#f8fafc;text-align:center;padding:5px;font-size:11px;color:#94a3b8;border-top:1px solid #e2e8f0;z-index:1000;">AP Analyzing Suite | {display_unit} View | Developed by Can Adiguzel</div>""", unsafe_allow_html=True)
