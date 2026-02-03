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
VERSION_NO = "v31.0" 

if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if 'view_currency' not in st.session_state: st.session_state['view_currency'] = "Local"
if 'results' not in st.session_state: st.session_state['results'] = None

# ==========================================
# 2. DATA ENGINE (v26.1 STABLE ENGINE)
# ==========================================

def get_live_rate(base_currency):
    if base_currency == "EUR": return 1.0
    try:
        ticker = yf.Ticker(f"EUR{base_currency}=X")
        history = ticker.history(period="1d")
        return history['Close'].iloc[-1] if not history.empty else None
    except: return None

def smart_parse_tb(file):
    try:
        df_tb = pd.read_excel(file)
        gl_name_map, gl_solar_map, gl_balance_map = {}, {}, {}
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

def generate_html_report(dfs, titles, display_curr, rate):
    html = f"<html><head><style>body{{font-family:sans-serif;padding:20px;}}h2{{color:#1e40af;border-bottom:2px solid #5b21b6;}}table{{border-collapse:collapse;width:100%;margin-bottom:20px;font-size:11px;}}th{{background:#f1f5f9;padding:8px;border:1px solid #cbd5e1;}}td{{padding:6px;border:1px solid #cbd5e1;text-align:right;}}td:first-child{{text-align:left;font-weight:bold;}}</style></head><body>"
    html += f"<h1>AP Analyzing Suite Audit Report ({display_curr})</h1><p>Date: {datetime.now().strftime('%Y-%m-%d %H:%M')} | FX Rate: {rate:,.4f}</p>"
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
# 3. INTERFACE
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
    tb_file = st.file_uploader("2. Trial Balance F.01 (Optional)", type=["xlsx", "xls"])
    currency = st.selectbox("Base Currency", ["EGP", "TRY", "USD", "TND"], index=1)
    if 'cur_val' not in st.session_state: st.session_state['cur_val'] = 52.50
    if st.button("🌐 Sync Online EUR Rate"):
        live = get_live_rate(currency)
        if live: st.session_state['cur_val'] = live; st.toast("Synced!")
    eur_rate = st.number_input(f"EUR/{currency}", value=st.session_state['cur_val'], format="%.4f")
    if st.button("🔒 Logout"): st.session_state['logged_in'] = False; st.rerun()

st.markdown(f"""<div style="display: flex; justify-content: space-between; align-items: center;"><div><h1 style="margin:0; color:#1e3a8a;">📊 AP Analyzing Suite</h1><p style="color:#64748b; margin:0;">HFO Operational Audit Dashboard</p></div><div style="text-align: right; color: #94a3b8; font-size: 11px;">Dev by <b>Can Adiguzel</b><br>{VERSION_NO} | Gemini AI</div></div>""", unsafe_allow_html=True)

col_t1, col_t2 = st.columns([8, 2])
with col_t2:
    toggle_label = "View kEUR" if st.session_state['view_currency'] == "Local" else f"View k{currency}"
    if st.button(f"🔄 {toggle_label}", use_container_width=True):
        st.session_state['view_currency'] = "EUR" if st.session_state['view_currency'] == "Local" else "Local"
        st.rerun()

display_unit = f"k{currency}" if st.session_state['view_currency'] == "Local" else "kEUR"
scalar = 1000 if st.session_state['view_currency'] == "Local" else (1000 * eur_rate)

if uploaded_file:
    if st.button("🚀 Run Audit Analysis", type="primary", use_container_width=True):
        with st.status("🔄 Harmonizing Data..."):
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
            
            gl_final = gl_pivot.reset_index()
            gl_final['GL Name'] = gl_final['GL'].map(name_map).fillna("-")
            gl_final['SOLAR Code'] = gl_final['GL'].map(solar_map).fillna("-")
            gl_final['Main Driver'] = gl_final['GL'].map(drivers).fillna("-")
            gl_final = gl_final[['GL', 'GL Name', 'SOLAR Code', 'Main Driver'] + buckets + ['Total Balance']].sort_values('Total Balance', key=abs, ascending=False)

            v_raw = df.pivot_table(index='Vendor', columns='Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
            v_raw['Total Balance'] = v_raw.sum(axis=1)
            v_ap = v_raw[v_raw['Total Balance'] < 0].sort_values('Total Balance').head(20).reset_index()
            v_db = v_raw[v_raw['Total Balance'] > 0].sort_values('Total Balance', ascending=False).head(20).reset_index()

            # SAFE Prepayments Logic
            dp_gls = ['16740100', '16740110', '16740000']
            v_dp = df[df['GL'].isin(dp_gls)].pivot_table(index='Vendor', columns='Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
            if not v_dp.empty:
                v_dp['Total Balance'] = v_dp.sum(axis=1)
                v_dp = v_dp.sort_values('Total Balance', ascending=False).reset_index()

            rec_list, gap_list = [], []
            if tb_file:
                for gl in gl_pivot.index:
                    tb_v, fbl_v = tb_bal_map.get(gl, 0), gl_pivot.loc[gl, 'Total Balance']
                    rec_list.append({"GL": gl, "GL Name": name_map.get(gl, "-"), "SOLAR Code": solar_map.get(gl, "-"), "TB_Balance": tb_v, "FBL1n_Sum": fbl_v, "Difference": tb_v - fbl_v})
                for gl, s_c in solar_map.items():
                    if str(s_c).strip() == '40000' and gl not in gl_pivot.index:
                        gap_list.append({"GL": gl, "GL Name": name_map.get(gl, "-"), "SOLAR Code": "40000", "TB Balance": tb_bal_map.get(gl, 0)})
            
            st.session_state['results'] = {
                'gl_final': gl_final, 'v_ap': v_ap, 'v_db': v_db, 
                'v_dp': v_dp if not v_dp.empty else pd.DataFrame(),
                'rec_df': pd.DataFrame(rec_list) if rec_list else pd.DataFrame(),
                'gap_df': pd.DataFrame(gap_list) if gap_list else pd.DataFrame(),
                'buckets': buckets, 'raw_data': df
            }

if st.session_state['results']:
    res = st.session_state['results']
    def scale_clean(df_in, cols):
        if df_in is None or df_in.empty: return pd.DataFrame()
        df_out = df_in.copy()
        for c in cols:
            if c in df_out.columns:
                df_out[c] = pd.to_numeric(df_out[c], errors='coerce').fillna(0) / scalar
        return df_out

    gl_disp = scale_clean(res['gl_final'], res['buckets'] + ['Total Balance'])
    ap_disp = scale_clean(res['v_ap'], res['buckets'] + ['Total Balance'])
    db_disp = scale_clean(res['v_db'], res['buckets'] + ['Total Balance'])
    rec_disp = scale_clean(res['rec_df'], ['TB_Balance', 'FBL1n_Sum', 'Difference'])
    gap_disp = scale_clean(res['gap_df'], ['TB Balance'])

    col_e1, col_e2 = st.columns(2)
    with col_e1:
        titles = ["0. Reconciliation", "1. GL Aging Summary", "2. Top Payables", "3. Top Debit Balances", "🛡️ TB Check"]
        html_out = generate_html_report([rec_disp, gl_disp, ap_disp, db_disp, gap_disp], [f"{t} ({display_unit})" for t in titles], display_unit, eur_rate)
        st.download_button("📄 Print Dashboard (HTML Export)", html_out, f"Audit_{display_unit}.html", "text/html", use_container_width=True)
    with col_e2:
        output_dash = io.BytesIO()
        with pd.ExcelWriter(output_dash, engine='xlsxwriter') as writer:
            gl_disp.to_excel(writer, sheet_name='GL Aging', index=False)
            ap_disp.to_excel(writer, sheet_name='Top Payables', index=False)
            db_disp.to_excel(writer, sheet_name='Top Debit Balances', index=False)
            if not rec_disp.empty: rec_disp.to_excel(writer, sheet_name='Reconciliation', index=False)
            if not gap_disp.empty: gap_disp.to_excel(writer, sheet_name='TB Check', index=False)
        st.download_button("📥 Export Dashboard to Excel", output_dash.getvalue(), f"Dash_{display_unit}.xlsx", use_container_width=True)

    if not rec_disp.empty:
        st.markdown(f"### 0. Reconciliation: FBL1n vs TB ({display_unit})")
        try:
            st.dataframe(rec_disp.round(0).style.apply(lambda x: ['color: red' if abs(v) > 1 else 'color: green' for v in x], subset=['Difference']), use_container_width=True)
        except: st.dataframe(rec_disp.round(0), use_container_width=True)

    st.markdown(f"### 1. GL & SOLAR Aging Summary ({display_unit})")
    st.dataframe(gl_disp.round(0), use_container_width=True)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"### 2. Top Payables ({display_unit})")
        if not ap_disp.empty: st.dataframe(ap_disp.round(0), use_container_width=True)
    with c2:
        st.markdown(f"### 3. Top Debit Balances ({display_unit})")
        if not db_disp.empty: st.dataframe(db_disp.round(0), use_container_width=True)

    if not gap_disp.empty:
        st.divider()
        st.markdown(f"### 🛡️ TB Check : Other Payables (Not Reported in FBL1n) ({display_unit})")
        st.dataframe(gap_disp.round(0), use_container_width=True)

    # --- DETAILED REPORT (SAFE WRITER) ---
    st.divider()
    output_full = io.BytesIO()
    with pd.ExcelWriter(output_full, engine='xlsxwriter') as writer:
        res['gl_final'].to_excel(writer, sheet_name='Full GL Aging', index=False)
        res['v_ap'].to_excel(writer, sheet_name='Full Vendor Aging', index=False)
        res['v_db'].to_excel(writer, sheet_name='All Debit Balances', index=False)
        if not res['v_dp'].empty:
            res['v_dp'].to_excel(writer, sheet_name='Prepayments Analysis', index=False)
        if not res['rec_df'].empty:
            res['rec_df'].to_excel(writer, sheet_name='Reconciliation Audit', index=False)
        if not res['gap_df'].empty:
            res['gap_df'].to_excel(writer, sheet_name='Other Payables Audit', index=False)
        res['raw_data'].head(5000).to_excel(writer, sheet_name='Raw Sample Data', index=False)

    st.download_button("📥 Download Detailed Audit Report (Full Data Pack)", output_full.getvalue(), f"Full_Audit_Pack_{datetime.now().strftime('%Y%m%d')}.xlsx", use_container_width=True, type="primary")

    st.markdown(f"""<div style="position:fixed;bottom:0;left:0;width:100%;background:#f8fafc;text-align:center;padding:5px;font-size:11px;color:#94a3b8;border-top:1px solid #e2e8f0;z-index:1000;">AP Analyzing Suite | Dev by Can Adiguzel</div>""", unsafe_allow_html=True)
else:
    st.info("👋 Welcome! Please upload FBL1N and F.01 (Mizan) reports and click 'Run Audit Analysis'.")
