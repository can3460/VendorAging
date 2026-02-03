import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import os
import time

# --- GÜVENLİ IMPORT (Kütüphane yoksa patlamasın) ---
try:
    import yfinance as yf
    YFINANCE_AVAILABLE = True
except ImportError:
    YFINANCE_AVAILABLE = False

# ==========================================
# 1. CONFIGURATION & SETUP
# ==========================================
st.set_page_config(page_title="Vendor Analysis Tool | Opella Finance", layout="wide", page_icon="🛡️")
USER_DB_FILE = "users.xlsx"
ADMIN_EMAIL = "can.adiguzel@sanofi.com" 

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
            {"Email": "Cedric.Fallu@sanofi.com", "Name": "Cedric Fallu", "Role": "User"}
        ]
        df = pd.DataFrame(initial_users)
        df['Email'] = df['Email'].str.lower().str.strip()
        df['Added_Date'] = datetime.now().strftime("%Y-%m-%d")
        df.to_excel(USER_DB_FILE, index=False)
        return df
    else:
        return pd.read_excel(USER_DB_FILE)

def add_user_to_db(email, name):
    df = load_user_db()
    email = email.lower().strip()
    if email in df['Email'].values:
        return False, "User already exists!"
    new_user = pd.DataFrame({
        "Email": [email], "Name": [name], "Role": ["User"],
        "Added_Date": [datetime.now().strftime("%Y-%m-%d")]
    })
    df = pd.concat([df, new_user], ignore_index=True)
    try:
        df.to_excel(USER_DB_FILE, index=False)
        return True, "User authorized successfully."
    except Exception as e:
        return False, f"Error saving DB: {e}"

if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False

# --- LOGIN SCREEN ---
if not st.session_state['logged_in']:
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        if os.path.exists("logo.png"):
            st.image("logo.png", width=200)
        else:
            st.markdown("<h1 style='text-align: center; color:#5b21b6;'>Opella</h1>", unsafe_allow_html=True)
        
        # İSİM GÜNCELLENDİ
        st.markdown("<h3 style='text-align: center;'>Vendor Analysis Tool</h3>", unsafe_allow_html=True)
        st.info("Please enter your company email address.")
        
        with st.form("login_form"):
            email_input = st.text_input("Email Address").strip().lower()
            submit_button = st.form_submit_button("Secure Login", type="primary", use_container_width=True)

        if submit_button:
            allowed_domains = ["sanofi.com", "opella.com"]
            is_valid_domain = any(email_input.endswith(dom) for dom in allowed_domains)
            
            if not is_valid_domain:
                st.error("⛔ Invalid Domain.")
            else:
                users_df = load_user_db()
                user_record = users_df[users_df['Email'] == email_input]
                
                if not user_record.empty:
                    st.session_state['logged_in'] = True
                    st.session_state['user_email'] = email_input
                    st.session_state['user_name'] = user_record.iloc[0]['Name']
                    st.session_state['user_role'] = user_record.iloc[0]['Role']
                    st.rerun()
                else:
                    st.warning("⚠️ Access Denied.")
    st.stop() 

# ==========================================
# 3. HELPER FUNCTIONS
# ==========================================

def get_live_rate(base_currency):
    """Yahoo Finance'ten EUR kurunu çeker (Safe Mode)"""
    if not YFINANCE_AVAILABLE: return None
    if base_currency == "EUR": return 1.0
    try:
        ticker_symbol = f"EUR{base_currency}=X" 
        data = yf.Ticker(ticker_symbol)
        history = data.history(period="1d")
        if not history.empty:
            return history['Close'].iloc[-1]
        else:
            return None
    except:
        return None

def get_aging_bucket(payment_date, report_date):
    if pd.isna(payment_date): return "Not Due"
    days = (report_date - payment_date).days
    if days < 0: return "Not Due"
    elif days <= 30: return "1-30 Days"
    elif days <= 60: return "31-60 Days"
    elif days <= 90: return "61-90 Days"
    else: return "90+ Days"

def clean_sap_data(df):
    if 'Document Number' in df.columns:
        return df.dropna(subset=['Document Number'])
    return df

def write_optimized_excel(writer, df, sheet_name):
    if df.empty: return
    workbook = writer.book
    worksheet = workbook.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = worksheet
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#5b21b6', 'font_color': 'white', 'border': 1, 'align': 'center'})
    num_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
    txt_fmt = workbook.add_format({'border': 1})
    for col_num, value in enumerate(df.columns.values): worksheet.write(0, col_num, value, header_fmt)
    for row_idx, row in enumerate(df.itertuples(index=False), start=1):
        for col_idx, value in enumerate(row): worksheet.write(row_idx, col_idx, value, num_fmt if isinstance(value, (int, float)) else txt_fmt)
    worksheet.set_column(0, 0, 15); worksheet.set_column(1, 1, 35)

# ==========================================
# 4. SIDEBAR & NAVIGATION
# ==========================================
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    else:
        st.header("Opella Finance")
        
    st.markdown(f"👤 **{st.session_state['user_name']}**")
    st.caption(f"Role: {st.session_state['user_role']}")
    
    page_mode = "📊 Dashboard"
    if st.session_state['user_role'] == 'Admin':
        st.markdown("---")
        page_mode = st.radio("Navigate", ["📊 Dashboard", "⚙️ Admin Panel"])
    
    st.markdown("---")
    
    uploaded_file = None
    if page_mode == "📊 Dashboard":
        st.header("📂 Data Import")
        uploaded_file = st.file_uploader("Upload FBL1N Report (Excel)", type=["xlsx", "xls"])
        
        st.markdown("### ⚙️ Parameters")
        currency_list = ["EGP", "TRY", "EUR", "USD", "TND", "AED", "SAR", "GBP"]
        selected_currency = st.selectbox("Local Currency", currency_list, index=0)
        
        default_val = 52.50 if selected_currency == "EGP" else (35.00 if selected_currency == "TRY" else 1.00)
        
        if 'current_eur_rate' not in st.session_state:
            st.session_state['current_eur_rate'] = default_val
            
        col_p1, col_p2 = st.columns([2, 1])
        with col_p2:
            st.write("") 
            st.write("")
            if YFINANCE_AVAILABLE:
                if st.button("🌐 Get Rate", help="Fetch live rate"):
                    with st.spinner("Fetching..."):
                        live_rate = get_live_rate(selected_currency)
                        if live_rate:
                            st.session_state['current_eur_rate'] = live_rate
                            st.toast(f"Updated: {live_rate:.2f}", icon="✅")
                        else:
                            st.error("Failed.")
            else:
                st.caption("⚠️ Library Missing")
        
        with col_p1:
            eur_rate = st.number_input(
                f"EUR / {selected_currency}", 
                value=st.session_state['current_eur_rate'], 
                step=0.01,
                format="%.4f"
            )
    
    st.markdown("---")
    if st.button("🔒 Logout"):
        st.session_state['logged_in'] = False
        st.rerun()

# ==========================================
# 5. ADMIN PANEL LOGIC
# ==========================================
if page_mode == "⚙️ Admin Panel":
    st.title("⚙️ User Management")
    tab1, tab2 = st.tabs(["📂 **User List**", "➕ **Add New User**"])
    with tab1:
        current_users_df = load_user_db()
        edited_users_df = st.data_editor(current_users_df, num_rows="dynamic", use_container_width=True, key="user_editor")
        if st.button("💾 Save Changes", type="primary"):
            try:
                edited_users_df.to_excel(USER_DB_FILE, index=False)
                st.success("✅ Saved!"); time.sleep(1); st.rerun()
            except Exception as e: st.error(str(e))
    with tab2:
        with st.form("add_user_form"):
            new_email = st.text_input("Email").strip().lower()
            new_name = st.text_input("Name").strip()
            if st.form_submit_button("Add User", type="primary"):
                success, msg = add_user_to_db(new_email, new_name)
                if success: st.success(msg); time.sleep(1); st.rerun()
                else: st.error(msg)

# ==========================================
# 6. DASHBOARD LOGIC
# ==========================================
elif page_mode == "📊 Dashboard":
    # İSİM GÜNCELLENDİ
    st.title("📊 Vendor Analysis Tool")
    
    if uploaded_file:
        st.info("File ready. Check parameters and click Start.")
        if st.button("🚀 Start Analysis", type="primary"):
            with st.status("🔄 Processing...", expanded=True) as status:
                st.write("🧹 Cleaning Data...")
                try:
                    df_raw = pd.read_excel(uploaded_file)
                    df = clean_sap_data(df_raw)
                    
                    df['Posting Date'] = pd.to_datetime(df['Posting Date'], errors='coerce')
                    df['Payment date'] = pd.to_datetime(df['Payment date'], errors='coerce')
                    df['Amount'] = pd.to_numeric(df['Amount in local currency'], errors='coerce').fillna(0)
                    df['Supplier'] = df['Supplier'].fillna('N/A').astype(str)
                    df['Vendor name'] = df['Vendor name'].fillna(df['Supplier'])
                    df['G/L Account'] = df['G/L Account'].astype(str).apply(lambda x: x.split('.')[0] if '.' in x else x)
                    
                    safe_rate = eur_rate if eur_rate > 0 else 1.0
                    df['Amount_EUR'] = df['Amount'] / safe_rate
                    report_date = df['Posting Date'].max()
                    df['Aging Bucket'] = df['Payment date'].apply(lambda x: get_aging_bucket(x, report_date))
                    buckets_order = ["Not Due", "1-30 Days", "31-60 Days", "61-90 Days", "90+ Days"]

                    # ==========================================
                    # CORE LOGIC & PIVOTS
                    # ==========================================
                    
                    # 1. GL Summary
                    gl_pivot = df.pivot_table(index=['G/L Account'], columns='Aging Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets_order, fill_value=0)
                    gl_pivot['Total Balance'] = gl_pivot.sum(axis=1)
                    
                    def get_top_driver(sub_df):
                        if sub_df.empty: return "None"
                        return sub_df.groupby('Vendor name')['Amount'].sum().abs().idxmax()
                    
                    top_vendors = df.groupby('G/L Account').apply(get_top_driver).reset_index(name='Top Driver Vendor')
                    gl_final = gl_pivot.reset_index().merge(top_vendors, on='G/L Account', how='left')
                    cols = ['G/L Account', 'Top Driver Vendor'] + buckets_order + ['Total Balance']
                    gl_final = gl_final[cols].sort_values(by='Total Balance', key=abs, ascending=False)

                    # 2. VENDOR AGING
                    v_raw = df.pivot_table(index=['Supplier', 'Vendor name'], columns='Aging Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets_order, fill_value=0)
                    v_raw['Total Balance'] = v_raw.sum(axis=1)
                    
                    v_ap = v_raw[v_raw['Total Balance'] < 0].copy()
                    v_ap = v_ap.sort_values(by='Total Balance', ascending=True).reset_index()

                    # 3. DEBIT BALANCES
                    v_debit = v_raw[v_raw['Total Balance'] > 0].copy()
                    v_debit = v_debit.sort_values(by='Total Balance', ascending=False).reset_index()

                    # 4. PREPAYMENTS
                    dp_gls = ['16740100', '16740110', '16740000']
                    dp_df = df[df['G/L Account'].isin(dp_gls)]
                    dp_final = pd.DataFrame()
                    if not dp_df.empty:
                        dp_piv = dp_df.pivot_table(index=['Supplier', 'Vendor name'], columns='Aging Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets_order, fill_value=0)
                        dp_piv['Total Balance'] = dp_piv.sum(axis=1)
                        dp_final = dp_piv.sort_values(by='Total Balance', ascending=False).reset_index()

                    # EXPORT
                    output_excel = io.BytesIO()
                    with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                        write_optimized_excel(writer, gl_final, 'GL Summary (BS Review)')
                        write_optimized_excel(writer, v_ap, 'AP Vendor Aging (Credit)')
                        write_optimized_excel(writer, v_debit, 'Debit Balances (Debit)')
                        write_optimized_excel(writer, dp_final, 'Prepayments')
                    
                    status.update(label="✅ Ready!", state="complete", expanded=False)
                except Exception as e:
                    st.error(f"Error: {e}")
                    st.stop()

            # ==========================================
            # VISUALIZATION
            # ==========================================
            st.caption(f"📅 Report Date: {report_date.strftime('%d-%b-%Y')} | 💱 FX: {safe_rate:,.2f}")
            
            # Helper for 'k' view
            def to_k_view(df_in):
                if df_in.empty: return pd.DataFrame()
                cols = buckets_order + ['Total Balance']
                return (df_in.set_index(df_in.columns[0])[cols] / 1000)

            # 1. GL
            st.markdown("### 1. GL Account Aging Summary")
            gl_eur_raw = df.pivot_table(index=['G/L Account'], columns='Aging Bucket', values='Amount_EUR', aggfunc='sum', fill_value=0).reindex(columns=buckets_order, fill_value=0)
            gl_eur_raw['Total Balance'] = gl_eur_raw.sum(axis=1)
            gl_eur_final = gl_eur_raw.sort_values(by='Total Balance', key=abs, ascending=False)

            c1, c2 = st.columns(2)
            with c1:
                st.info(f"**k{selected_currency}**")
                gl_disp = gl_final.set_index('G/L Account')[buckets_order + ['Total Balance']] / 1000
                st.dataframe(gl_disp.style.format("{:,.0f}"), use_container_width=True)
            with c2:
                st.warning("**kEUR**")
                st.dataframe((gl_eur_final/1000).style.format("{:,.0f}"), use_container_width=True)
            
            st.divider()

            # 2. VENDOR AGING
            st.markdown("### 2. Vendor Aging (Payables Only)")
            top_ap_local = v_ap.head(10)
            
            v_eur_raw = df.pivot_table(index=['Supplier', 'Vendor name'], columns='Aging Bucket', values='Amount_EUR', aggfunc='sum', fill_value=0).reindex(columns=buckets_order, fill_value=0)
            v_eur_raw['Total Balance'] = v_eur_raw.sum(axis=1)
            v_eur_ap = v_eur_raw[v_eur_raw['Total Balance'] < 0].sort_values(by='Total Balance', ascending=True).reset_index().head(10)

            c3, c4 = st.columns(2)
            with c3:
                st.info(f"**Top 10 Payables (k{selected_currency})**")
                if not top_ap_local.empty:
                    st.dataframe(to_k_view(top_ap_local.drop(columns=['Supplier'])).style.format("{:,.0f}"), use_container_width=True)
                else: st.write("No Payables found.")
            with c4:
                st.warning("**Top 10 Payables (kEUR)**")
                if not v_eur_ap.empty:
                    st.dataframe(to_k_view(v_eur_ap.drop(columns=['Supplier'])).style.format("{:,.0f}"), use_container_width=True)
                else: st.write("No Payables found.")

            st.divider()

            # 3. DEBIT BALANCES
            st.markdown("### 3. Top Debit Balances")
            top_debit_local = v_debit.head(10)
            
            # --- FIX: SADECE SAYISAL KOLONLARA FORMAT UYGULA ---
            numeric_cols = buckets_order + ['Total Balance']

            if not top_debit_local.empty:
                disp_debit = top_debit_local[['Vendor name', 'Total Balance'] + buckets_order].copy()
                st.dataframe(disp_debit.style.format("{:,.2f}", subset=numeric_cols), use_container_width=True)
            else: st.write("No Debit Balances found.")

            # 4. PREPAYMENTS
            st.markdown("### 4. Prepayments")
            if not dp_final.empty:
                st.dataframe(
                    dp_final[['Vendor name', 'Total Balance'] + buckets_order].head(10).style.format("{:,.2f}", subset=numeric_cols), 
                    use_container_width=True
                )
            else: st.success("No Data")

            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button("📥 Download Excel Report", output_excel.getvalue(), f"Opella_AP_{datetime.now().strftime('%Y%m%d')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

    else:
        st.info("👋 Upload FBL1N Excel file to start.")
