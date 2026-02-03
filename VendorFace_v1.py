import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import os
import time
import yfinance as yf # Yahoo Finance Kütüphanesi

# ==========================================
# 1. CONFIGURATION & SETUP
# ==========================================
st.set_page_config(page_title="VendorFace | Opella Finance", layout="wide", page_icon="🛡️")
USER_DB_FILE = "users.xlsx"
ADMIN_EMAIL = "can.adiguzel@sanofi.com" 

# ==========================================
# 2. AUTHENTICATION SYSTEM
# ==========================================

def load_user_db():
    if not os.path.exists(USER_DB_FILE):
        # Initial User List
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
        # Logo Check
        if os.path.exists("logo.png"):
            st.image("logo.png", width=200)
        else:
            st.markdown("<h1 style='text-align: center; color:#5b21b6;'>Opella</h1>", unsafe_allow_html=True)
        
        st.markdown("<h3 style='text-align: center;'>VendorFace Login</h3>", unsafe_allow_html=True)
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
# 3. HELPER FUNCTIONS (Yahoo & Logic)
# ==========================================

def get_live_rate(base_currency):
    """Yahoo Finance'ten EUR kurunu çeker"""
    if base_currency == "EUR": return 1.0
    try:
        # Yahoo sembol formatı: EURTRY=X (1 EUR kaç TRY)
        ticker_symbol = f"EUR{base_currency}=X" 
        data = yf.Ticker(ticker_symbol)
        history = data.history(period="1d")
        if not history.empty:
            rate = history['Close'].iloc[-1]
            return rate # Örn: 35.50 (1 EUR = 35.50 TRY)
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
    # Removes subtotals based on missing Document Numbers
    if 'Document Number' in df.columns:
        return df.dropna(subset=['Document Number'])
    return df

def create_k_pivot(data, index_col, value_col, buckets):
    piv = data.pivot_table(index=index_col, columns='Aging Bucket', values=value_col, aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
    piv['Total Balance'] = piv.sum(axis=1)
    # Sort Descending by Absolute Value
    return (piv / 1000).sort_values(by='Total Balance', key=abs, ascending=False)

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
    
    # Navigation Logic (Admin Only)
    page_mode = "📊 Dashboard"
    if st.session_state['user_role'] == 'Admin':
        st.markdown("---")
        page_mode = st.radio("Navigate", ["📊 Dashboard", "⚙️ Admin Panel"])
    
    st.markdown("---")
    
    # Dashboard Controls
    uploaded_file = None
    if page_mode == "📊 Dashboard":
        st.header("📂 Data Import")
        uploaded_file = st.file_uploader("Upload FBL1N Report (Excel)", type=["xlsx", "xls"])
        
        st.markdown("### ⚙️ Parameters")
        currency_list = ["EGP", "TRY", "EUR", "USD", "TND", "AED", "SAR", "GBP"]
        selected_currency = st.selectbox("Local Currency", currency_list, index=0)
        
        # --- LIVE RATE LOGIC ---
        # Default değerler
        default_val = 52.50 if selected_currency == "EGP" else (35.00 if selected_currency == "TRY" else 1.00)
        
        # Session state ile kuru tutuyoruz ki sayfa yenilenince gitmesin
        if 'current_eur_rate' not in st.session_state:
            st.session_state['current_eur_rate'] = default_val
            
        col_p1, col_p2 = st.columns([2, 1])
        with col_p2:
            st.write("") # Boşluk
            st.write("")
            if st.button("🌐 Get Rate", help="Fetch live rate from Yahoo Finance"):
                with st.spinner("Fetching..."):
                    live_rate = get_live_rate(selected_currency)
                    if live_rate:
                        st.session_state['current_eur_rate'] = live_rate
                        st.toast(f"Updated: 1 EUR = {live_rate:.2f} {selected_currency}", icon="✅")
                    else:
                        st.error("Failed.")
        
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
    st.markdown("Manage authorized users for VendorFace Dashboard.")
    
    tab1, tab2 = st.tabs(["📂 **User List**", "➕ **Add New User**"])
    
    with tab1:
        st.info("💡 Edit directly in the table below. Select rows and press 'Delete' to remove users.")
        current_users_df = load_user_db()
        edited_users_df = st.data_editor(current_users_df, num_rows="dynamic", use_container_width=True, key="user_editor")
        
        if st.button("💾 Save Changes", type="primary"):
            try:
                edited_users_df.to_excel(USER_DB_FILE, index=False)
                st.success("✅ User list updated successfully!")
                time.sleep(1)
                st.rerun()
            except Exception as e:
                st.error(f"Error saving changes: {e}")

    with tab2:
        with st.form("add_user_form"):
            new_email = st.text_input("New User Email").strip().lower()
            new_name = st.text_input("New User Name").strip()
            if st.form_submit_button("Add User", type="primary"):
                success, msg = add_user_to_db(new_email, new_name)
                if success: st.success(msg); time.sleep(1); st.rerun()
                else: st.error(msg)

# ==========================================
# 6. DASHBOARD LOGIC
# ==========================================
elif page_mode == "📊 Dashboard":
    st.title("📊 Executive BS Review Dashboard")
    st.markdown("Analyze AP Aging, Prepayments, and Debit Balances.")

    if uploaded_file:
        st.info("File uploaded successfully. Please verify currency settings and click Start.")
        
        # START BUTTON
        if st.button("🚀 Start Analysis", type="primary"):
            
            with st.status("🔄 Processing Data...", expanded=True) as status:
                st.write("🧹 Cleaning SAP Data...")
                df_raw = pd.read_excel(uploaded_file)
                df = clean_sap_data(df_raw)
                
                # Transformations
                df['Posting Date'] = pd.to_datetime(df['Posting Date'], errors='coerce')
                df['Payment date'] = pd.to_datetime(df['Payment date'], errors='coerce')
                df['Amount'] = pd.to_numeric(df['Amount in local currency'], errors='coerce').fillna(0)
                df['Supplier'] = df['Supplier'].fillna('N/A').astype(str)
                df['Vendor name'] = df['Vendor name'].fillna(df['Supplier'])
                
                # GL Cleaning
                df['G/L Account'] = df['G/L Account'].astype(str).apply(lambda x: x.split('.')[0] if '.' in x else x)
                
                # FX
                safe_rate = eur_rate if eur_rate > 0 else 1.0
                df['Amount_EUR'] = df['Amount'] / safe_rate
                
                # Buckets
                report_date = df['Posting Date'].max()
                df['Aging Bucket'] = df['Payment date'].apply(lambda x: get_aging_bucket(x, report_date))
                buckets_order = ["Not Due", "1-30 Days", "31-60 Days", "61-90 Days", "90+ Days"]

                # --- PREPARE DATASETS ---
                
                # 1. GL Summary
                gl_pivot = df.pivot_table(index=['G/L Account'], columns='Aging Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets_order, fill_value=0)
                gl_pivot['Total Balance'] = gl_pivot.sum(axis=1)
                
                # Top Driver Logic
                def get_top_driver(sub_df):
                    if sub_df.empty: return "None"
                    return sub_df.groupby('Vendor name')['Amount'].sum().abs().idxmax()
                
                top_vendors = df.groupby('G/L Account').apply(get_top_driver).reset_index(name='Top Driver Vendor')
                gl_summary_final = gl_pivot.reset_index().merge(top_vendors, on='G/L Account', how='left')
                cols = ['G/L Account', 'Top Driver Vendor'] + buckets_order + ['Total Balance']
                gl_summary_final = gl_summary_final[cols].sort_values(by='Total Balance', key=abs, ascending=False)

                # 2. Vendor Aging
                vendor_pivot = df.pivot_table(index=['Supplier', 'Vendor name'], columns='Aging Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets_order, fill_value=0)
                vendor_pivot['Total Balance'] = vendor_pivot.sum(axis=1)
                vendor_pivot = vendor_pivot.sort_values(by='Total Balance', key=abs, ascending=False).reset_index()

                # 3. Debit Balances
                debit_df = vendor_pivot[vendor_pivot['Total Balance'] > 0].copy()
                
                # 4. Prepayments
                dp_gls = ['16740100', '16740110', '16740000'] # Modify if needed
                dp_df = df[df['G/L Account'].isin(dp_gls)]
                
                if not dp_df.empty:
                    dp_pivot = dp_df.pivot_table(index=['Supplier', 'Vendor name'], columns='Aging Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets_order, fill_value=0)
                    dp_pivot['Total Balance'] = dp_pivot.sum(axis=1)
                    dp_pivot_sorted = dp_pivot.sort_values(by='Total Balance', key=abs, ascending=False).reset_index()
                else:
                    dp_pivot_sorted = pd.DataFrame()

                # --- EXPORT ---
                output_excel = io.BytesIO()
                with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                    write_optimized_excel(writer, gl_summary_final, 'GL Summary (BS Review)')
                    write_optimized_excel(writer, vendor_pivot, 'AP Vendor Aging')
                    write_optimized_excel(writer, dp_pivot_sorted, 'Prepayments (Downpayments)')
                    write_optimized_excel(writer, debit_df, 'Debit Balances')
                
                status.update(label="✅ Analysis Complete!", state="complete", expanded=False)

            # ==========================================
            # DASHBOARD VISUALIZATION
            # ==========================================
            st.caption(f"📅 Report Date: {report_date.strftime('%d-%b-%Y')} | 💱 FX Rate: 1 EUR = {safe_rate:,.2f} {selected_currency}")
            
            # 1. GL AGING
            st.markdown("### 1. GL Account Aging Summary")
            c1, c2 = st.columns(2)
            with c1:
                st.info(f"**Local Currency (k{selected_currency})**")
                st.dataframe(create_k_pivot(df, 'G/L Account', 'Amount', buckets_order).style.format("{:,.0f}"), use_container_width=True)
            with c2:
                st.warning("**Group Currency (kEUR)**")
                st.dataframe(create_k_pivot(df, 'G/L Account', 'Amount_EUR', buckets_order).style.format("{:,.0f}"), use_container_width=True)
            
            st.divider()

            # 2. TOP VENDORS
            st.markdown("### 2. Top High Value Vendors")
            top_list = df.groupby('Vendor name')['Amount'].sum().abs().sort_values(ascending=False).head(10).index.tolist()
            df_top = df[df['Vendor name'].isin(top_list)]
            
            c3, c4 = st.columns(2)
            with c3:
                st.info(f"**Top 10 Vendors (k{selected_currency})**")
                st.dataframe(create_k_pivot(df_top, 'Vendor name', 'Amount', buckets_order).style.format("{:,.0f}"), use_container_width=True)
            with c4:
                st.warning("**Top 10 Vendors (kEUR)**")
                st.dataframe(create_k_pivot(df_top, 'Vendor name', 'Amount_EUR', buckets_order).style.format("{:,.0f}"), use_container_width=True)

            st.divider()

            # 3. PREPAYMENTS
            st.markdown("### 3. Prepayments (Downpayments) Overview")
            if not dp_df.empty:
                c5, c6 = st.columns(2)
                with c5:
                    st.success(f"**Prepayments (k{selected_currency})**")
                    st.dataframe(create_k_pivot(dp_df, 'Vendor name', 'Amount', buckets_order).style.format("{:,.0f}"), use_container_width=True)
                with c6:
                    st.success("**Prepayments (kEUR)**")
                    st.dataframe(create_k_pivot(dp_df, 'Vendor name', 'Amount_EUR', buckets_order).style.format("{:,.0f}"), use_container_width=True)
            else:
                st.success("No Prepayment/Downpayment GL activity found.")

            st.divider()

            # 4. DEBIT BALANCES
            st.markdown("### 4. Top Debit Balances")
            if not debit_df.empty:
                debit_top = debit_df.head(10) # Already sorted
                st.dataframe(debit_top[['Vendor name', 'Total Balance'] + buckets_order].style.format("{:,.2f}"), use_container_width=True)
            else:
                st.write("No Debit Balances found.")

            # DOWNLOAD
            st.markdown("<br>", unsafe_allow_html=True)
            st.download_button(
                label="📥 Download Full Report (Excel)",
                data=output_excel.getvalue(),
                file_name=f"Opella_AP_Analysis_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

    else:
        st.info("👋 Waiting for file upload...")
