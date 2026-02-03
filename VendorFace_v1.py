import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import os
import time
import streamlit.components.v1 as components

# ==========================================
# 1. CONFIGURATION & SETUP
# ==========================================
st.set_page_config(page_title="VendorFace | Smart Finance Suite", layout="wide", page_icon="🛡️")
USER_DB_FILE = "users.xlsx"
ADMIN_EMAIL = "can.adiguzel@sanofi.com" 

# ==========================================
# 2. AUTHENTICATION SYSTEM
# ==========================================

def load_user_db():
    if not os.path.exists(USER_DB_FILE):
        # İlk çalıştırmada admin kullanıcısını oluşturur
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

# Session State Başlangıcı
if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if 'user_email' not in st.session_state: st.session_state['user_email'] = ""

# --- LOGIN EKRANI ---
if not st.session_state['logged_in']:
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown("<h1 style='text-align: center; color:#1e3a8a;'>🛡️ VendorFace</h1>", unsafe_allow_html=True)
        st.markdown("<h3 style='text-align: center;'>Financial Intelligence Suite</h3>", unsafe_allow_html=True)
        
        st.info("Please enter your company email address to access the dashboard.")
        
        with st.form("login_form"):
            email_input = st.text_input("Email Address", placeholder="name.surname@sanofi.com").strip().lower()
            submit_button = st.form_submit_button("Secure Login", type="primary", use_container_width=True)

        if submit_button:
            # Domain Kontrolü
            allowed_domains = ["sanofi.com", "opella.com"]
            is_valid_domain = any(email_input.endswith(dom) for dom in allowed_domains)
            
            if not is_valid_domain:
                st.error("⛔ Invalid Domain. Please enter a valid company email domain.")
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
                    st.warning("⚠️ Access Denied: Your email is not authorized.")
    st.stop() 

# ==========================================
# 3. SIDEBAR & NAVIGASYON
# ==========================================

with st.sidebar:
    st.markdown(f"👤 **{st.session_state['user_name']}**")
    st.caption(f"Role: {st.session_state['user_role']}")
    st.markdown("---")
    
    st.header("📂 Data Import")
    uploaded_file = st.file_uploader("Upload FBL1N Report (Excel)", type=["xlsx", "xls"])
    
    st.markdown("### ⚙️ Parameters")
    currency_list = ["EGP", "TRY", "EUR", "USD", "TND", "AED", "SAR", "GBP"]
    selected_currency = st.selectbox("Local Currency", currency_list, index=0)
    
    default_rate = 52.50 if selected_currency == "EGP" else (35.00 if selected_currency == "TRY" else 1.00)
    eur_rate = st.number_input(f"EUR / {selected_currency} Rate", value=default_rate, step=0.01)
    
    st.markdown("---")
    if st.button("🔒 Logout"):
        st.session_state['logged_in'] = False
        st.rerun()

# --- Custom Styling ---
st.markdown("""
    <style>
    .main { background-color: #f8fafc; }
    h1 { color: #1e3a8a; font-family: 'Segoe UI', sans-serif; font-weight: 800; }
    h3 { color: #1e40af; border-left: 5px solid #3b82f6; padding-left: 10px; margin-top: 30px; }
    .stDataFrame { border: 1px solid #e2e8f0; border-radius: 8px; }
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# 4. HELPER FUNCTIONS
# ==========================================

def get_aging_bucket(payment_date, report_date):
    if pd.isna(payment_date): return "Not Due"
    days = (report_date - payment_date).days
    if days < 0: return "Not Due"
    elif days <= 30: return "1-30 Days"
    elif days <= 60: return "31-60 Days"
    elif days <= 90: return "61-90 Days"
    else: return "90+ Days"

def write_optimized_excel(writer, df, sheet_name):
    """Excel sheet'lerini otomatik formatlayan yardımcı fonksiyon"""
    if df.empty: return
    workbook = writer.book
    worksheet = workbook.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = worksheet
    
    # Formatlar
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1e3a8a', 'font_color': 'white', 'border': 1, 'align': 'center'})
    num_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
    txt_fmt = workbook.add_format({'border': 1})
    
    # Header Yazdırma
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_fmt)
        
    # Data Yazdırma
    for row_idx, row in enumerate(df.itertuples(index=False), start=1):
        for col_idx, value in enumerate(row):
            worksheet.write(row_idx, col_idx, value, num_fmt if isinstance(value, (int, float)) else txt_fmt)
            
    worksheet.set_column(0, 0, 15) # İlk kolon genişliği
    worksheet.set_column(1, 1, 35) # İsim kolonu genişliği

# ==========================================
# 5. MAIN LOGIC & DASHBOARD
# ==========================================

st.title("📊 Executive BS Review Dashboard")

if uploaded_file:
    with st.status("🚀 Processing Data & Generating Report...", expanded=True) as status:
        st.write("📂 1. Reading Excel file...")
        try:
            df = pd.read_excel(uploaded_file)
            
            # --- DATA CLEANING ---
            st.write("🧹 2. Cleaning & Transforming...")
            df['Posting Date'] = pd.to_datetime(df['Posting Date'], errors='coerce')
            df['Payment date'] = pd.to_datetime(df['Payment date'], errors='coerce')
            df['Amount'] = pd.to_numeric(df['Amount in local currency'], errors='coerce').fillna(0)
            df['Supplier'] = df['Supplier'].fillna('N/A').astype(str)
            df['Vendor name'] = df['Vendor name'].fillna('Unknown')
            
            # GL Account Temizliği
            df['G/L Account'] = df['G/L Account'].apply(lambda x: str(int(float(x))) if str(x).replace('.','',1).isdigit() else str(x))
            
            # EUR Hesaplama
            safe_rate = eur_rate if eur_rate > 0 else 1.0
            df['Amount_EUR'] = df['Amount'] / safe_rate

            # Aging Bucket
            report_date = df['Posting Date'].max()
            df['Aging Bucket'] = df['Payment date'].apply(lambda x: get_aging_bucket(x, report_date))
            buckets_order = ["Not Due", "1-30 Days", "31-60 Days", "61-90 Days", "90+ Days"]

            # --- HAZIRLIK: EXCEL EXPORT DATALARI ---
            
            # 1. GL Summary (BS Review)
            gl_pivot = df.pivot_table(
                index=['G/L Account'], 
                columns='Aging Bucket', 
                values='Amount', 
                aggfunc='sum', 
                fill_value=0
            ).reindex(columns=buckets_order, fill_value=0)
            gl_pivot['Total Balance'] = gl_pivot.sum(axis=1)
            
            # Top Driver Vendor Bulma
            def get_top_driver(sub_df):
                if sub_df.empty: return "None"
                vendor_sums = sub_df.groupby('Vendor name')['Amount'].sum().abs()
                return vendor_sums.idxmax() if not vendor_sums.empty else "None"

            top_vendors = df.groupby('G/L Account').apply(get_top_driver).reset_index(name='Top Driver Vendor')
            gl_summary_final = gl_pivot.reset_index().merge(top_vendors, on='G/L Account', how='left')
            cols = ['G/L Account', 'Top Driver Vendor'] + buckets_order + ['Total Balance']
            gl_summary_final = gl_summary_final[cols].sort_values(by='Total Balance', key=abs, ascending=False)

            # 2. Vendor Detail (Operasyonel)
            vendor_pivot = df.pivot_table(
                index=['Supplier', 'Vendor name'],
                columns='Aging Bucket',
                values='Amount',
                aggfunc='sum',
                fill_value=0
            ).reindex(columns=buckets_order, fill_value=0)
            vendor_pivot['Total Balance'] = vendor_pivot.sum(axis=1)
            vendor_pivot = vendor_pivot.sort_values(by='Total Balance', ascending=True).reset_index()

            # 3. Debit Balances & Downpayments
            debit_df = vendor_pivot[vendor_pivot['Total Balance'] > 0]
            dp_gls = ['16740100', '16740110'] # Örnek GL'ler
            dp_df = vendor_pivot[vendor_pivot['Supplier'].isin(df[df['G/L Account'].isin(dp_gls)]['Supplier'])]

            # --- EXCEL OLUŞTURMA ---
            output_excel = io.BytesIO()
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                write_optimized_excel(writer, gl_summary_final, 'GL Summary (BS Review)')
                write_optimized_excel(writer, vendor_pivot, 'AP Vendor Aging')
                write_optimized_excel(writer, debit_df, 'Debit Balances')
                write_optimized_excel(writer, dp_df, 'Downpayments')
            
            status.update(label="✅ Analysis Complete!", state="complete", expanded=False)

        except Exception as e:
            st.error(f"Critical Error: {e}")
            st.stop()

    # =========================================================
    # DASHBOARD VISUALIZATION
    # =========================================================
    
    st.caption(f"📅 Report Date: {report_date.strftime('%d-%b-%Y')} | 💱 FX Rate: 1 EUR = {safe_rate:,.2f} {selected_currency}")
    
    # Helper for 'k' (Thousands) View
    def create_k_pivot(data, index_col, value_col):
        piv = data.pivot_table(
            index=index_col, columns='Aging Bucket', values=value_col, aggfunc='sum', fill_value=0
        ).reindex(columns=buckets_order, fill_value=0)
        piv['Total Balance'] = piv.sum(axis=1)
        return (piv / 1000).sort_values(by='Total Balance', key=abs, ascending=False)

    # --- SECTION 1: GL ACCOUNT AGING ---
    st.markdown("### 1. GL Account Aging Summary")
    
    col_gl1, col_gl2 = st.columns(2)
    with col_gl1:
        st.info(f"**Local Currency (k{selected_currency})**")
        gl_k_local = create_k_pivot(df, 'G/L Account', 'Amount')
        st.dataframe(gl_k_local.style.format("{:,.0f}"), use_container_width=True)
    
    with col_gl2:
        st.warning(f"**Group Currency (kEUR)**")
        gl_k_eur = create_k_pivot(df, 'G/L Account', 'Amount_EUR')
        st.dataframe(gl_k_eur.style.format("{:,.0f}"), use_container_width=True)
        
    st.divider()

    # --- SECTION 2: TOP 10 VENDORS ---
    st.markdown("### 2. Top 10 High Value Vendors")
    
    # Identify Top 10 by Absolute Value
    top10_names = df.groupby('Vendor name')['Amount'].sum().abs().sort_values(ascending=False).head(10).index.tolist()
    df_top10 = df[df['Vendor name'].isin(top10_names)]
    
    col_v1, col_v2 = st.columns(2)
    with col_v1:
        st.info(f"**Top 10 Vendors (k{selected_currency})**")
        v_k_local = create_k_pivot(df_top10, 'Vendor name', 'Amount')
        st.dataframe(v_k_local.style.format("{:,.0f}"), use_container_width=True)
        
    with col_v2:
        st.warning("**Top 10 Vendors (kEUR)**")
        v_k_eur = create_k_pivot(df_top10, 'Vendor name', 'Amount_EUR')
        st.dataframe(v_k_eur.style.format("{:,.0f}"), use_container_width=True)
        
    st.divider()

    # --- SECTION 3: DEBIT BALANCES ---
    st.markdown("### 3. Top 10 Debit Balances (Advances/Returns)")
    
    # Logic: Sum > 0
    vendor_sums = df.groupby('Vendor name')[['Amount', 'Amount_EUR']].sum()
    debit_vendors = vendor_sums[vendor_sums['Amount'] > 0].sort_values(by='Amount', ascending=False).head(10)
    
    if not debit_vendors.empty:
        df_debit_top = df[df['Vendor name'].isin(debit_vendors.index)]
        
        # Pivot by Aging
        debit_pivot = df_debit_top.pivot_table(
            index='Vendor name', columns='Aging Bucket', values='Amount', aggfunc='sum', fill_value=0
        ).reindex(columns=buckets_order, fill_value=0)
        
        # Add Totals Columns
        debit_pivot = debit_pivot / 1000 # Convert buckets to 'k'
        debit_pivot[f'Total k{selected_currency}'] = debit_vendors['Amount'] / 1000
        debit_pivot['Total kEUR'] = debit_vendors['Amount_EUR'] / 1000
        
        # Reorder
        final_cols = buckets_order + [f'Total k{selected_currency}', 'Total kEUR']
        debit_pivot = debit_pivot[final_cols].sort_values(by=f'Total k{selected_currency}', ascending=False)
        
        st.dataframe(
            debit_pivot.style.format("{:,.1f}").background_gradient(cmap="Greens", subset=[f'Total k{selected_currency}']),
            use_container_width=True
        )
    else:
        st.success("No significant Debit Balances found.")

    # --- DOWNLOAD BUTTON ---
    st.markdown("<br>", unsafe_allow_html=True)
    st.download_button(
        label="📥 Download Full Analysis Report (Excel)",
        data=output_excel.getvalue(),
        file_name=f"AP_Dashboard_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )

else:
    # Empty State
    col_c1, col_c2 = st.columns([1, 2])
    with col_c1:
        st.info("👈 **Start Here**\nPlease upload your FBL1N Excel file from the sidebar.")
    with col_c2:
        st.markdown("### Ready for Analysis")
        st.markdown("""
        - **GL Review:** Automatic Balance Sheet aging per account.
        - **Currency:** Dual view (Local & EUR).
        - **Risk:** Identify 90+ days items instantly.
        """)
