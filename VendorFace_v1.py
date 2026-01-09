import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import os
import time
import streamlit.components.v1 as components

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="VendorFace | Enterprise Login", layout="wide")
USER_DB_FILE = "users.xlsx"
# Admin Yetkisi (Bu mail ile girildiğinde solda Admin Paneli açılır)
ADMIN_EMAIL = "can.adiguzel@sanofi.com"

# --- 2. AUTHENTICATION SYSTEM ---

def load_user_db():
    """
    Kullanıcı veritabanını yükler. 
    Eğer dosya yoksa, ön tanımlı VIP listesi ile oluşturur.
    """
    if not os.path.exists(USER_DB_FILE):
        # --- ÖN TANIMLI KULLANICI LİSTESİ (VIP LIST) ---
        initial_users = [
            # Admin
            {"Email": ADMIN_EMAIL, "Name": "Can Adiguzel", "Role": "Admin"},
            # Ekip Üyeleri
            {"Email": "AyseDeniz.Sen@sanofi.com", "Name": "AyseDeniz Sen", "Role": "User"},
            {"Email": "Hassan.Sadek@sanofi.com", "Name": "Hassan Sadek", "Role": "User"},
            {"Email": "Omar.Kordy@sanofi.com", "Name": "Omar Kordy", "Role": "User"},
            {"Email": "Rishabh.Tiwari@sanofi.com", "Name": "Rishabh Tiwari", "Role": "User"},
            {"Email": "Molka.Mathlouthi@sanofi.com", "Name": "Molka Mathlouthi", "Role": "User"},
            {"Email": "Shweta.Sharma3@sanofi.com", "Name": "Shweta Sharma", "Role": "User"},
            {"Email": "Cedric.Fallu@sanofi.com", "Name": "Cedric Fallu", "Role": "User"}
        ]
        
        df = pd.DataFrame(initial_users)
        # Mailleri küçük harfe çevir ve temizle
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
        "Email": [email],
        "Name": [name],
        "Role": ["User"],
        "Added_Date": [datetime.now().strftime("%Y-%m-%d")]
    })
    df = pd.concat([df, new_user], ignore_index=True)
    try:
        df.to_excel(USER_DB_FILE, index=False)
        return True, "User authorized successfully."
    except Exception as e:
        return False, f"Error saving DB (Close Excel file if open): {e}"

# Session State Kontrolü
if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
if 'user_email' not in st.session_state:
    st.session_state['user_email'] = ""

# --- LOGIN SCREEN ---
if not st.session_state['logged_in']:
    col1, col2, col3 = st.columns([1, 1, 1])
    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)
        if os.path.exists("logo.png"):
            st.image("logo.png", width=200)
        else:
            st.markdown("<h1 style='text-align: center; color:#1e3a8a;'>🛡️ VendorFace</h1>", unsafe_allow_html=True)
        
        st.markdown("### 👋 Corporate Login")
        st.info("Please enter your company email address to access the dashboard.")
        
        # Enter tuşu ile giriş yapabilmek için form kullanımı
        with st.form("login_form"):
            email_input = st.text_input("Email Address", placeholder="name.surname@sanofi.com").strip().lower()
            submit_button = st.form_submit_button("Secure Login", type="primary", use_container_width=True)

        if submit_button:
            # Domain Kontrolü
            allowed_domains = ["sanofi.com", "opella.com"]
            is_valid_domain = any(email_input.endswith(dom) for dom in allowed_domains)
            
            if not is_valid_domain:
                st.error("⛔ Invalid Domain. Only @sanofi.com and @opella.com addresses are allowed.")
            else:
                # DB Kontrolü
                users_df = load_user_db()
                user_record = users_df[users_df['Email'] == email_input]
                
                if not user_record.empty:
                    # Başarılı Giriş
                    st.session_state['logged_in'] = True
                    st.session_state['user_email'] = email_input
                    st.session_state['user_name'] = user_record.iloc[0]['Name']
                    st.session_state['user_role'] = user_record.iloc[0]['Role']
                    st.rerun()
                else:
                    # Yetkisiz Kullanıcı -> Mail Linki
                    st.warning("⚠️ Access Denied: Your email is not authorized.")
                    
                    subject = "VendorFace Access Request"
                    body = f"Hello Admin,%0D%0A%0D%0AI would like to request access to the VendorFace Dashboard.%0D%0A%0D%0AMy Email: {email_input}%0D%0AThank you."
                    mailto_link = f"mailto:{ADMIN_EMAIL}?subject={subject}&body={body}"
                    
                    st.markdown(f"""
                        <a href="{mailto_link}" target="_blank" style="text-decoration:none;">
                            <div style="background-color:#fefce8; border:1px solid #facc15; color:#854d0e; padding:10px; border-radius:5px; text-align:center; font-weight:bold;">
                                📩 Click here to Request Access via Outlook
                            </div>
                        </a>
                        """, unsafe_allow_html=True)
                    st.caption("Admin will receive your request and notify you once approved.")
    
    st.stop() 

# ==========================================
# ANA UYGULAMA (Giriş Başarılı)
# ==========================================

# --- Sidebar: User Info & Admin Panel ---
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    
    st.markdown(f"👤 **{st.session_state['user_name']}**")
    st.caption(f"Logged in as: {st.session_state['user_email']}")
    
    if st.button("🔒 Logout"):
        st.session_state['logged_in'] = False
        st.rerun()
    
    st.divider()
    
    # --- ADMIN PANEL (Sadece Admin Görür) ---
    if st.session_state['user_email'] == ADMIN_EMAIL.lower() or st.session_state.get('user_role') == 'Admin':
        with st.expander("⚙️ Admin Panel (Add User)"):
            st.write("Authorize new colleague:")
            new_u_email = st.text_input("New User Email").strip().lower()
            new_u_name = st.text_input("New User Name").strip()
            
            if st.button("Add User"):
                if new_u_email and new_u_name:
                    success, msg = add_user_to_db(new_u_email, new_u_name)
                    if success:
                        st.success(f"✅ {new_u_name} added!")
                    else:
                        st.error(msg)
                else:
                    st.error("Please fill both fields.")
    
    # Normal Ayarlar
    st.header("Settings")
    eur_rate = st.number_input("EUR / EGP Rate", value=52.50, step=0.01)
    uploaded_file = st.file_uploader("Upload SAP FBL1N Report", type=["xlsx", "xls"])

# --- Custom Styling ---
st.markdown("""
    <style>
    .main { background-color: #f8fafc; }
    h1 { color: #1e3a8a; font-family: 'Segoe UI', sans-serif; font-weight: 800; letter-spacing: -0.5px; }
    h3 { color: #1e40af; border-left: 5px solid #3b82f6; padding-left: 10px; margin-top: 30px; }
    .footer { text-align: center; color: #94a3b8; font-family: 'Courier New', monospace; font-size: 0.8rem; margin-top: 50px; border-top: 1px solid #e2e8f0; padding-top: 20px; }
    thead tr th:first-child { display:none }
    tbody th { display:none }
    .stTable { width: 100% !important; }
    .print-btn { background-color: #475569; color: white; border: none; padding: 10px 20px; text-align: center; display: inline-block; font-size: 16px; cursor: pointer; border-radius: 8px; font-weight: bold; width: 100%; }
    .print-btn:hover { background-color: #334155; transform: scale(1.02); }
    @media print { [data-testid="stSidebar"] { display: none; } .stButton { display: none; } .print-hide { display: none; } iframe { display: none; } .footer { position: fixed; bottom: 0; width: 100%; } }
    </style>
    """, unsafe_allow_html=True)

# --- Main App Logic ---
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.title("Account Payable Aging Analysis")
    st.caption(f"Financial Intelligence Dashboard | Welcome, {st.session_state['user_name']}")

# Optimized Excel Writer Helper
def write_optimized_excel(writer, df, sheet_name):
    workbook = writer.book
    worksheet = workbook.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = worksheet
    # Kurumsal Mavi Başlık
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#1e3a8a', 'font_color': 'white', 'border': 1, 'align': 'center'})
    str_fmt = workbook.add_format({'border': 1, 'align': 'left'})
    num_fmt = workbook.add_format({'border': 1, 'num_format': '#,##0.00', 'align': 'right'})
    
    # Write Headers
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_fmt)
    
    # Write Data
    for row_idx, row in enumerate(df.itertuples(index=False), start=1):
        for col_idx, value in enumerate(row):
            worksheet.write(row_idx, col_idx, value, str_fmt if col_idx < 2 else num_fmt)
    
    # Set Widths
    worksheet.set_column('A:A', 15)
    worksheet.set_column('B:B', 45)
    worksheet.set_column('C:H', 18)

if uploaded_file:
    with st.status("🚀 Generating Report... Please wait.", expanded=True) as status:
        st.write("📂 1. Reading file...")
        progress_bar = st.progress(0)
        try:
            df = pd.read_excel(uploaded_file)
            progress_bar.progress(25)
            
            st.write("🧹 2. Cleaning data...")
            df['Posting Date'] = pd.to_datetime(df['Posting Date'], errors='coerce')
            df['Payment date'] = pd.to_datetime(df['Payment date'], errors='coerce')
            df['Amount'] = pd.to_numeric(df['Amount in local currency'], errors='coerce').fillna(0)
            df['Supplier'] = df['Supplier'].fillna('N/A')
            df['Vendor name'] = df['Vendor name'].fillna('Unknown')
            df['G/L Account'] = df['G/L Account'].apply(lambda x: str(int(float(x))) if not pd.isna(x) and str(x).replace('.','').isdigit() else str(x))
            progress_bar.progress(50)
            
            st.write("🧮 3. Calculating Aging...")
            report_date = df['Posting Date'].max()
            buckets = ["Not Due", "1-30 Days", "31-60 Days", "61-90 Days", "90+ Days"]
            
            def get_bucket(pay_date):
                if pd.isna(pay_date): return "Not Due"
                days = (report_date - pay_date).days
                if days < 0: return "Not Due"
                elif days <= 30: return "1-30 Days"
                elif days <= 60: return "31-60 Days"
                elif days <= 90: return "61-90 Days"
                else: return "90+ Days"
            
            df['Aging Bucket'] = df['Payment date'].apply(get_bucket)
            progress_bar.progress(75)
            
            st.write("📊 4. Creating tables...")
            def create_pivot(d):
                if d.empty: return pd.DataFrame()
                p = d.pivot_table(index=['Supplier', 'Vendor name'], columns='Aging Bucket', values='Amount', aggfunc='sum', fill_value=0).reindex(columns=buckets, fill_value=0)
                p['Total Balance'] = p.sum(axis=1)
                return p.sort_values(by='Total Balance', ascending=True).reset_index()

            ap_pivot = create_pivot(df[df['G/L Account'].isin(['16740100', '31210100'])])
            dp_pivot = create_pivot(df[df['G/L Account'] == '16740100'])
            p_only = df[df['G/L Account'] == '31210100']
            db_pivot = pd.DataFrame()
            if not p_only.empty:
                temp = create_pivot(p_only)
                db_pivot = temp[temp['Total Balance'] > 0]
            
            progress_bar.progress(90)
            output_excel = io.BytesIO()
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                write_optimized_excel(writer, ap_pivot, 'AP Aging')
                write_optimized_excel(writer, dp_pivot, 'Downpayments')
                write_optimized_excel(writer, db_pivot, 'Debit Balances')
            
            progress_bar.progress(100)
            status.update(label="✅ Analysis Complete!", state="complete", expanded=False)
            
        except Exception as e:
            st.error(f"Error: {e}")
            st.stop()

    with col_h2:
        st.download_button("📥 Download Excel Report", output_excel.getvalue(), f"AP_Report_{datetime.now().strftime('%d%m%Y')}.xlsx")
        components.html(f"""<button onclick="window.print()" class="print-btn">🖨️ Save as PDF</button>""", height=50)

    def display_narrow_table(pivot_df, title):
        c1, c2, c3 = st.columns([1, 6, 1]) 
        with c2:
            st.markdown(f"### {title}")
            if pivot_df.empty:
                st.warning("No data found.")
                return
            
            total_cols = pivot_df[buckets].sum()
            grand_total = pivot_df['Total Balance'].sum()
            summ_data = {"Unit": ["kEGP", "kEUR"], "Total": [round(grand_total/1000), round((grand_total/eur_rate)/1000)]}
            for b in buckets:
                val = total_cols[b]
                summ_data[b] = [round(val/1000), round((val/eur_rate)/1000)]
            
            st.table(pd.DataFrame(summ_data).style.format(precision=0, thousands=",").set_properties(**{'text-align': 'center'}).set_properties(subset=['Total'], **{'font-weight': 'bold', 'background-color': '#f1f5f9'}).set_table_styles([{'selector': 'th', 'props': [('background-color', '#e2e8f0'), ('color', '#1e3a8a')]}]))

    st.divider()
    display_narrow_table(ap_pivot, "1. Total AP Aging Summary")
    display_narrow_table(dp_pivot, "2. Prepayments (DP) Summary")
    display_narrow_table(db_pivot, "3. Debit Balances Summary")

    st.markdown(f"""<div class="footer">© {datetime.now().year} | <b>Account Payable Intelligence Suite</b><br>Developed by <b>Can Adiguzel</b></div>""", unsafe_allow_html=True)
else:
    st.info("👋 Welcome! Please upload your FBL1N Excel file from the sidebar to start.")
