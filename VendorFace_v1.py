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
st.set_page_config(page_title="Vendor 360° | Opella Finance", layout="wide", page_icon="🛡️")
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
        
        # --- LOGIN BAŞLIK ALANI ---
        st.markdown("""
        <div style='text-align: center;'>
            <h2 style='color:#1e293b; margin-bottom: 0px;'>Vendor 360° Intelligence</h2>
            <p style='color: #94a3b8; font-size: 13px; margin-top: 5px; font-style: italic;'>
                Developed by <b>Can Adiguzel</b> with <span style="background: -webkit-linear-gradient(45deg, #4285F4, #9B72CB); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-weight: bold;">Gemini AI</span> technologies
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        st.write("")
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

def process_tb_file(file, fbl1n_gl_summary):
    try:
        df_tb = pd.read_excel(file)
        
        acc_col = next((col for col in df_tb.columns if 'Account' in str(col) and 'Number' in str(col)), None)
        amt_col = next((col for col in df_tb.columns if 'Total' in str(col) and 'reporting' in str(col)), None)
        
        name_col = next((col for col in df_tb.columns if ('Text' in str(col) or 'Description' in str(col)) and 'B/S' in str(col)), None)
        if not name_col:
            name_col = next((col for col in df_tb.columns if 'Text' in str(col) and 'Account' not in str(col)), None)

        if not acc_col or not amt_col:
            return None, {}, "Error: Could not identify Account/Balance columns."

        df_tb = df_tb.dropna(subset=[acc_col])
        df_tb['GL_Account'] = df_tb[acc_col].astype(str).str.strip()
        df_tb = df_tb[df_tb['GL_Account'].str.match(r'^\d+$')] 
        
        gl_name_map = {}
        if name_col:
            gl_name_map = df_tb.set_index('GL_Account')[name_col].to_dict()

        tb_summary = df_tb.groupby('GL_Account')[amt_col].sum().reset_index()
        tb_summary.rename(columns={amt_col: 'TB_Balance'}, inplace=True)
        
        fbl1n_check = fbl1n_gl_summary.reset_index()[['G/L Account', 'Total Balance']]
        fbl1n_check.rename(columns={'G/L Account': 'GL_Account', 'Total Balance': 'FBL1n_Sum'}, inplace=True)
        
        merged = pd.merge(tb_summary, fbl1n_check, on='GL_Account', how='inner')
        merged['Difference'] = merged['TB_Balance'] - merged['FBL1n_Sum']
        
        return merged, gl_name_map, "Success"
        
    except Exception as e:
        return None, {}, str(e)

def write_optimized_excel(writer, df, sheet_name):
    if df.empty: return
    workbook = writer.book
    worksheet = workbook.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = worksheet
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#5b21b6', 'font_color': 'white', 'border': 1, 'align': 'center'})
    num_fmt = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
    txt_fmt = workbook.add_format({'border': 1})
    
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_fmt)
        worksheet.set_column(col_num, col_num, len(str(value)) + 5) 

    for row_idx, row in enumerate(df.itertuples(index=False), start=1):
        for col_idx, value in enumerate(row):
            worksheet.write(row_idx, col_idx, value, num_fmt if isinstance(value, (int, float)) else txt_fmt)
            
    worksheet.set_column(0, 0, 15) 
    worksheet.set_column(1, 1, 35)

# ==========================================
# 4. SIDEBAR & NAVIGATION
# ==========================================
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo
