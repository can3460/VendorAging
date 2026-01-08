import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import os
import time
import streamlit.components.v1 as components

# --- Page Configuration ---
st.set_page_config(page_title="VendorFace | AP Analysis", layout="wide")

# --- Custom Styling ---
st.markdown("""
    <style>
    .main { background-color: #f8fafc; }
    h1 { color: #1e3a8a; font-family: 'Segoe UI', sans-serif; font-weight: 800; letter-spacing: -0.5px; }
    h3 { color: #1e40af; border-left: 5px solid #3b82f6; padding-left: 10px; margin-top: 30px; }
    .footer { text-align: center; color: #94a3b8; font-family: 'Courier New', monospace; font-size: 0.8rem; margin-top: 50px; border-top: 1px solid #e2e8f0; padding-top: 20px; }
    
    /* Table Headers Styling */
    thead tr th:first-child { display:none }
    tbody th { display:none }
    .stTable { width: 100% !important; }
    
    /* Print Button Style */
    .print-btn {
        background-color: #475569; color: white; border: none; padding: 10px 20px;
        text-align: center; text-decoration: none; display: inline-block;
        font-size: 16px; margin: 4px 2px; cursor: pointer; border-radius: 8px;
        font-family: 'Segoe UI', sans-serif; font-weight: bold;
        transition: 0.3s; width: 100%;
    }
    .print-btn:hover { background-color: #334155; transform: scale(1.02); }
    
    /* Print Mode Settings (Save as PDF) */
    @media print {
        [data-testid="stSidebar"] { display: none; }
        .stButton { display: none; }
        .print-hide { display: none; }
        iframe { display: none; } /* Hide the button itself */
        .footer { position: fixed; bottom: 0; width: 100%; text-align: center; }
        .block-container { padding-top: 1rem; } 
    }
    </style>
    """, unsafe_allow_html=True)

# --- Sidebar ---
with st.sidebar:
    if os.path.exists("logo.png"):
        st.image("logo.png", use_container_width=True)
    else:
        st.markdown("<div style='text-align: center; font-size: 50px;'>🛡️</div>", unsafe_allow_html=True)
    
    st.header("Settings")
    eur_rate = st.number_input("EUR / EGP Rate", value=52.50, step=0.01)
    st.divider()
    uploaded_file = st.file_uploader("Upload SAP FBL1N Report", type=["xlsx", "xls"])

# --- Main Page Header ---
col_h1, col_h2 = st.columns([3, 1])
with col_h1:
    st.title("Account Payable Aging Analysis")
    st.caption("Financial Intelligence Dashboard | Vendor Liquidity Tracking")

# --- Helper Function: Optimized Excel Writer ---
def write_optimized_excel(writer, df, sheet_name):
    """Writes data with formatting APPLIED ONLY TO DATA CELLS to save memory."""
    workbook = writer.book
    worksheet = workbook.add_worksheet(sheet_name)
    writer.sheets[sheet_name] = worksheet

    # Define Formats
    header_fmt = workbook.add_format({
        'bold': True, 'bg_color': '#1e3a8a', 'font_color': 'white', 
        'border': 1, 'align': 'center', 'valign': 'vcenter'
    })
    
    str_fmt = workbook.add_format({'border': 1, 'align': 'left'})
    num_fmt = workbook.add_format({'border': 1, 'num_format': '#,##0.00', 'align': 'right'})

    # 1. Write Header
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_fmt)

    # 2. Write Data (Iterating rows to apply format ONLY to data cells)
    # Using itertuples is faster than iterrows
    for row_idx, row in enumerate(df.itertuples(index=False), start=1):
        for col_idx, value in enumerate(row):
            # First 2 columns are Strings (Supplier, Name), rest are Numbers
            if col_idx < 2:
                worksheet.write(row_idx, col_idx, value, str_fmt)
            else:
                worksheet.write(row_idx, col_idx, value, num_fmt)

    # 3. Set Column Widths (No format applied here to keep empty cells clean)
    worksheet.set_column('A:A', 15)  # Supplier ID
    worksheet.set_column('B:B', 45)  # Vendor Name
    worksheet.set_column('C:H', 18)  # Amount Columns

# --- Logic ---
if uploaded_file:
    # --- PROCESS START ---
    with st.status("🚀 Generating Report... Please wait.", expanded=True) as status:
        
        st.write("📂 1. Reading file...")
        progress_bar = st.progress(0)
        
        try:
            df = pd.read_excel(uploaded_file)
            progress_bar.progress(25)
            time.sleep(0.3)

            st.write("🧹 2. Cleaning and standardizing data...")
            df['Posting Date'] = pd.to_datetime(df['Posting Date'], errors='coerce')
            df['Payment date'] = pd.to_datetime(df['Payment date'], errors='coerce')
            df['Amount'] = pd.to_numeric(df['Amount in local currency'], errors='coerce').fillna(0)
            df['Supplier'] = df['Supplier'].fillna('N/A')
            df['Vendor name'] = df['Vendor name'].fillna('Unknown')
            
            def clean_gl(val):
                try: return str(int(float(val)))
                except: return str(val).strip()
            df['G/L Account'] = df['G/L Account'].apply(clean_gl)
            progress_bar.progress(50)

            st.write("🧮 3. Calculating Aging buckets...")
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

            st.write("📊 4. Creating pivot tables...")
            
            # Pivot Functions
            def create_detailed_pivot(dataframe):
                if dataframe.empty: return pd.DataFrame()
                pivot = dataframe.pivot_table(
                    index=['Supplier', 'Vendor name'],
                    columns='Aging Bucket',
                    values='Amount',
                    aggfunc='sum', fill_value=0
                ).reindex(columns=buckets, fill_value=0)
                pivot['Total Balance'] = pivot.sum(axis=1)
                return pivot.sort_values(by='Total Balance', ascending=True).reset_index()

            ap_pivot = create_detailed_pivot(df[df['G/L Account'].isin(['16740100', '31210100'])])
            dp_pivot = create_detailed_pivot(df[df['G/L Account'] == '16740100'])
            
            p_only_raw = df[df['G/L Account'] == '31210100']
            if not p_only_raw.empty:
                p_only_pivot_temp = create_detailed_pivot(p_only_raw)
                db_pivot = p_only_pivot_temp[p_only_pivot_temp['Total Balance'] > 0]
            else:
                db_pivot = pd.DataFrame()
            
            progress_bar.progress(90)
            st.write("💾 5. Optimizing Excel file size...")

            # Optimized Excel Creation
            output_excel = io.BytesIO()
            with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
                write_optimized_excel(writer, ap_pivot, 'AP Aging Detail')
                write_optimized_excel(writer, dp_pivot, 'Downpayments Detail')
                write_optimized_excel(writer, db_pivot, 'Debit Balances Detail')
            
            progress_bar.progress(100)
            status.update(label="✅ Analysis Complete!", state="complete", expanded=False)

        except Exception as e:
            st.error(f"Error occurred: {e}")
            st.stop()

    # --- BUTTONS AREA (Top Right) ---
    with col_h2:
        # Excel Download Button
        st.download_button(
            label="📥 Download Excel Report",
            data=output_excel.getvalue(),
            file_name=f"AP_Analysis_Detailed_{datetime.now().strftime('%d%m%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        # HTML/JS Print Button (Save as PDF)
        components.html(
            f"""<button onclick="window.print()" class="print-btn">🖨️ Save Report as PDF</button>""", 
            height=50
        )

    # --- DISPLAY TABLES ---
    def display_narrow_table(pivot_df, title):
        c1, c2, c3 = st.columns([1, 6, 1]) 
        with c2:
            st.markdown(f"### {title}")
            if pivot_df.empty:
                st.warning("No data found in this category.")
                return

            total_cols = pivot_df[buckets].sum()
            grand_total = pivot_df['Total Balance'].sum()
            
            summary_data = {
                "Unit": ["kEGP", "kEUR"],
                "Total": [round(grand_total/1000), round((grand_total/eur_rate)/1000)]
            }
            for b in buckets:
                val = total_cols[b]
                summary_data[b] = [round(val/1000), round((val/eur_rate)/1000)]
            
            summ_df = pd.DataFrame(summary_data)
            
            st.table(
                summ_df.style.format(precision=0, thousands=",")
                .set_properties(**{'text-align': 'center'})
                .set_properties(subset=['Total'], **{'font-weight': 'bold', 'background-color': '#f1f5f9'})
                .set_table_styles([{'selector': 'th', 'props': [('background-color', '#e2e8f0'), ('color', '#1e3a8a'), ('font-weight', 'bold')]}])
            )

    st.divider()
    display_narrow_table(ap_pivot, "1. Total AP Aging Summary")
    display_narrow_table(dp_pivot, "2. Prepayments (DP) Summary")
    display_narrow_table(db_pivot, "3. Debit Balances Summary")

    # --- Footer ---
    st.markdown(f"""
        <div class="footer">
            © {datetime.now().year} | <b>Account Payable Intelligence Suite</b><br>
            Developed by <b>Can Adiguzel</b> & Powered by <b>Gemini Flash 2.0</b>
        </div>
        """, unsafe_allow_html=True)

else:
    st.info("👋 Welcome! Please upload your FBL1N Excel file from the sidebar to start.")