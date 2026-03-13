"""
╔══════════════════════════════════════════════════════════════════════════════╗
║                    VENDORFACE v7.1 PERFECT HYBRID                            ║
║                                                                              ║
║        Original v1 Motor + Claude UX + Opella Branding                       ║
║        Bulletproof Excel | Live FX | Dual Theme | Production Ready           ║
║                                                                              ║
║        Opella Healthcare Finance Operations | Internal Use Only             ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, date
import io
import os
import plotly.express as px
import plotly.graph_objects as go

# Live FX rates
try:
    import yfinance as yf
    YF_AVAILABLE = True
except ImportError:
    YF_AVAILABLE = False

# ══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="AP Analyzing Suite | Opella",
    layout="wide",
    page_icon="🛡️",
    initial_sidebar_state="expanded",
    menu_items={'Get Help': None, 'Report a bug': None, 'About': "VendorFace v7.1 | Opella Finance"}
)

# Hide Streamlit branding
st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
.stDeployButton {display:none;}
[data-testid="stToolbar"] {display: none;}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS & STATE
# ══════════════════════════════════════════════════════════════════════════════
VERSION_NO = "v7.1 HYBRID"
MASTER_ADMIN = "can.adiguzel@sanofi.com"
USER_DB = "users.xlsx"

# Opella brand colors
OPELLA_COLORS = {
    'primary': '#00A3E0',    # Bright blue
    'secondary': '#7C3AED',  # Purple
    'accent': '#F59E0B',     # Amber
    'success': '#10B981',    # Green
    'danger': '#EF4444',     # Red
}

if 'logged_in' not in st.session_state: st.session_state['logged_in'] = False
if 'view_currency' not in st.session_state: st.session_state['view_currency'] = "Local"
if 'results' not in st.session_state: st.session_state['results'] = None
if 'analysis_run' not in st.session_state: st.session_state['analysis_run'] = False
if 'theme' not in st.session_state: st.session_state['theme'] = 'light'
if 'eur_rate' not in st.session_state: st.session_state['eur_rate'] = 1.0

# ══════════════════════════════════════════════════════════════════════════════
# THEME SYSTEM WITH OPELLA COLORS
# ══════════════════════════════════════════════════════════════════════════════
def get_theme():
    if st.session_state['theme'] == 'dark':
        return {
            'bg': '#0F172A',
            'card_bg': '#1E293B',
            'text': '#F1F5F9',
            'text_sec': '#CBD5E1',
            'border': '#334155',
            'accent': OPELLA_COLORS['primary'],
            'sidebar_bg': 'linear-gradient(135deg, #00A3E0 0%, #7C3AED 100%)',
            'input_bg': '#1E293B',
            'input_text': '#F1F5F9'
        }
    return {
        'bg': '#F8FAFC',
        'card_bg': '#FFFFFF',
        'text': '#0F172A',
        'text_sec': '#475569',
        'border': '#E2E8F0',
        'accent': OPELLA_COLORS['primary'],
        'sidebar_bg': 'linear-gradient(135deg, #00A3E0 0%, #7C3AED 100%)',
        'input_bg': '#FFFFFF',
        'input_text': '#0F172A'
    }

theme = get_theme()

# ══════════════════════════════════════════════════════════════════════════════
# CSS WITH OPELLA BRANDING + DARK MODE FIX
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

* {{
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
}}

html, body, [data-testid="stAppViewContainer"] {{
    background: {theme['bg']};
    color: {theme['text']};
}}

/* Sidebar with Opella gradient */
[data-testid="stSidebar"] {{
    background: {theme['sidebar_bg']};
}}

[data-testid="stSidebar"] * {{
    color: #F1F5F9 !important;
}}

[data-testid="stSidebar"] label {{
    color: #E0E7FF !important;
    font-size: 0.75rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}}

/* Input fields in sidebar - CRITICAL FIX */
[data-testid="stSidebar"] input,
[data-testid="stSidebar"] select,
[data-testid="stSidebar"] textarea {{
    background: rgba(255, 255, 255, 0.15) !important;
    color: #FFFFFF !important;
    border: 1px solid rgba(255, 255, 255, 0.3) !important;
    border-radius: 8px !important;
    padding: 8px 12px !important;
    font-size: 0.9rem !important;
    font-weight: 600 !important;
}}

[data-testid="stSidebar"] input::placeholder {{
    color: rgba(255, 255, 255, 0.5) !important;
}}

/* Number input arrows */
[data-testid="stSidebar"] input[type="number"] {{
    color: #FFFFFF !important;
}}

/* Selectbox dropdown text */
[data-testid="stSidebar"] [data-baseweb="select"] {{
    color: #FFFFFF !important;
}}

[data-testid="stSidebar"] [data-baseweb="select"] > div {{
    background: rgba(255, 255, 255, 0.15) !important;
    color: #FFFFFF !important;
    border: 1px solid rgba(255, 255, 255, 0.3) !important;
}}

/* KPI Cards with Opella colors */
.kpi-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
    gap: 20px;
    margin-bottom: 24px;
}}

.kpi-card {{
    background: {theme['card_bg']};
    border: 1px solid {theme['border']};
    border-radius: 16px;
    padding: 24px;
    box-shadow: 0 4px 12px rgba(0,0,0,0.08);
    transition: all 0.3s ease;
    position: relative;
    overflow: hidden;
}}

.kpi-card:hover {{
    transform: translateY(-4px);
    box-shadow: 0 8px 24px rgba(0,0,0,0.15);
}}

.kpi-card::before {{
    content: '';
    position: absolute;
    top: 0;
    left: 0;
    width: 4px;
    height: 100%;
    border-radius: 16px 0 0 16px;
}}

.kpi-card.opella-blue::before {{ background: {OPELLA_COLORS['primary']}; }}
.kpi-card.opella-purple::before {{ background: {OPELLA_COLORS['secondary']}; }}
.kpi-card.opella-amber::before {{ background: {OPELLA_COLORS['accent']}; }}
.kpi-card.opella-green::before {{ background: {OPELLA_COLORS['success']}; }}
.kpi-card.opella-red::before {{ background: {OPELLA_COLORS['danger']}; }}

.kpi-icon {{
    font-size: 2.5rem;
    margin-bottom: 12px;
    opacity: 0.9;
}}

.kpi-label {{
    font-size: 0.8rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    color: {theme['text_sec']};
    margin-bottom: 8px;
}}

.kpi-value {{
    font-size: 2.2rem;
    font-weight: 800;
    color: {theme['text']};
    line-height: 1.2;
    margin-bottom: 6px;
}}

.kpi-sub {{
    font-size: 0.9rem;
    color: {theme['text_sec']};
    font-weight: 500;
}}

/* Info Boxes */
.info-box {{
    background: rgba(0, 163, 224, 0.1);
    border: 1px solid rgba(0, 163, 224, 0.3);
    border-radius: 12px;
    padding: 16px;
    margin: 12px 0;
    font-size: 0.95rem;
    line-height: 1.6;
    color: {theme['text']};
}}

.success-box {{
    background: rgba(16, 185, 129, 0.1);
    border: 1px solid rgba(16, 185, 129, 0.3);
}}

.warning-box {{
    background: rgba(245, 158, 11, 0.1);
    border: 1px solid rgba(245, 158, 11, 0.3);
}}

/* Header with Opella gradient */
.main-header {{
    background: linear-gradient(135deg, {OPELLA_COLORS['primary']} 0%, {OPELLA_COLORS['secondary']} 100%);
    padding: 36px;
    border-radius: 20px;
    margin-bottom: 28px;
    color: white;
    box-shadow: 0 12px 40px rgba(0, 163, 224, 0.25);
}}

.header-title {{
    font-size: 2.4rem;
    font-weight: 800;
    margin-bottom: 6px;
    text-shadow: 0 2px 4px rgba(0,0,0,0.1);
}}

.header-subtitle {{
    font-size: 1.1rem;
    opacity: 0.95;
    font-weight: 500;
}}

/* Dark mode text fixes - ENHANCED */
.stMarkdown, .stMarkdown p, .stMarkdown div, 
[data-testid="stMarkdownContainer"],
[data-testid="stText"],
.element-container,
.stTextInput label,
.stSelectbox label,
.stNumberInput label {{
    color: {theme['text']} !important;
}}

/* Metric labels with better contrast */
[data-testid="stMetricLabel"] {{
    color: {theme['text_sec']} !important;
    font-weight: 600 !important;
}}

[data-testid="stMetricValue"] {{
    color: {theme['text']} !important;
    font-weight: 800 !important;
}}

/* Section headers with Opella accent */
.sec-hdr {{
    font-size: 1.2rem;
    font-weight: 700;
    color: {theme['text']};
    margin: 28px 0 14px;
    padding-bottom: 10px;
    border-bottom: 3px solid {OPELLA_COLORS['primary']};
}}

/* Buttons with Opella colors */
.stButton > button {{
    border-radius: 8px;
    font-weight: 600;
    border: none;
    transition: all 0.3s ease;
}}

.stButton > button[kind="primary"] {{
    background: {OPELLA_COLORS['primary']};
    color: white;
}}

.stButton > button[kind="primary"]:hover {{
    background: {OPELLA_COLORS['secondary']};
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(0, 163, 224, 0.3);
}}

/* Dataframe styling */
[data-testid="stDataFrame"] {{
    border-radius: 12px;
    overflow: hidden;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
}}

/* Tab styling */
.stTabs [data-baseweb="tab-list"] {{
    gap: 8px;
}}

.stTabs [data-baseweb="tab"] {{
    border-radius: 8px 8px 0 0;
    padding: 12px 20px;
    font-weight: 600;
}}

.stTabs [aria-selected="true"] {{
    background: {OPELLA_COLORS['primary']};
    color: white;
}}
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════════════════
# AUTHENTICATION (Original v1)
# ══════════════════════════════════════════════════════════════════════════════
def load_users():
    if not os.path.exists(USER_DB):
        default_users = pd.DataFrame([
            {"email": MASTER_ADMIN, "role": "admin"},
            {"email": "can.adiguzel@opella.com", "role": "admin"},
            {"email": "admin@opella.com", "role": "admin"}
        ])
        default_users.to_excel(USER_DB, index=False)
        return default_users
    df = pd.read_excel(USER_DB)
    df.columns = [str(c).lower().strip() for c in df.columns]
    return df

def add_user(new_email):
    users = load_users()
    new_email = new_email.strip().lower()
    if 'email' in users.columns and new_email not in users['email'].values:
        new_row = pd.DataFrame([{"email": new_email, "role": "user"}])
        users = pd.concat([users, new_row], ignore_index=True)
        users.to_excel(USER_DB, index=False)
        return True
    return False

# ══════════════════════════════════════════════════════════════════════════════
# CORE ENGINE (Original v1 - PRESERVED WITH BUG FIXES)
# ══════════════════════════════════════════════════════════════════════════════
def get_live_rate(base_currency):
    """Live FX rates from Yahoo Finance"""
    if base_currency == "EUR": 
        return 1.0
    if not YF_AVAILABLE:
        return None
    try:
        ticker = yf.Ticker(f"EUR{base_currency}=X")
        history = ticker.history(period="1d")
        if not history.empty:
            return float(history['Close'].iloc[-1])
        return None
    except Exception: 
        return None

def smart_read(file):
    """Multi-encoding file reader"""
    name = file.name.lower()
    if name.endswith((".xlsx", ".xls")): 
        return pd.read_excel(file)
    raw = file.getvalue()
    for enc in ["utf-8-sig", "utf-8", "iso-8859-9", "cp1254", "latin-1", "windows-1252"]:
        try:
            return pd.read_csv(io.BytesIO(raw), encoding=enc, sep=None, engine="python", on_bad_lines="skip")
        except: 
            continue
    raise ValueError(f"Could not read '{file.name}'. Please ensure it is an Excel or UTF-8 CSV file.")

def smart_parse_tb(file):
    """Trial Balance parser with SOLAR code support"""
    try:
        df_tb = smart_read(file)
        gl_name_map, gl_solar_map, gl_balance_map = {}, {}, {}
        
        acc_col = next((c for c in df_tb.columns if 'Account Number' in str(c) or 'G/L' in str(c) or 'Account' in str(c)), None)
        name_col = next((c for c in df_tb.columns if 'Text' in str(c) or 'Description' in str(c)), None)
        solar_col = next((c for c in df_tb.columns if any(x in str(c).lower() for x in ['financial', 'fs item', 'solar', 'group'])), None)
        amt_col = next((c for c in df_tb.columns if any(x in str(c).lower() for x in ['total', 'balance', 'reporting'])), None)

        if not acc_col or not amt_col: 
            return {}, {}, {}

        for _, row in df_tb.iterrows():
            raw_val = str(row[acc_col]).strip()
            clean_acc = raw_val.split('.')[0]
            if clean_acc.isdigit() and len(clean_acc) >= 6:
                gl_name_map[clean_acc] = str(row[name_col]).strip() if name_col and not pd.isna(row[name_col]) else "-"
                gl_solar_map[clean_acc] = str(row[solar_col]).strip() if solar_col and not pd.isna(row[solar_col]) else "-"
                gl_balance_map[clean_acc] = row[amt_col] if not pd.isna(row[amt_col]) else 0
        return gl_name_map, gl_solar_map, gl_balance_map
    except Exception as e: 
        st.error(f"Error reading Trial Balance: {e}")
        return {}, {}, {}

def append_totals(df, numeric_cols, label_col='Vendor'):
    """Append TOTAL row to DataFrame"""
    if df is None or df.empty: 
        return df
    tot_dict = {c: df[c].sum() for c in numeric_cols if c in df.columns}
    tot_dict[label_col] = 'TOTAL'
    tot_df = pd.DataFrame([tot_dict])
    return pd.concat([df, tot_df], ignore_index=True)

def generate_html_report(dfs, titles, display_curr, rate):
    """HTML report generator"""
    html = f"<html><head><style>body{{font-family:sans-serif;padding:20px;}}h2{{color:#00A3E0;border-bottom:2px solid #00A3E0;padding-bottom:5px;}}table{{border-collapse:collapse;width:100%;margin-bottom:30px;font-size:12px;}}th{{background:#f1f5f9;padding:10px;border:1px solid #cbd5e1;text-align:right;}}td{{padding:8px;border:1px solid #cbd5e1;text-align:right;}}td:first-child, th:first-child{{text-align:left;font-weight:bold;}}</style></head><body>"
    html += f"<h1>Opella AP Analyzing Suite</h1><p><b>Currency View:</b> {display_curr} | <b>Date:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')} | <b>EUR Rate:</b> {rate:,.4f}</p>"
    for df, title in zip(dfs, titles):
        if df is not None and not df.empty:
            html += f"<h2>{title}</h2>"
            df_fmt = df.copy()
            for col in df_fmt.select_dtypes(include=[np.number]).columns:
                df_fmt[col] = df_fmt[col].apply(lambda x: f"{x:,.0f}" if not pd.isna(x) else "")
            html += df_fmt.to_html(index=False)
    html += "</body></html>"
    return html

# ══════════════════════════════════════════════════════════════════════════════
# BULLETPROOF EXCEL ENGINE (Original v1)
# ══════════════════════════════════════════════════════════════════════════════
def format_excel_sheet(writer, df, sheet_name):
    """Bulletproof Excel formatter"""
    if df is None or df.empty:
        df = pd.DataFrame({'Data': ['No data available in this category.']})
        
    df = df.copy()
    
    for col in df.select_dtypes(include=['datetimetz']).columns:
        df[col] = df[col].dt.tz_localize(None)
        
    df = df.replace([np.inf, -np.inf], np.nan)
    df.to_excel(writer, sheet_name=sheet_name, index=False, na_rep="")
    
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    header_format = workbook.add_format({
        'bold': True,
        'font_color': '#FFFFFF',
        'bg_color': '#00A3E0',
        'border': 1
    })
    cell_format = workbook.add_format({'border': 1})
    num_format = workbook.add_format({'border': 1, 'num_format': '#,##0'})
    
    for col_num, col_name in enumerate(df.columns):
        worksheet.write(0, col_num, col_name, header_format)
        max_len = max(df[col_name].astype(str).apply(len).max(), len(str(col_name)))
        worksheet.set_column(col_num, col_num, min(max_len + 2, 50))
        
        if df[col_name].dtype in [np.int64, np.float64]:
            for row_num in range(1, len(df) + 1):
                val = df[col_name].iloc[row_num - 1]
                if pd.notna(val) and not np.isinf(val):
                    worksheet.write(row_num, col_num, val, num_format)
                else:
                    worksheet.write(row_num, col_num, "", cell_format)
        else:
            for row_num in range(1, len(df) + 1):
                val = df[col_name].iloc[row_num - 1]
                worksheet.write(row_num, col_num, val, cell_format)

def build_excel_output(results):
    """Build complete Excel workbook"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        format_excel_sheet(writer, results['aging_matrix'], 'Aging Matrix')
        format_excel_sheet(writer, results['summary_vendor'], 'Summary by Vendor')
        format_excel_sheet(writer, results['vendor_aging'], 'Vendor Aging')
        format_excel_sheet(writer, results['gl_breakdown'], 'GL Breakdown')
        format_excel_sheet(writer, results['prepayments'], 'Prepayments')
        format_excel_sheet(writer, results['debit_balances'], 'Debit Balances')
        format_excel_sheet(writer, results['tb_reconciliation'], 'TB Reconciliation')
        format_excel_sheet(writer, results['raw_data'], 'Raw Data')
    return output.getvalue()

# ══════════════════════════════════════════════════════════════════════════════
# ANALYSIS ENGINE WITH DYNAMIC COLUMN DETECTION (BUG FIX)
# ══════════════════════════════════════════════════════════════════════════════
def analyze_ap(df, gl_solar_map, display_curr, eur_rate):
    """Main AP analysis engine with smart column detection"""
    
    if df.empty:
        return None
    
    df = df.copy()
    
    # ✅ CRITICAL FIX: Dynamic column detection
    col_map = {}
    for c in df.columns:
        cl = str(c).lower()
        if 'amount' in cl and 'local' in cl:
            col_map['Amount'] = c
        elif 'payment date' in cl or 'due date' in cl or 'net due' in cl:
            col_map['Due Date'] = c
        elif 'document date' in cl or 'posting date' in cl:
            col_map['Document Date'] = c
        elif 'vendor name' in cl or 'name1' in cl:
            col_map['Vendor Name'] = c
        elif 'supplier' in cl or 'vendor' in cl:
            col_map['Vendor'] = c
        elif 'g/l' in cl or 'gl account' in cl:
            col_map['GL Account'] = c
        elif 'document number' in cl or 'doc' in cl:
            col_map['Document Number'] = c
    
    # Rename columns
    df.rename(columns={v: k for k, v in col_map.items()}, inplace=True)
    
    # Set defaults for missing columns
    if 'Amount' not in df.columns:
        st.error("❌ Could not find 'Amount in local currency' column")
        return None
    if 'Due Date' not in df.columns:
        df['Due Date'] = pd.Timestamp.today()
    if 'Vendor Name' not in df.columns:
        df['Vendor Name'] = df.get('Vendor', 'Unknown')
    if 'GL Account' not in df.columns:
        df['GL Account'] = '160000'
    if 'Document Number' not in df.columns:
        df['Document Number'] = range(len(df))
    
    # Amount conversion
    if display_curr == "EUR":
        df['View_Amount'] = df['Amount'] / eur_rate
    else:
        df['View_Amount'] = df['Amount']
    
    # Date parsing
    df['Due Date'] = pd.to_datetime(df['Due Date'], errors='coerce')
    if 'Document Date' in df.columns:
        df['Document Date'] = pd.to_datetime(df['Document Date'], errors='coerce')
    
    # Days calculation
    today = pd.Timestamp.today()
    df['Days Overdue'] = (today - df['Due Date']).dt.days
    df['Days Overdue'] = df['Days Overdue'].clip(lower=0)
    
    # Aging buckets
    def assign_bucket(days):
        if pd.isna(days) or days <= 0: return 'Current'
        elif days <= 30: return '1-30 Days'
        elif days <= 60: return '31-60 Days'
        elif days <= 90: return '61-90 Days'
        else: return '90+ Days'
    
    df['Aging Bucket'] = df['Days Overdue'].apply(assign_bucket)
    
    # GL mapping
    df['GL_6'] = df['GL Account'].astype(str).str[:6]
    df['SOLAR'] = df['GL_6'].map(gl_solar_map).fillna('-')
    
    # Segment classification
    def classify_segment(row):
        solar = str(row['SOLAR']).strip()
        gl = str(row['GL_6']).strip()
        vname = str(row.get('Vendor Name', '')).lower()
        
        if solar == '42905' or gl.startswith(('160100', '161', '162')):
            return 'ICO'
        if any(kw in vname for kw in ['interco', 'ico', 'sanofi', 'opella']):
            return 'ICO'
        
        if solar == '42006' or gl.startswith(('165', '163')):
            return 'Employee'
        if any(kw in vname for kw in ['employee', 'personnel', 'travel']):
            return 'Employee'
        
        return '3rd Party'
    
    df['Segment'] = df.apply(classify_segment, axis=1)
    
    # Aging matrix
    aging_cols = ['Current', '1-30 Days', '31-60 Days', '61-90 Days', '90+ Days']
    aging_matrix = pd.pivot_table(
        df,
        values='View_Amount',
        index='Vendor Name',
        columns='Aging Bucket',
        aggfunc='sum',
        fill_value=0
    )
    for col in aging_cols:
        if col not in aging_matrix.columns:
            aging_matrix[col] = 0
    aging_matrix = aging_matrix[aging_cols]
    aging_matrix['Total'] = aging_matrix.sum(axis=1)
    aging_matrix = aging_matrix.sort_values('Total', ascending=False)
    aging_matrix = append_totals(aging_matrix, aging_cols + ['Total'], 'Vendor Name')
    
    # Summary by vendor
    summary_vendor = df.groupby('Vendor Name').agg({
        'View_Amount': 'sum',
        'Document Number': 'count',
        'Days Overdue': 'max'
    }).reset_index()
    summary_vendor.columns = ['Vendor', 'Balance', 'Invoice Count', 'Max Days Overdue']
    summary_vendor = summary_vendor.sort_values('Balance', ascending=False)
    summary_vendor = append_totals(summary_vendor, ['Balance', 'Invoice Count'], 'Vendor')
    
    # GL breakdown
    gl_breakdown = df.groupby('GL_6').agg({
        'View_Amount': 'sum',
        'Document Number': 'count'
    }).reset_index()
    gl_breakdown.columns = ['GL Account', 'Balance', 'Count']
    gl_breakdown = gl_breakdown.sort_values('Balance', ascending=False)
    gl_breakdown = append_totals(gl_breakdown, ['Balance', 'Count'], 'GL Account')
    
    # Vendor aging
    vendor_aging = df.groupby(['Vendor Name', 'Aging Bucket']).agg({
        'View_Amount': 'sum'
    }).reset_index()
    vendor_aging.columns = ['Vendor', 'Bucket', 'Amount']
    
    # Prepayments
    prepayments = df[df['View_Amount'] > 0].copy()
    
    # Debit balances
    vendor_net = df.groupby('Vendor Name')['View_Amount'].sum()
    debit_vendors = vendor_net[vendor_net > 0].index
    debit_balances = df[df['Vendor Name'].isin(debit_vendors)].copy()
    
    return {
        'aging_matrix': aging_matrix,
        'summary_vendor': summary_vendor,
        'gl_breakdown': gl_breakdown,
        'vendor_aging': vendor_aging,
        'prepayments': prepayments,
        'debit_balances': debit_balances,
        'tb_reconciliation': pd.DataFrame(),
        'raw_data': df
    }

# ══════════════════════════════════════════════════════════════════════════════
# LOGIN PAGE
# ══════════════════════════════════════════════════════════════════════════════
if not st.session_state['logged_in']:
    st.markdown(f"""
    <div class='main-header'>
        <div class='header-title'>🛡️ AP Analyzing Suite</div>
        <div class='header-subtitle'>Opella Healthcare Finance Operations | {VERSION_NO}</div>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 1.5, 1])
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        with st.form("login_form"):
            email_input = st.text_input("Corporate Email (@sanofi.com / @opella.com)").strip().lower()
            if st.form_submit_button("🔐 Secure Login", use_container_width=True):
                if not (email_input.endswith("@sanofi.com") or email_input.endswith("@opella.com")):
                    st.error("❌ Only @sanofi.com and @opella.com emails are allowed.")
                else:
                    users = load_users()
                    if 'email' in users.columns and email_input in users['email'].str.lower().values:
                        st.session_state['logged_in'] = True
                        st.session_state['user_email'] = email_input
                        st.session_state['user_name'] = email_input.split('@')[0].replace('.', ' ').title()
                        st.rerun()
                    else:
                        st.error("❌ Access denied. Contact Can Adiguzel for authorization.")
    
    st.markdown(f"""
    <div class='info-box' style='text-align:center;margin-top:40px;'>
        🔒 <b>Zero Data Retention</b> | All data processed in temporary memory only<br/>
        No files stored. Session data deleted on browser close.
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# MAIN APP
# ══════════════════════════════════════════════════════════════════════════════

# Sidebar
with st.sidebar:
    st.markdown(f"""
    <div style='text-align:center;padding:24px 0;'>
        <div style='font-size:3rem;margin-bottom:10px;'>🛡️</div>
        <div style='font-size:1.4rem;font-weight:800;color:#FFFFFF;'>Opella</div>
        <div style='font-size:0.75rem;color:#E0E7FF;letter-spacing:0.1em;margin-top:4px;'>AP SUITE {VERSION_NO}</div>
    </div>
    <hr style='border-color:rgba(255,255,255,0.2);margin:16px 0;'/>
    """, unsafe_allow_html=True)
    
    st.markdown(f"""
    <div style='background:rgba(255,255,255,0.1);border:1px solid rgba(255,255,255,0.2);
                border-radius:12px;padding:14px;margin-bottom:16px;'>
      <div style='font-size:.7rem;color:#E0E7FF;margin-bottom:6px;font-weight:600;'>LOGGED IN AS</div>
      <div style='font-size:1rem;font-weight:700;color:#FFFFFF;'>
        {st.session_state.get('user_name', 'User')}
      </div>
      <div style='font-size:.75rem;color:#CBD5E1;margin-top:4px;'>
        {st.session_state.get('user_email', '')}
      </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Theme toggle
    col_t1, col_t2 = st.columns([3, 1])
    with col_t2:
        if st.button("🌙" if st.session_state['theme'] == 'light' else "☀️", 
                     help="Toggle theme", use_container_width=True):
            st.session_state['theme'] = 'dark' if st.session_state['theme'] == 'light' else 'light'
            st.rerun()
    
    st.markdown("### 📁 Upload Files")
    fbl1n_file = st.file_uploader("FBL1N — AP Line Items", type=["xlsx", "xls", "csv"], key="fbl1n")
    tb_file = st.file_uploader("F.01 — Trial Balance", type=["xlsx", "xls", "csv"], key="tb")
    
    st.markdown("---")
    st.markdown("### 💱 Currency")
    
    view_currency = st.selectbox(
        "Display Currency",
        ["Local", "EUR"],
        index=0 if st.session_state['view_currency'] == "Local" else 1,
        key="curr_select"
    )
    st.session_state['view_currency'] = view_currency
    
    if view_currency == "Local":
        st.info("💡 Showing in local currency")
        eur_rate = 1.0
    else:
        base_curr = "TRY"
        if fbl1n_file:
            try:
                temp_df = smart_read(fbl1n_file)
                curr_col = next((c for c in temp_df.columns if 'currency' in str(c).lower()), None)
                if curr_col:
                    base_curr = str(temp_df[curr_col].iloc[0]).strip()
            except:
                pass
        
        col_r, col_b = st.columns([3, 1])
        with col_r:
            eur_rate = st.number_input(
                f"1 EUR = ? {base_curr}",
                min_value=0.00001,
                value=float(st.session_state.get('eur_rate', 1.0)),
                format="%.4f",
                key="rate_input"
            )
            st.session_state['eur_rate'] = eur_rate
        with col_b:
            if st.button("🌐", help=f"Fetch live rate", use_container_width=True, key="fetch_rate"):
                if YF_AVAILABLE:
                    with st.spinner("..."):
                        live_rate = get_live_rate(base_curr)
                        if live_rate:
                            st.session_state['eur_rate'] = live_rate
                            st.success(f"✅ {live_rate:.4f}")
                            st.rerun()
                        else:
                            st.error("❌ Failed")
                else:
                    st.error("yfinance not installed")
    
    st.markdown("---")
    
    if st.button("🔄 Run Analysis", use_container_width=True, type="primary", key="run_btn"):
        if fbl1n_file:
            st.session_state['analysis_run'] = True
            st.rerun()
        else:
            st.error("❌ Upload FBL1N first")
    
    st.markdown("---")
    
    if st.button("🚪 Logout", use_container_width=True, key="logout_btn"):
        st.session_state['logged_in'] = False
        st.session_state['analysis_run'] = False
        st.session_state['results'] = None
        st.rerun()

# Main content
st.markdown(f"""
<div class='main-header'>
    <div style='display:flex;justify-content:space-between;align-items:center;'>
        <div>
            <div class='header-title'>📊 AP Intelligence Dashboard</div>
            <div class='header-subtitle'>Real-time Analysis • Powered by Opella Finance</div>
        </div>
        <div style='text-align:right;'>
            <div style='font-size:0.85rem;opacity:0.9;'>
                {datetime.now().strftime('%d %b %Y, %H:%M')}
            </div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

if st.session_state.get('analysis_run') and fbl1n_file:
    with st.spinner("🔄 Processing data..."):
        try:
            df = smart_read(fbl1n_file)
            
            gl_name_map, gl_solar_map, gl_balance_map = {}, {}, {}
            if tb_file:
                gl_name_map, gl_solar_map, gl_balance_map = smart_parse_tb(tb_file)
            
            results = analyze_ap(df, gl_solar_map, view_currency, eur_rate)
            
            if results:
                st.session_state['results'] = results
                
                # KPI Cards with Opella colors
                total_ap = results['summary_vendor']['Balance'].iloc[-1] if not results['summary_vendor'].empty else 0
                total_invoices = int(results['summary_vendor']['Invoice Count'].iloc[-1]) if not results['summary_vendor'].empty else 0
                total_vendors = len(results['summary_vendor']) - 1
                
                overdue = results['aging_matrix'][['1-30 Days', '31-60 Days', '61-90 Days', '90+ Days']].iloc[-1].sum() if not results['aging_matrix'].empty else 0
                critical = results['aging_matrix']['90+ Days'].iloc[-1] if not results['aging_matrix'].empty else 0
                
                st.markdown(f"""
                <div class='kpi-grid'>
                    <div class='kpi-card opella-blue'>
                        <div class='kpi-icon'>💰</div>
                        <div class='kpi-label'>Total AP Balance</div>
                        <div class='kpi-value'>{total_ap:,.0f}</div>
                        <div class='kpi-sub'>{total_invoices:,} invoices • {total_vendors} vendors</div>
                    </div>
                    <div class='kpi-card opella-red'>
                        <div class='kpi-icon'>⏰</div>
                        <div class='kpi-label'>Overdue Amount</div>
                        <div class='kpi-value'>{overdue:,.0f}</div>
                        <div class='kpi-sub'>{overdue/total_ap*100 if total_ap else 0:.1f}% of total AP</div>
                    </div>
                    <div class='kpi-card opella-amber'>
                        <div class='kpi-icon'>⚠️</div>
                        <div class='kpi-label'>Critical 90+ Days</div>
                        <div class='kpi-value'>{critical:,.0f}</div>
                        <div class='kpi-sub'>{critical/total_ap*100 if total_ap else 0:.1f}% risk exposure</div>
                    </div>
                </div>
                """, unsafe_allow_html=True)
                
                # Tabs
                tab1, tab2, tab3, tab4 = st.tabs([
                    "📊 Aging Analysis",
                    "👥 Vendor Summary",
                    "🏦 GL Breakdown",
                    "📤 Export"
                ])
                
                with tab1:
                    st.markdown('<div class="sec-hdr">Aging Matrix</div>', unsafe_allow_html=True)
                    st.dataframe(results['aging_matrix'], use_container_width=True, height=400)
                    
                    if not results['vendor_aging'].empty:
                        top20 = results['vendor_aging'].groupby('Vendor')['Amount'].sum().nlargest(20).reset_index()
                        vendor_aging_top = results['vendor_aging'][results['vendor_aging']['Vendor'].isin(top20['Vendor'])]
                        
                        fig = px.bar(
                            vendor_aging_top,
                            x='Vendor',
                            y='Amount',
                            color='Bucket',
                            title='Top 20 Vendors by Aging',
                            template='plotly_white',
                            color_discrete_map={
                                'Current': OPELLA_COLORS['success'],
                                '1-30 Days': OPELLA_COLORS['primary'],
                                '31-60 Days': OPELLA_COLORS['accent'],
                                '61-90 Days': '#EF4444',
                                '90+ Days': OPELLA_COLORS['danger']
                            }
                        )
                        fig.update_layout(height=500)
                        st.plotly_chart(fig, use_container_width=True)
                
                with tab2:
                    st.markdown('<div class="sec-hdr">Vendor Summary</div>', unsafe_allow_html=True)
                    st.dataframe(results['summary_vendor'], use_container_width=True, height=400)
                
                with tab3:
                    st.markdown('<div class="sec-hdr">GL Account Breakdown</div>', unsafe_allow_html=True)
                    st.dataframe(results['gl_breakdown'], use_container_width=True, height=400)
                
                with tab4:
                    st.markdown('<div class="sec-hdr">Export Options</div>', unsafe_allow_html=True)
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        excel_data = build_excel_output(results)
                        st.download_button(
                            "📥 Download Excel Report",
                            excel_data,
                            file_name=f"AP_Analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                    
                    with col2:
                        html_data = generate_html_report(
                            [results['aging_matrix'], results['summary_vendor'], results['gl_breakdown']],
                            ['Aging Matrix', 'Vendor Summary', 'GL Breakdown'],
                            view_currency,
                            eur_rate
                        )
                        st.download_button(
                            "📥 Download HTML Report",
                            html_data.encode('utf-8'),
                            file_name=f"AP_Analysis_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
                            mime="text/html",
                            use_container_width=True
                        )
                
        except Exception as e:
            st.error(f"❌ Analysis error: {e}")
            import traceback
            with st.expander("🔍 Technical Details"):
                st.code(traceback.format_exc())

else:
    st.markdown(f"""
    <div class='info-box' style='text-align:center;padding:60px;'>
        <div style='font-size:3.5rem;margin-bottom:20px;'>📂</div>
        <h2 style='color:{theme['text']};margin-bottom:12px;font-weight:700;'>Upload Files to Begin</h2>
        <p style='color:{theme['text_sec']};font-size:1.1rem;'>FBL1N & F.01 files required for analysis</p>
    </div>
    """, unsafe_allow_html=True)

# Footer
st.markdown(f"""
<div style='text-align:center;margin-top:60px;padding-top:28px;border-top:2px solid {theme['border']};
            color:{theme['text_sec']};font-size:0.85rem;'>
    🛡️ <b>Opella AP Suite {VERSION_NO}</b> | Finance Operations | Zero Data Retention<br/>
    <span style='font-size:0.75rem;opacity:0.8;'>Powered by Streamlit • Secured by Enterprise Auth</span>
</div>
""", unsafe_allow_html=True)
