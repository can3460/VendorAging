"""
╔══════════════════════════════════════════════════════════════════════════════╗
║                          VendorFace v6.0 ULTIMATE                            ║
║                   AP Intelligence Dashboard | Opella Finance                 ║
║                  Perfect Engine + Stunning UX | Production Ready             ║
╚══════════════════════════════════════════════════════════════════════════════╝
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from io import BytesIO
import io, json, warnings, urllib.request, os, base64
from datetime import datetime, date, timedelta

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════════════════════════════════════
st.set_page_config(
    page_title="VendorFace v6.0 | AP Intelligence",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={'Get Help': None, 'Report a bug': None, 'About': "VendorFace v6.0 | Opella Finance"}
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
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════
AGING_LABELS = ["Current", "1-30 Days", "31-60 Days", "61-90 Days", "90+ Days"]
AGING_COLORS = ["#10B981", "#3B82F6", "#F59E0B", "#EF4444", "#991B1B"]
AGING_BG = ["rgba(16,185,129,0.1)", "rgba(59,130,246,0.1)", "rgba(245,158,11,0.1)", 
            "rgba(239,68,68,0.1)", "rgba(153,27,27,0.1)"]

CURRENCIES = {
    "EUR": "EUR", "USD": "USD", "GBP": "GBP", "TRY": "TRY", "EGP": "EGP", "TND": "TND",
    "AED": "AED", "SAR": "SAR", "CNY": "CNY", "JPY": "JPY", "INR": "INR", "BRL": "BRL"
}

GL_NAMES = {
    "160000": "AP Trade", "160100": "AP Interco", "160200": "AP Services",
    "168000": "AP Accruals", "165000": "Employee Pay", "163000": "Travel Accrual"
}

ICO_GL = ("160100", "161", "162")
EMP_GL = ("165", "163")
ICO_KW = ["interco", "ico", "affiliated", "related party", "intragroup", "group payable"]
EMP_KW = ["employee", "personnel", "staff", "travel", "expense", "salary", "bonus"]

def is_ico(gl, vname=""):
    gl, vn = str(gl).strip(), f" {str(vname).lower()} "
    return gl.startswith(ICO_GL) or any(f" {k} " in vn for k in ICO_KW) or "intercompany" in vn

def is_employee(gl, vname=""):
    gl, vn = str(gl).strip(), f" {str(vname).lower()} "
    return gl.startswith(EMP_GL) or any(f" {k} " in vn for k in EMP_KW)


# ══════════════════════════════════════════════════════════════════════════════
# THEME SYSTEM (Light/Dark Mode)
# ══════════════════════════════════════════════════════════════════════════════
if 'theme' not in st.session_state:
    st.session_state.theme = 'light'

def get_theme():
    if st.session_state.theme == 'dark':
        return {
            'bg': '#0F172A', 'card_bg': '#1E293B', 'text': '#F1F5F9', 
            'text_sec': '#94A3B8', 'border': '#334155', 'accent': '#3B82F6'
        }
    return {
        'bg': '#F8FAFC', 'card_bg': '#FFFFFF', 'text': '#0F172A', 
        'text_sec': '#64748B', 'border': '#E2E8F0', 'accent': '#2563EB'
    }

theme = get_theme()


# ══════════════════════════════════════════════════════════════════════════════
# GLASSMORPHISM CSS
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

[data-testid="stSidebar"] {{
    background: linear-gradient(180deg, #1e3a8a 0%, #312e81 100%);
}}

[data-testid="stSidebar"] * {{
    color: #F1F5F9 !important;
}}

/* Glassmorphism Cards */
.glass-card {{
    background: rgba(255, 255, 255, 0.05);
    backdrop-filter: blur(10px);
    border: 1px solid rgba(255, 255, 255, 0.1);
    border-radius: 16px;
    padding: 24px;
    box-shadow: 0 8px 32px rgba(0, 0, 0, 0.1);
    transition: all 0.3s ease;
}}

.glass-card:hover {{
    transform: translateY(-4px);
    box-shadow: 0 12px 48px rgba(0, 0, 0, 0.15);
}}

/* KPI Cards */
.kpi-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
    gap: 16px;
    margin-bottom: 24px;
}}

.kpi-card {{
    background: {theme['card_bg']};
    border: 1px solid {theme['border']};
    border-radius: 12px;
    padding: 20px;
    position: relative;
    overflow: hidden;
    transition: all 0.3s ease;
}}

.kpi-card:hover {{
    transform: translateY(-2px);
    box-shadow: 0 8px 24px rgba(0,0,0,0.12);
}}

.kpi-card::before {{
    content: '';
    position: absolute;
    top: 0;
    left: 0;
    width: 4px;
    height: 100%;
    background: linear-gradient(180deg, var(--accent), transparent);
}}

.kpi-label {{
    font-size: 0.75rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    color: {theme['text_sec']};
    margin-bottom: 8px;
}}

.kpi-value {{
    font-size: 2rem;
    font-weight: 800;
    color: {theme['text']};
    line-height: 1.2;
    margin-bottom: 4px;
}}

.kpi-sub {{
    font-size: 0.85rem;
    color: {theme['text_sec']};
}}

.kpi-badge {{
    position: absolute;
    top: 16px;
    right: 16px;
    background: rgba(59, 130, 246, 0.1);
    color: #3B82F6;
    font-size: 0.7rem;
    font-weight: 700;
    padding: 4px 10px;
    border-radius: 999px;
}}

/* Progress Bars */
.progress-bar {{
    height: 8px;
    background: {theme['border']};
    border-radius: 999px;
    overflow: hidden;
    margin-top: 12px;
}}

.progress-fill {{
    height: 100%;
    background: linear-gradient(90deg, #3B82F6, #8B5CF6);
    transition: width 0.6s ease;
}}

/* Aging Pills */
.pill {{
    display: inline-block;
    padding: 4px 12px;
    border-radius: 999px;
    font-size: 0.75rem;
    font-weight: 600;
    margin: 2px;
}}

.pill-current {{ background: #D1FAE5; color: #065F46; }}
.pill-30 {{ background: #DBEAFE; color: #1E40AF; }}
.pill-60 {{ background: #FEF3C7; color: #92400E; }}
.pill-90 {{ background: #FEE2E2; color: #991B1B; }}
.pill-120p {{ background: #FEE2E2; color: #7F1D1D; border: 2px solid #991B1B; }}

/* Section Headers */
.sec-hdr {{
    font-size: 1.1rem;
    font-weight: 700;
    color: {theme['text']};
    margin: 24px 0 12px;
    padding-bottom: 8px;
    border-bottom: 2px solid {theme['accent']};
}}

/* Info Boxes */
.info-box {{
    background: rgba(59, 130, 246, 0.1);
    border: 1px solid rgba(59, 130, 246, 0.3);
    border-radius: 12px;
    padding: 16px;
    margin: 12px 0;
    font-size: 0.9rem;
    line-height: 1.6;
}}

.warning-box {{
    background: rgba(239, 68, 68, 0.1);
    border: 1px solid rgba(239, 68, 68, 0.3);
}}

.success-box {{
    background: rgba(16, 185, 129, 0.1);
    border: 1px solid rgba(16, 185, 129, 0.3);
}}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════
def assign_aging(days):
    if days <= 0: return "Current"
    elif days <= 30: return "1-30 Days"
    elif days <= 60: return "31-60 Days"
    elif days <= 90: return "61-90 Days"
    else: return "90+ Days"

def fa(val):
    """Format Amount"""
    curr = st.session_state.get("disp_curr", "EUR")
    kk = st.session_state.get("in_k", False)
    if pd.isna(val): return "—"
    v = abs(float(val)) / (1000 if kk else 1)
    return f"{v:,.0f} {'k' if kk else ''}{curr}"

def smart_read(file):
    """Encoding-safe file reader"""
    name = file.name.lower()
    if name.endswith((".xlsx", ".xls")):
        return pd.read_excel(file)
    raw = file.getvalue()
    for enc in ["utf-8-sig", "utf-8", "iso-8859-9", "cp1254", "latin-1", "cp1252"]:
        try:
            return pd.read_csv(io.BytesIO(raw), encoding=enc, sep=None, 
                             engine="python", on_bad_lines="skip")
        except:
            continue
    raise ValueError(f"Could not read '{file.name}'")


# ══════════════════════════════════════════════════════════════════════════════
# 🔥 PERFECT DATA LOADER (Bug Fix: Correct Amount Column)
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def load_fbl1n(file):
    """Perfect FBL1N loader with correct amount mapping"""
    df = smart_read(file)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()].copy()
    
    rename, mapped = {}, set()
    
    for c in df.columns:
        cl, target = c.lower(), None
        
        # ✅ CRITICAL: Amount mapping priority
        if "amount in local currency" in cl:
            target = "Amount (LC)"
        elif "vendor name" in cl or "name1" in cl:
            target = "Vendor Name"
        elif "supplier" in cl or "vendor" in cl:
            target = "Vendor"
        elif "payment date" in cl or "due date" in cl or "vade" in cl:
            target = "Due Date"
        elif "document date" in cl or "belge tarihi" in cl:
            target = "Document Date"
        elif "g/l account" in cl or "gl account" in cl:
            target = "GL Account"
        elif "document number" in cl:
            target = "Document No"
        elif "company code" in cl:
            target = "Company Code"
        
        if target and target not in mapped:
            rename[c] = target
            mapped.add(target)
    
    df.rename(columns=rename, inplace=True)
    
    # Defaults
    defaults = {
        "Amount (LC)": 0,
        "Vendor": "Unknown",
        "Due Date": pd.Timestamp(date.today()),
        "GL Account": "160000",
    }
    for col, val in defaults.items():
        if col not in df.columns:
            df[col] = val
    
    # Type conversions
    df["Due Date"] = pd.to_datetime(df["Due Date"], errors="coerce").fillna(pd.Timestamp(date.today()))
    df["Amount (LC)"] = pd.to_numeric(df["Amount (LC)"], errors="coerce").fillna(0)
    
    today = pd.Timestamp(date.today())
    df["Days Overdue"] = (today - df["Due Date"]).dt.days.clip(lower=0)
    df["Aging Bucket"] = df["Days Overdue"].apply(assign_aging)
    df["GL Account"] = df["GL Account"].astype(str).str.strip()
    
    # Segment classification
    vn = "Vendor Name" if "Vendor Name" in df.columns else "Vendor"
    df["Segment"] = df.apply(
        lambda r: "ICO" if is_ico(r["GL Account"], r[vn])
        else ("Employee" if is_employee(r["GL Account"], r[vn]) else "3rd Party"),
        axis=1
    )
    
    return df.loc[:, ~df.columns.duplicated()].copy()


@st.cache_data(show_spinner=False)
def load_f01(file):
    """F.01 Trial Balance loader"""
    df = smart_read(file)
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()].copy()
    
    rename, mapped = {}, set()
    for c in df.columns:
        cl, target = c.lower(), None
        if "g/l account" in cl or "gl account" in cl: target = "GL Account"
        elif "balance" in cl or "saldo" in cl: target = "Balance"
        elif "fs item" in cl or "solar" in cl: target = "SOLAR"
        elif "description" in cl or "text" in cl: target = "Description"
        
        if target and target not in mapped:
            rename[c] = target
            mapped.add(target)
    
    df.rename(columns=rename, inplace=True)
    
    for col, val in [("GL Account", "Unknown"), ("Balance", 0)]:
        if col not in df.columns: df[col] = val
    
    df["Balance"] = pd.to_numeric(df["Balance"], errors="coerce").fillna(0)
    df["GL Account"] = df["GL Account"].astype(str).str.strip()
    
    if "SOLAR" not in df.columns: df["SOLAR"] = ""
    df["SOLAR"] = df["SOLAR"].astype(str).str.strip()
    
    if "Description" not in df.columns:
        df["Description"] = df["GL Account"]
    
    return df.loc[:, ~df.columns.duplicated()].copy()


# ══════════════════════════════════════════════════════════════════════════════
# DEMO DATA
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def demo_fbl1n(n=500):
    np.random.seed(42)
    today = pd.Timestamp(date.today())
    vendors = [f"V{str(i).zfill(6)}" for i in range(1001, 1101)]
    vnames = ["Siemens AG", "Bosch GmbH", "SAP SE", "BASF SE", "Henkel AG",
              "Sanofi EMEA (ICO)", "Sanofi US (ICO)", "Group Treasury (ICO)",
              "John Smith (Employee)", "Travel Expense Pool"]
    vmap = {v: vnames[i % len(vnames)] for i, v in enumerate(vendors)}
    
    gl_pool = [("160000", 0.35), ("160100", 0.20), ("168000", 0.18), 
               ("165000", 0.15), ("163000", 0.12)]
    gls, gprob = zip(*gl_pool)
    gprob = np.array(gprob) / sum(gprob)
    
    rows = []
    for i in range(n):
        v = np.random.choice(vendors)
        gl = np.random.choice(gls, p=gprob)
        offset = np.random.choice([-5, 0, 15, 35, 55, 80, 110], 
                                  p=[0.15, 0.20, 0.20, 0.15, 0.12, 0.10, 0.08])
        due = today - pd.Timedelta(days=int(offset))
        amt = np.random.choice([-1, 1]) * round(np.random.lognormal(9, 1.5), 2)
        
        rows.append({
            "Vendor": v,
            "Vendor Name": vmap[v],
            "GL Account": gl,
            "Due Date": due,
            "Amount (LC)": amt,
            "Document No": f"DOC{np.random.randint(1000000, 9999999)}",
            "Company Code": np.random.choice(["DE01", "FR01", "TR01", "EG03"])
        })
    
    df = pd.DataFrame(rows)
    df["Days Overdue"] = (today - df["Due Date"]).dt.days.clip(lower=0)
    df["Aging Bucket"] = df["Days Overdue"].apply(assign_aging)
    df["Segment"] = df.apply(
        lambda r: "ICO" if is_ico(r["GL Account"], r["Vendor Name"])
        else ("Employee" if is_employee(r["GL Account"], r["Vendor Name"]) else "3rd Party"),
        axis=1
    )
    return df


@st.cache_data(show_spinner=False)
def demo_f01():
    return pd.DataFrame([
        {"GL Account": "160000", "Balance": -3500000, "SOLAR": "40000", "Description": "AP Trade"},
        {"GL Account": "160100", "Balance": -2200000, "SOLAR": "42905", "Description": "AP Interco"},
        {"GL Account": "165000", "Balance": -450000, "SOLAR": "42006", "Description": "Employee Pay"},
    ])


# ══════════════════════════════════════════════════════════════════════════════
# RECONCILIATION
# ══════════════════════════════════════════════════════════════════════════════
def reconcile(fbl1n, f01):
    """SOLAR-aware GL reconciliation"""
    f01_pay = f01[f01["SOLAR"].isin(["40000", "42905", "42006"])].copy()
    
    ap = fbl1n.groupby("GL Account")["Amount (LC)"].sum().reset_index()
    ap.columns = ["GL Account", "AP Subledger"]
    
    m = pd.merge(
        f01_pay[["GL Account", "Description", "Balance", "SOLAR"]],
        ap, on="GL Account", how="outer"
    ).fillna(0)
    
    m["Difference"] = m["Balance"] - m["AP Subledger"]
    m["Match"] = m["Difference"].abs() < 1.0
    
    matched = m[m["Match"]].copy()
    gaps = m[~m["Match"] & (m["AP Subledger"] != 0)].copy()
    missing = m[(m["Balance"] != 0) & (m["AP Subledger"] == 0)].copy()
    
    return matched, gaps, missing


# ══════════════════════════════════════════════════════════════════════════════
# 🔥 PERFECT EXCEL EXPORT (Bug Fix)
# ══════════════════════════════════════════════════════════════════════════════
def build_excel(df_full, recon_match, recon_gap, recon_miss, aging_summary):
    """
    Perfect Excel export - uses FULL dataset, no filtering
    
    BUG FIX: Previous version was exporting filtered df, 
    causing mismatch between dashboard totals and Excel totals
    """
    try:
        from openpyxl import Workbook
        from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        # Fallback to basic export
        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            df_full.to_excel(w, sheet_name="AP Line Items", index=False)
            aging_summary.to_excel(w, sheet_name="Aging Summary", index=False)
            recon_match.to_excel(w, sheet_name="Reconciliation Matched", index=False)
            recon_gap.to_excel(w, sheet_name="Reconciliation Gaps", index=False)
            recon_miss.to_excel(w, sheet_name="Missing GL", index=False)
        return buf.getvalue()
    
    wb = Workbook()
    
    # Styles
    HF = PatternFill("solid", fgColor="1A56DB")
    HFT = Font(bold=True, color="FFFFFF", size=10)
    BD = Side(style="thin", color="E2E8F0")
    CBR = Border(left=BD, right=BD, top=BD, bottom=BD)
    
    AGING_FILLS = {
        "Current": PatternFill("solid", fgColor="D1FAE5"),
        "1-30 Days": PatternFill("solid", fgColor="DBEAFE"),
        "31-60 Days": PatternFill("solid", fgColor="FEF3C7"),
        "61-90 Days": PatternFill("solid", fgColor="FEE2E2"),
        "90+ Days": PatternFill("solid", fgColor="FEE2E2"),
    }
    
    def hrow(ws, nc):
        for c in range(1, nc + 1):
            cell = ws.cell(row=1, column=c)
            if cell.value:
                cell.fill = HF
                cell.font = HFT
                cell.border = CBR
                cell.alignment = Alignment(horizontal="center", vertical="center")
    
    def aw(ws):
        for col in ws.columns:
            mx = max((len(str(c.value or "")) for c in col), default=8)
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(mx + 4, 50)
    
    # ─── Cover Sheet ───
    ws0 = wb.active
    ws0.title = "Cover"
    ws0["B2"] = "VendorFace v6.0 — AP Intelligence Report"
    ws0["B2"].font = Font(bold=True, size=18, color="1A56DB")
    ws0["B3"] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws0["B3"].font = Font(size=11, color="64748B")
    ws0["B5"] = f"Total Records: {len(df_full):,}"
    ws0["B6"] = f"Total Amount: {df_full['Amount (LC)'].sum():,.2f}"
    ws0.column_dimensions["B"].width = 50
    
    # ─── 1. AP Line Items (FULL DATASET) ───
    ws1 = wb.create_sheet("AP Line Items")
    cols = [c for c in ["Document No", "Vendor", "Vendor Name", "GL Account", "Segment",
                         "Due Date", "Days Overdue", "Aging Bucket", "Amount (LC)", 
                         "Company Code"] if c in df_full.columns]
    
    ws1.append(cols)
    hrow(ws1, len(cols))
    ws1.freeze_panes = "A2"
    
    # ✅ CRITICAL: Use full dataset
    for _, row in df_full[cols].iterrows():
        ws1.append(list(row))
        r = ws1.max_row
        bk = row.get("Aging Bucket", "Current")
        fill = AGING_FILLS.get(bk, PatternFill())
        
        for ci in range(1, len(cols) + 1):
            cell = ws1.cell(row=r, column=ci)
            cell.border = CBR
            cell.fill = fill
            if cols[ci - 1] == "Amount (LC)":
                cell.number_format = '#,##0.00'
    
    aw(ws1)
    
    # ─── 2. Aging Summary ───
    ws2 = wb.create_sheet("Aging Summary")
    ws2.append(list(aging_summary.columns))
    hrow(ws2, len(aging_summary.columns))
    
    for _, row in aging_summary.iterrows():
        ws2.append(list(row))
        r = ws2.max_row
        bk = row.get("Aging Bucket", "Current")
        fill = AGING_FILLS.get(bk, PatternFill())
        
        for ci in range(1, len(aging_summary.columns) + 1):
            cell = ws2.cell(row=r, column=ci)
            cell.border = CBR
            cell.fill = fill
            if "Amount" in aging_summary.columns[ci - 1]:
                cell.number_format = '#,##0.00'
    
    aw(ws2)
    
    # ─── 3. Vendor Summary ───
    ws3 = wb.create_sheet("Vendor Summary")
    vc = ["Vendor", "Vendor Name", "Segment"] if "Vendor Name" in df_full.columns else ["Vendor", "Segment"]
    vdf = (df_full.groupby(vc)
           .agg(Balance=("Amount (LC)", "sum"),
                Count=("Amount (LC)", "count"),
                MaxOverdue=("Days Overdue", "max"))
           .reset_index()
           .sort_values("Balance"))
    
    ws3.append(list(vdf.columns))
    hrow(ws3, len(vdf.columns))
    
    for _, row in vdf.iterrows():
        ws3.append(list(row))
        r = ws3.max_row
        for ci in range(1, len(vdf.columns) + 1):
            cell = ws3.cell(row=r, column=ci)
            cell.border = CBR
            if vdf.columns[ci - 1] == "Balance":
                cell.number_format = '#,##0.00'
    
    aw(ws3)
    
    # ─── 4. Reconciliation ───
    ws4 = wb.create_sheet("Reconciliation Matched")
    ws4.append(list(recon_match.columns))
    hrow(ws4, len(recon_match.columns))
    for _, row in recon_match.iterrows():
        ws4.append(list(row))
    aw(ws4)
    
    ws5 = wb.create_sheet("Reconciliation Gaps")
    ws5.append(list(recon_gap.columns))
    hrow(ws5, len(recon_gap.columns))
    for _, row in recon_gap.iterrows():
        ws5.append(list(row))
    aw(ws5)
    
    ws6 = wb.create_sheet("Missing GL")
    ws6.append(list(recon_miss.columns))
    hrow(ws6, len(recon_miss.columns))
    for _, row in recon_miss.iterrows():
        ws6.append(list(row))
    aw(ws6)
    
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(f"""
    <div style='text-align:center;padding:20px 0;'>
        <div style='font-size:2.5rem;margin-bottom:8px;'>💎</div>
        <div style='font-size:1.3rem;font-weight:800;color:#F1F5F9;'>VendorFace</div>
        <div style='font-size:0.7rem;color:#94A3B8;letter-spacing:0.1em;'>v6.0 ULTIMATE</div>
    </div>
    <hr style='border-color:#475569;margin:16px 0;'/>
    """, unsafe_allow_html=True)
    
    # Theme Toggle
    col_t1, col_t2 = st.columns([3, 1])
    with col_t2:
        if st.button("🌙" if st.session_state.theme == 'light' else "☀️", 
                     help="Toggle theme", use_container_width=True):
            st.session_state.theme = 'dark' if st.session_state.theme == 'light' else 'light'
            st.rerun()
    
    st.markdown("### 📁 Data Source")
    use_demo = st.toggle("Demo Mode", value=True)
    fbl1n_file, f01_file = None, None
    
    if not use_demo:
        fbl1n_file = st.file_uploader("FBL1N — AP Line Items", type=["xlsx", "xls", "csv"])
        f01_file = st.file_uploader("F.01 — Trial Balance", type=["xlsx", "xls", "csv"])
    
    st.markdown("<hr style='border-color:#475569;margin:16px 0;'/>", unsafe_allow_html=True)
    
    st.markdown("### 💱 Currency")
    cur_sel = st.selectbox("Local Currency", list(CURRENCIES.keys()), index=0)
    cur_code = CURRENCIES[cur_sel]
    in_k = st.toggle(f"Show in thousands (k{cur_code})", value=False)
    
    st.session_state["disp_curr"] = cur_code
    st.session_state["in_k"] = in_k
    
    st.markdown("<hr style='border-color:#475569;margin:16px 0;'/>", unsafe_allow_html=True)
    
    st.markdown("""
    <div style='background:rgba(59,130,246,0.1);border:1px solid rgba(59,130,246,0.3);
                border-radius:8px;padding:12px;font-size:0.75rem;line-height:1.5;'>
        🔒 <b>Zero Data Retention</b><br/>
        All files processed in temporary memory only. No data stored on servers.
    </div>
    """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# LOAD DATA
# ══════════════════════════════════════════════════════════════════════════════
if use_demo:
    fbl1n_df, f01_df, data_ok = demo_fbl1n(500), demo_f01(), True
else:
    data_ok, fbl1n_df, f01_df = False, None, demo_f01()
    
    if fbl1n_file:
        try:
            with st.spinner("📊 Processing FBL1N data..."):
                fbl1n_df = load_fbl1n(fbl1n_file)
            st.success("✅ FBL1N loaded successfully")
            data_ok = True
        except Exception as e:
            st.error(f"❌ Error loading FBL1N: {e}")
            st.stop()
    
    if f01_file:
        try:
            with st.spinner("📊 Processing F.01 data..."):
                f01_df = load_f01(f01_file)
        except Exception as e:
            st.warning(f"F.01 loading failed: {e}")

df_full = fbl1n_df.copy() if data_ok else pd.DataFrame()


# ══════════════════════════════════════════════════════════════════════════════
# HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div style='background:linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding:32px;border-radius:20px;margin-bottom:24px;color:white;
            box-shadow:0 20px 60px rgba(102,126,234,0.3);'>
    <div style='display:flex;justify-content:space-between;align-items:center;'>
        <div>
            <div style='font-size:2.2rem;font-weight:800;margin-bottom:4px;'>
                💎 VendorFace v6.0
            </div>
            <div style='font-size:1rem;opacity:0.9;'>
                AP Intelligence Dashboard | Opella Finance Operations
            </div>
        </div>
        <div style='text-align:right;'>
            <div style='font-size:0.8rem;opacity:0.8;'>
                {datetime.now().strftime('%d %b %Y, %H:%M')}
            </div>
            <div style='font-size:1.1rem;font-weight:600;margin-top:4px;'>
                {len(df_full):,} Records
            </div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

if use_demo:
    st.markdown("""
    <div class='info-box'>
        🎯 <b>Demo Mode Active</b> — Viewing synthetic data. 
        Upload real FBL1N & F.01 files via sidebar to analyze actual AP data.
    </div>
    """, unsafe_allow_html=True)

if not data_ok:
    st.markdown("""
    <div style='text-align:center;padding:80px;'>
        <div style='font-size:4rem;margin-bottom:16px;'>📂</div>
        <h2 style='color:#64748B;margin-bottom:8px;'>Upload Files to Begin</h2>
        <p style='color:#94A3B8;'>FBL1N & F.01 files required</p>
    </div>
    """, unsafe_allow_html=True)
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# KPI CARDS
# ══════════════════════════════════════════════════════════════════════════════
tot_ap = df_full["Amount (LC)"].sum()
ov_df = df_full[df_full["Days Overdue"] > 0]
tot_ov = ov_df["Amount (LC)"].sum()
crit_df = df_full[df_full["Aging Bucket"] == "90+ Days"]
n_vend = df_full["Vendor"].nunique()
ico_bal = df_full[df_full["Segment"] == "ICO"]["Amount (LC)"].sum()
emp_bal = df_full[df_full["Segment"] == "Employee"]["Amount (LC)"].sum()

kpi_data = [
    ("Total AP", tot_ap, f"{len(df_full):,} items · {n_vend} vendors", "#3B82F6", "💰"),
    ("Overdue", tot_ov, f"{len(ov_df):,} invoices", "#EF4444", "⏰"),
    ("Critical 90+", crit_df["Amount (LC)"].sum(), f"{len(crit_df):,} items", "#DC2626", "⚠️"),
    ("ICO Balance", ico_bal, "Intercompany", "#8B5CF6", "🔗"),
    ("Employee", emp_bal, "Staff payables", "#10B981", "👤"),
]

cols = st.columns(len(kpi_data))
for i, (label, value, sub, color, icon) in enumerate(kpi_data):
    with cols[i]:
        st.markdown(f"""
        <div class='kpi-card' style='--accent:{color};'>
            <div style='font-size:2rem;margin-bottom:8px;'>{icon}</div>
            <div class='kpi-label'>{label}</div>
            <div class='kpi-value'>{fa(value)}</div>
            <div class='kpi-sub'>{sub}</div>
            <div class='progress-bar'>
                <div class='progress-fill' style='width:{min(100, abs(value/tot_ap*100) if tot_ap else 0):.0f}%;background:{color};'></div>
            </div>
        </div>
        """, unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 Aging Analysis",
    "🔄 Reconciliation",
    "📈 Vendor Intelligence",
    "📤 Export & Reports"
])

with tab1:
    st.markdown('<div class="sec-hdr">Aging Distribution</div>', unsafe_allow_html=True)
    
    # Aging summary
    aging = (df_full.groupby("Aging Bucket")
             .agg(Count=("Amount (LC)", "count"),
                  **{"Total Amount": ("Amount (LC)", "sum")})
             .reset_index())
    aging["Aging Bucket"] = pd.Categorical(aging["Aging Bucket"], AGING_LABELS, ordered=True)
    aging = aging.sort_values("Aging Bucket")
    aging["Total Amount"] = aging["Total Amount"].abs()
    
    # 3D Bar Chart
    fig = go.Figure()
    for i, row in aging.iterrows():
        bk = row["Aging Bucket"]
        amt = row["Total Amount"] / (1000 if st.session_state.get("in_k") else 1)
        color = AGING_COLORS[AGING_LABELS.index(bk)]
        
        fig.add_trace(go.Bar(
            name=bk,
            x=[bk],
            y=[amt],
            marker_color=color,
            text=[f"{fa(row['Total Amount'])}<br>{int(row['Count']):,} items"],
            textposition="outside",
            hovertemplate=f"<b>{bk}</b><br>Amount: {fa(row['Total Amount'])}<br>Count: {int(row['Count']):,}<extra></extra>"
        ))
    
    fig.update_layout(
        showlegend=False,
        height=400,
        margin=dict(l=0, r=0, t=20, b=0),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        font=dict(color=theme['text']),
        yaxis=dict(
            gridcolor=theme['border'],
            title=f"Amount ({'k' if st.session_state.get('in_k') else ''}{st.session_state.get('disp_curr', 'EUR')})"
        )
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Segment breakdown
    st.markdown('<div class="sec-hdr">Segment Analysis</div>', unsafe_allow_html=True)
    
    seg_cols = st.columns(3)
    for i, seg in enumerate(["3rd Party", "ICO", "Employee"]):
        seg_df = df_full[df_full["Segment"] == seg]
        seg_amt = seg_df["Amount (LC)"].sum()
        
        with seg_cols[i]:
            st.markdown(f"""
            <div class='glass-card'>
                <div style='font-size:1.1rem;font-weight:700;margin-bottom:8px;'>
                    {seg}
                </div>
                <div style='font-size:1.8rem;font-weight:800;color:{AGING_COLORS[i]};margin-bottom:4px;'>
                    {fa(seg_amt)}
                </div>
                <div style='font-size:0.85rem;color:{theme['text_sec']};'>
                    {len(seg_df):,} items
                </div>
            </div>
            """, unsafe_allow_html=True)

with tab2:
    st.markdown('<div class="sec-hdr">GL Reconciliation</div>', unsafe_allow_html=True)
    
    rm, rg, rms = reconcile(df_full, f01_df)
    
    gap_tot = rg["Difference"].abs().sum() if not rg.empty else 0
    
    if gap_tot < 1 and rms.empty:
        st.markdown('<div class="success-box">✅ All payable GL accounts matched. No variances detected.</div>', 
                   unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="warning-box">⚠️ {len(rg)} GL(s) with variances. Total gap: {fa(gap_tot)}</div>',
                   unsafe_allow_html=True)
    
    if not rms.empty:
        st.markdown('<div class="sec-hdr">Missing GL Accounts in Sub-Ledger</div>', unsafe_allow_html=True)
        st.markdown("""
        <div class='info-box'>
            📋 These GL accounts have payable balances in the Trial Balance (F.01) but do not appear 
            in the AP Sub-Ledger report. Common causes: manual GL postings, accruals, or period-end adjustments.
        </div>
        """, unsafe_allow_html=True)
        
        # Chart
        fig = go.Figure(go.Bar(
            x=rms["Balance"].abs(),
            y=rms["GL Account"] + " — " + rms.get("Description", ""),
            orientation='h',
            marker_color='#EF4444',
            text=[fa(v) for v in rms["Balance"]],
            textposition='outside'
        ))
        fig.update_layout(
            height=max(300, len(rms) * 50),
            margin=dict(l=0, r=100, t=20, b=0),
            xaxis_title=f"Balance ({st.session_state.get('disp_curr', 'EUR')})",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color=theme['text'])
        )
        st.plotly_chart(fig, use_container_width=True)

with tab3:
    st.markdown('<div class="sec-hdr">Top 20 Vendors by Balance</div>', unsafe_allow_html=True)
    
    lbl = "Vendor Name" if "Vendor Name" in df_full.columns else "Vendor"
    top20 = (df_full.groupby(lbl)["Amount (LC)"]
             .sum().abs().sort_values(ascending=False).head(20).reset_index())
    
    fig = go.Figure(go.Bar(
        x=top20[lbl].str[:30],
        y=top20["Amount (LC)"] / (1000 if st.session_state.get("in_k") else 1),
        marker=dict(
            color=top20["Amount (LC)"],
            colorscale='RdYlGn_r',
            showscale=True
        ),
        text=[fa(v) for v in top20["Amount (LC)"]],
        textposition='outside'
    ))
    fig.update_layout(
        height=500,
        margin=dict(l=0, r=0, t=20, b=0),
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color=theme['text']),
        yaxis_title=f"Balance ({'k' if st.session_state.get('in_k') else ''}{st.session_state.get('disp_curr', 'EUR')})"
    )
    st.plotly_chart(fig, use_container_width=True)

with tab4:
    st.markdown('<div class="sec-hdr">Export & Reports</div>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        <div class='glass-card'>
            <div style='font-size:1.5rem;margin-bottom:8px;'>📊</div>
            <div style='font-weight:700;margin-bottom:8px;'>Excel Report</div>
            <div style='font-size:0.85rem;color:{};margin-bottom:16px;'>
                Complete workbook with all data
            </div>
        </div>
        """.format(theme['text_sec']), unsafe_allow_html=True)
        
        if st.button("⬇️ Generate Excel", use_container_width=True, type="primary"):
            with st.spinner("Building workbook..."):
                rm, rg, rms = reconcile(df_full, f01_df)
                xdata = build_excel(df_full, rm, rg, rms, aging)
            
            st.download_button(
                "📥 Download .xlsx",
                xdata,
                file_name=f"VendorFace_AP_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    
    with col2:
        st.markdown("""
        <div class='glass-card'>
            <div style='font-size:1.5rem;margin-bottom:8px;'>⚠️</div>
            <div style='font-weight:700;margin-bottom:8px;'>Critical Items</div>
            <div style='font-size:0.85rem;color:{};margin-bottom:16px;'>
                90+ day overdue only
            </div>
        </div>
        """.format(theme['text_sec']), unsafe_allow_html=True)
        
        crit = df_full[df_full["Aging Bucket"] == "90+ Days"]
        if len(crit):
            st.download_button(
                f"⬇️ Critical ({len(crit)} items)",
                crit.to_csv(index=False).encode("utf-8-sig"),
                file_name=f"VendorFace_Critical_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        else:
            st.success("✅ No critical items")


# ══════════════════════════════════════════════════════════════════════════════
# FOOTER
# ══════════════════════════════════════════════════════════════════════════════
st.markdown(f"""
<div style='text-align:center;margin-top:48px;padding-top:24px;border-top:1px solid {theme['border']};
            color:{theme['text_sec']};font-size:0.8rem;'>
    💎 VendorFace v6.0 ULTIMATE | Perfect Engine + Stunning UX | Opella Finance Operations
</div>
""", unsafe_allow_html=True)
