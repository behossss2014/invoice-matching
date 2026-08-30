import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

# ==========================================
# 1. Page Configuration & Executive Theme CSS
# ==========================================
st.set_page_config(
    page_title="Warehouse Intelligence Hub",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    /* Global Container & Direction Setup */
    [data-testid="stHeader"] { visibility: hidden; }
    
    .block-container {
        padding-top: 4.5rem !important;
        font-family: 'Inter', 'Segoe UI', -apple-system, sans-serif;
        max-width: 98%;
        direction: ltr !important;
        text-align: left !important;
    }
    
    h1, h2, h3, h4, h5, h6, p, label, div { 
        direction: ltr !important; 
        text-align: left !important; 
    }

    h1, h2, h3, h4 { 
        font-weight: 700;
        letter-spacing: -0.02em;
        margin-bottom: 0.5rem;
    }

    /* Modern Executive Navbar Theme */
    .ninja-navbar {
        position: fixed;
        top: 0;
        left: 0;
        width: 100%;
        height: 52px;
        background-color: #1e293b;
        z-index: 999999;
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 0 24px;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.15);
        border-bottom: 1px solid rgba(128, 128, 128, 0.2);
    }
    .ninja-logo {
        color: #00c9b1;
        font-size: 24px;
        font-weight: 800;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        letter-spacing: -0.5px;
    }
    .ninja-title {
        color: #f8fafc;
        font-size: 14px;
        font-weight: 600;
        letter-spacing: 0.05em;
        text-transform: uppercase;
    }

    /* Adaptive Sidebar Refinement */
    [data-testid="stSidebar"] {
        border-right: 1px solid rgba(128, 128, 128, 0.2);
    }

    /* Card & Alert Adaptive Styles */
    .alert-box {
        background-color: #f8fafc; 
        border-radius: 8px; 
        padding: 16px 20px;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05); 
        border: 1px solid #e2e8f0; 
        margin-bottom: 20px; 
        font-weight: 500; 
        font-size: 14px;
        color: #0f172a;
    }
    .card-title {
        font-size: 14px;
        font-weight: 700;
        margin-bottom: 10px;
        color: #00c9b1;
        border-bottom: 1px solid #e2e8f0;
        padding-bottom: 6px;
        text-transform: uppercase;
        letter-spacing: 0.03em;
    }

    /* Compact Executive Metric Banners */
    .exec-banner-card {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 10px 14px;
        margin-bottom: 6px;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.05);
    }
    .exec-banner-label {
        font-size: 11px;
        color: #64748b;
        font-weight: 600;
        margin-bottom: 4px;
        text-transform: uppercase;
        letter-spacing: 0.05em;
    }
    .exec-banner-val {
        font-size: 16px;
        font-weight: 700;
        color: #0f172a;
    }

    /* Automatic Dark Mode Support */
    @media (prefers-color-scheme: dark) {
        .alert-box {
            background-color: #1e293b !important;
            border-color: #334155 !important;
            color: #f8fafc !important;
        }
        .exec-banner-card {
            background-color: #1e293b !important;
            border-color: #334155 !important;
        }
        .exec-banner-label {
            color: #94a3b8 !important;
        }
        .exec-banner-val {
            color: #f8fafc !important;
        }
        .card-title {
            border-bottom-color: #334155 !important;
        }
    }

    /* Typography & Table Adjustments */
    [data-testid="stDataFrame"] {
        font-size: 13px !important;
    }
    [data-testid="stDataFrame"] div[role="gridcell"] {
        padding: 8px 12px !important;
        font-size: 13px !important;
    }
    </style>
    
    <div class="ninja-navbar">
        <div class="ninja-logo">ninja</div>
        <div class="ninja-title">Warehouse Intelligence Hub</div>
    </div>
    """, unsafe_allow_html=True)

def resolve_data_file():
    """Locate the source workbook robustly, tolerating spacing/underscore
    naming differences (e.g. 'Returns Sheet.xlsx' vs 'Returns_Sheet.xlsx')."""
    candidates = ["Returns Sheet.xlsx", "Returns_Sheet.xlsx", "Returns-Sheet.xlsx"]
    for c in candidates:
        if os.path.exists(c):
            return c
    # Fallback: any xlsx in the working directory whose normalized name matches "returnssheet"
    for f in os.listdir("."):
        if f.lower().endswith(".xlsx") and "returns" in f.lower() and "sheet" in f.lower():
            return f
    return candidates[0]

EXCEL_FILE = resolve_data_file()

# ==========================================
# Authentication & Session State Management
# ==========================================
if 'logged_in_user' not in st.session_state or not st.session_state.logged_in_user:
    st.session_state.logged_in_user = "Warehouse User"

if 'vendor_attendance' not in st.session_state:
    st.session_state.vendor_attendance = {}

if 'audit_log' not in st.session_state:
    st.session_state.audit_log = []

def record_audit(action_type, details, user="System"):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.session_state.audit_log.insert(0, {
        "Timestamp": timestamp,
        "User": user,
        "Action Type": action_type,
        "Details": details
    })

def convert_df_to_csv(df):
    return df.to_csv(index=False).encode('utf-8')

# ==========================================
# 2. Data Loading & Smart Processing
# ==========================================
@st.cache_data(ttl=60)
def load_all_sheets_live(file_path):
    if not os.path.exists(file_path): 
        return None, None, None, None, None, None
    
    xls = pd.ExcelFile(file_path)

    # Sheet names in the source workbook can carry stray leading/trailing
    # whitespace (e.g. "Closed " instead of "Closed"), which breaks exact
    # matching. Match on a stripped/lowered key instead so these variants
    # are no longer silently dropped.
    def find_sheet(candidates):
        norm_actual = {name.strip().lower(): name for name in xls.sheet_names}
        for cand in candidates:
            key = cand.strip().lower()
            if key in norm_actual:
                return norm_actual[key]
        return None

    def read_sheet(candidates):
        name = find_sheet(candidates)
        return pd.read_excel(xls, sheet_name=name) if name else pd.DataFrame()

    df_scheduled = read_sheet(['Next Day Scheduled'])
    df_returns = read_sheet(['Pending Returns'])
    df_pending_pros = read_sheet(['Pending PROs'])
    df_linked = read_sheet(['Linked_With_PO', 'Linked With PO', 'Linked_with_PO'])
    df_closed = read_sheet(['Closed', 'Closed Returns', 'Closed_Returns', 'Closed PROs', 'Closed_PROs'])
    df_supplier_damage = read_sheet(['Supplier Damage', 'Supplier_Damage', 'SupplierDamage', 'Damage'])
    # Line-item level root-cause detail for closed/linked PROs (reason per item).
    # This is the sheet that actually carries the discrepancy reason
    # (NEAR_EXPIRY, MISSED_ITEM, QUALITY_ISSUE, NOT_LISTED, NOT_ORDERED,
    # PRICE_ISSUE) — the Closed sheet itself has no reason column.
    df_issues = read_sheet(['pro_with_issues_linked_with_po_', 'pro_with_issues_linked_with_po'])

    def standardize_columns(df):
        if df.empty: 
            return
        rename_dict = {}
        for col in df.columns:
            clean_col = str(col).lower().replace(' ', '').replace('_', '').replace('-', '').replace('.', '')
            if clean_col in ['supplierno', 'buyfromvendorno', 'vendorno', 'vendorcode', 'suppliercode']:
                rename_dict[col] = 'buyFromVendorNo'
            elif clean_col in ['suppliername', 'buyfromvendorname', 'buyfromvendornam', 'vendorname', 'name']:
                rename_dict[col] = 'buyFromVendorName'
            elif clean_col in ['no', 'documentno', 'prono', 'number', 'purchasereturnorderno']:
                rename_dict[col] = 'no' 
            elif clean_col in ['reason', 'returnreason', 'discrepancyreason', 'cause']:
                rename_dict[col] = 'reason'
            elif clean_col in ['locationcode', 'location', 'whcode', 'warehouse']:
                rename_dict[col] = 'location_code'
            elif clean_col in ['tat', 'turnaroundtime', 'tatdays', 'closedays', 'responsedays', 'turnarounddays']:
                rename_dict[col] = 'tat_days'
        df.rename(columns=rename_dict, inplace=True)

    standardize_columns(df_scheduled)
    standardize_columns(df_returns)
    standardize_columns(df_linked)
    standardize_columns(df_pending_pros)
    standardize_columns(df_closed)
    standardize_columns(df_supplier_damage)
    standardize_columns(df_issues)

    if not df_supplier_damage.empty and 'Status' not in df_supplier_damage.columns:
        df_supplier_damage['Status'] = 'Supplier Damage (Awaiting CN)'
            
    def clean_vendor_code(val):
        if pd.isna(val) or str(val).strip() == "": 
            return ""
        val_str = str(val).strip().upper()
        if val_str.startswith('V'): 
            return val_str
        if val_str.replace('.', '', 1).isdigit(): 
            return f"V{int(float(val_str)):05d}"
        return val_str

    def clean_vendor_name(val):
        # Some source rows carry stray tabs/newlines/repeated spaces inside
        # the vendor name (e.g. "Rawaie Sima Trading Est\t\t\t\t\t\t"), which
        # would otherwise fragment the same real vendor into multiple
        # distinct groups in any aggregation/chart.
        if pd.isna(val):
            return val
        return " ".join(str(val).split())

    for df in [df_scheduled, df_returns, df_linked, df_pending_pros, df_closed, df_supplier_damage, df_issues]:
        if not df.empty and 'buyFromVendorNo' in df.columns:
            df['buyFromVendorNo'] = df['buyFromVendorNo'].apply(clean_vendor_code)
        if not df.empty and 'buyFromVendorName' in df.columns:
            df['buyFromVendorName'] = df['buyFromVendorName'].apply(clean_vendor_name)
            
    if not df_scheduled.empty and 'PO_Number' in df_scheduled.columns:
        df_scheduled['PO_Number'] = df_scheduled['PO_Number'].astype(str).str.strip().str.upper()
        
    for df in [df_returns, df_linked, df_pending_pros, df_closed, df_supplier_damage, df_issues]:
        if not df.empty and 'no' in df.columns:
            df['no'] = df['no'].astype(str).str.strip().str.upper()
    
    return df_scheduled, df_returns, df_linked, df_pending_pros, df_closed, df_supplier_damage, df_issues

def calculate_target_amounts(df):
    if df.empty: 
        return 0.0, 0
    amt_col = [c for c in df.columns if 'amount' in c.lower() and 'vat' in c.lower()]
    if not amt_col:
        amt_col = [c for c in df.columns if 'amount' in c.lower()] 
    total_amt = pd.to_numeric(df[amt_col[0]], errors='coerce').sum() if amt_col else 0.0
    return total_amt, len(df)

df_scheduled, df_returns, df_linked, df_pending_pros, df_closed, df_supplier_damage, df_issues = load_all_sheets_live(EXCEL_FILE)

if df_scheduled is None:
    st.error(f"System Error: Target file '{EXCEL_FILE}' not found in current directory.")
    st.stop()

# ==========================================
# Vendor Dictionary Build
# ==========================================
all_codes = set()
for df in [df_scheduled, df_returns, df_linked, df_pending_pros, df_closed, df_supplier_damage, df_issues]:
    if not df.empty and 'buyFromVendorNo' in df.columns:
        all_codes.update(df['buyFromVendorNo'].dropna().unique())

vendor_lookup = {}
for df in [df_scheduled, df_returns, df_linked, df_pending_pros, df_closed, df_supplier_damage, df_issues]:
    if not df.empty and 'buyFromVendorNo' in df.columns and 'buyFromVendorName' in df.columns:
        temp_dict = df.drop_duplicates(subset=['buyFromVendorNo']).set_index('buyFromVendorNo')['buyFromVendorName'].to_dict()
        vendor_lookup.update(temp_dict)

vendor_options_list = sorted([f"{code} - {vendor_lookup.get(code, 'Unknown Vendor')}" for code in all_codes if code])
vendor_mapping = {opt.split(" - ")[0]: opt.split(" - ")[1] for opt in vendor_options_list}

# ==========================================
# 3. Navigation Menu & Sidebar
# ==========================================
st.sidebar.markdown(f"<p style='font-size:12px; margin-top:5px;'>Logged in as:<br><b>{st.session_state.logged_in_user}</b></p>", unsafe_allow_html=True)

st.sidebar.markdown("<h3 style='margin-top: 15px; margin-bottom: 15px; font-size: 15px; font-weight:700;'>Navigation Menu</h3>", unsafe_allow_html=True)
page = st.sidebar.radio("Select Module:", [
    "Gate Operations", 
    "Today's Pre-Alerts", 
    "Data Analytics & Insights",
    "Executive Analytics",
    "Vendor SLA & Escalation Hub",
    "Audit Trail & Logs"
])
st.sidebar.markdown("---")
st.sidebar.write("**System Status:** Active (Live Mode)")

# ==========================================
# PAGE 1: GATE OPERATIONS
# ==========================================
if page == "Gate Operations":
    st.markdown("<h2>Gate Operations Dashboard</h2>", unsafe_allow_html=True)
    st.write("Search scheduled or unscheduled POs, review vendor liabilities, and stamp arrivals.")
    
    search_type = st.radio("Search Criteria:", ["PO Number", "Vendor Name/Code"], horizontal=True, key="gate_search_type")
    
    def display_vendor_status(v_no, v_name, v_time_slot, active_po_num, has_appointment=True):
        status_text = "Scheduled" if has_appointment else "Ad-hoc / Unscheduled Arrival"
        border_color = "#00C9B1" if has_appointment else "#f59e0b"
        
        st.markdown(f"""
        <div class="alert-box" style="border-left: 4px solid {border_color};">
            <div class="card-title">Vendor & Shipment Information</div>
            <p style="margin:4px 0;"><b>Name:</b> {v_name}</p>
            <p style="margin:4px 0;"><b>Code:</b> {v_no}</p>
            <p style="margin:4px 0;"><b>PO Number:</b> {active_po_num}</p>
            <p style="margin:4px 0;"><b>Type / Time Slot:</b> {v_time_slot} ({status_text})</p>
        </div>
        """, unsafe_allow_html=True)
        
        att_key = f"{v_no}_{active_po_num}"
        current_record = st.session_state.vendor_attendance.get(att_key, {"status": False, "user": "", "time": ""})
        
        if isinstance(current_record, bool):
            current_status = current_record
        elif isinstance(current_record, dict):
            current_status = current_record.get("status", False)
        else:
            current_status = False
        
        arrived_checked = st.checkbox(f"Mark Vendor as Arrived & Stamp Entry ({v_no} - {active_po_num})", value=current_status, key=f"chk_{att_key}")
        
        if arrived_checked != current_status:
            current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.session_state.vendor_attendance[att_key] = {
                "status": arrived_checked,
                "user": st.session_state.logged_in_user,
                "time": current_time,
                "po": active_po_num
            }
            action_label = "Arrived (Stamped)" if arrived_checked else "Attendance Reset"
            record_audit("Gate Arrival & Stamp Update", f"Vendor {v_no} ({v_name}) with PO {active_po_num} marked as {action_label} by {st.session_state.logged_in_user}", st.session_state.logged_in_user)
        
        if arrived_checked:
            saved_user = current_record.get('user', st.session_state.logged_in_user) if isinstance(current_record, dict) else st.session_state.logged_in_user
            saved_time = current_record.get('time', 'Just now') if isinstance(current_record, dict) else 'Just now'
            st.success(f"Verified: Arrival recorded and stamped by {saved_user} at {saved_time}.")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### Pending Returns Analysis")
            match_returns = df_returns[(df_returns['buyFromVendorNo'] == v_no) & (df_returns['Status'] == 'Pending for Collection')]
            if not match_returns.empty:
                ret_amt, _ = calculate_target_amounts(match_returns)
                
                temp_ret = match_returns.copy()
                if 'documentDate' in temp_ret.columns:
                    temp_ret['documentDate'] = pd.to_datetime(temp_ret['documentDate'], errors='coerce')
                    temp_ret['age'] = (pd.to_datetime('today') - temp_ret['documentDate']).dt.days
                    max_ret_age = temp_ret['age'].max()
                else:
                    max_ret_age = 0
                
                if max_ret_age > 30:
                    st.error(f"Critical: Returns exceeded 30 days (Max Age: {int(max_ret_age)} days)")
                elif max_ret_age > 20:
                    st.warning(f"Warning: Returns approaching limit (Max Age: {int(max_ret_age)} days)")
                else:
                    st.success(f"Status Normal (Max Age: {int(max_ret_age)} days)")
                
                st.metric(label="Total Return Amount (Inc. VAT)", value=f"{ret_amt:,.2f} SAR")
            else:
                st.success("Clear: No pending returns found.")
                
        with col2:
            st.markdown("### Outstanding Credit Notes Analysis")
            match_collected = df_returns[(df_returns['buyFromVendorNo'] == v_no) & (df_returns['Status'] == 'Collected')]
            match_linked = df_linked[df_linked['buyFromVendorNo'] == v_no] if not df_linked.empty else pd.DataFrame()
            match_damage = df_supplier_damage[df_supplier_damage['buyFromVendorNo'] == v_no] if not df_supplier_damage.empty else pd.DataFrame()
            
            amt_coll, _ = calculate_target_amounts(match_collected)
            amt_link, _ = calculate_target_amounts(match_linked)
            amt_damage, _ = calculate_target_amounts(match_damage)
            total_cn_amount = amt_coll + amt_link + amt_damage
            
            if total_cn_amount > 0:
                dfs_cn_list = [d for d in [match_collected, match_linked, match_damage] if not d.empty]
                combined_cn = pd.concat(dfs_cn_list, ignore_index=True) if dfs_cn_list else pd.DataFrame()
                cn_date_col = 'documentDate' if 'documentDate' in combined_cn.columns else ('postingDate' if 'postingDate' in combined_cn.columns else None)
                
                if cn_date_col:
                    combined_cn[cn_date_col] = pd.to_datetime(combined_cn[cn_date_col], errors='coerce')
                    combined_cn['age'] = (pd.to_datetime('today') - combined_cn[cn_date_col]).dt.days
                    max_cn_age = combined_cn['age'].max()
                else:
                    max_cn_age = 0
                
                if max_cn_age > 20:
                    st.error(f"Critical: Unclosed CN / Supplier Damage exceeded 20 days (Max Age: {int(max_cn_age)} days)")
                elif max_cn_age >= 10:
                    st.warning(f"Warning: Unclosed CN / Supplier Damage aging (Max Age: {int(max_cn_age)} days)")
                else:
                    st.success(f"Status Normal (Max Age: {int(max_cn_age)} days)")

                st.metric(label="Total Combined CN & Supplier Damage Amount (Inc. VAT)", value=f"{total_cn_amount:,.2f} SAR")
            else:
                st.success("Clear: No outstanding Credit Notes or Supplier Damage found.")

    if search_type == "PO Number":
        po_input = st.text_input("Enter PO Number (including split/suffix variants if any):", key="gate_po_field").strip().upper()
        if po_input:
            match_schedule = df_scheduled[df_scheduled['PO_Number'].str.contains(po_input, na=False)]
            if not match_schedule.empty:
                st.info(f"Found {len(match_schedule)} scheduled entry/entries for PO search:")
                for idx, row in match_schedule.iterrows():
                    t_slot = row.get('Time Slot', row.get('TimeSlot', 'N/A'))
                    display_vendor_status(row['buyFromVendorNo'], row.get('buyFromVendorName', 'Unknown Vendor'), t_slot, row['PO_Number'], True)
                    st.markdown("---")
            else:
                st.warning("PO Number not found in today's scheduled roster. Use Ad-hoc entry below to process unscheduled arrival.")
                
                with st.expander("Ad-hoc / Unscheduled PO Entry", expanded=True):
                    st.write("Register arriving shipment manually:")
                    adhoc_vendors_sel = st.multiselect("Select Vendor(s) for this PO:", options=vendor_options_list, key="adhoc_v_sel")
                    if adhoc_vendors_sel:
                        for v_sel in adhoc_vendors_sel:
                            v_code_ad = v_sel.split(" - ")[0]
                            v_name_ad = v_sel.split(" - ")[1]
                            display_vendor_status(v_code_ad, v_name_ad, "Ad-hoc Entry", po_input, False)
                            st.markdown("---")
                
    else: 
        selected_vendors = st.multiselect("Search / Select Vendor(s):", options=vendor_options_list, placeholder="Select one or multiple vendors...")
        if selected_vendors:
            for vendor_item in selected_vendors:
                v_code = vendor_item.split(" - ")[0]
                v_name = vendor_item.split(" - ")[1]
                
                match_vendor = df_scheduled[df_scheduled['buyFromVendorNo'] == v_code]
                if not match_vendor.empty:
                    for idx, row in match_vendor.iterrows():
                        t_slot = row.get('Time Slot', row.get('TimeSlot', 'N/A'))
                        display_vendor_status(v_code, v_name, t_slot, row.get('PO_Number', 'UNKNOWN-PO'), True)
                        st.markdown("---")
                else:
                    st.warning(f"Vendor ({v_name}) is not scheduled for today. Proceed with Ad-hoc PO input.")
                    adhoc_po_code = st.text_input(f"Enter Arriving PO Number for {v_name} ({v_code}):", key=f"adhoc_po_{v_code}").strip().upper()
                    if adhoc_po_code:
                        display_vendor_status(v_code, v_name, "Ad-hoc Entry", adhoc_po_code, False)
                        st.markdown("---")

    st.markdown("---")
    with st.expander("End-of-Day Gate Attendance & Stamping Report"):
        if len(st.session_state.vendor_attendance) > 0:
            att_report_data = []
            for key, info in st.session_state.vendor_attendance.items():
                if isinstance(info, bool):
                    is_arrived = info
                    u_val = "System"
                    t_val = "N/A"
                    p_val = key.split("_")[1] if "_" in key else "N/A"
                elif isinstance(info, dict):
                    is_arrived = info.get("status", False)
                    u_val = info.get("user", "System")
                    t_val = info.get("time", "N/A")
                    p_val = info.get("po", "N/A")
                else:
                    continue

                if is_arrived:
                    v_code = key.split("_")[0]
                    att_report_data.append({
                        "Vendor Code": v_code,
                        "Vendor Name": vendor_lookup.get(v_code, "Unknown"),
                        "PO Number": p_val,
                        "Processed By (User)": u_val,
                        "Timestamp": t_val,
                        "Status": "Arrived & Stamped"
                    })
            if att_report_data:
                df_att_report = pd.DataFrame(att_report_data)
                st.dataframe(df_att_report, use_container_width=True, hide_index=True)
                
                csv_attendance = convert_df_to_csv(df_att_report)
                st.download_button(
                    label="Download Attendance & Stamping Report (CSV)",
                    data=csv_attendance,
                    file_name=f"Gate_Attendance_Stamped_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv"
                )
            else:
                st.write("No vendors marked as arrived yet in this session.")
        else:
            st.write("No attendance records available.")

# ==========================================
# PAGE 2: TODAY'S PRE-ALERTS
# ==========================================
elif page == "Today's Pre-Alerts":
    st.markdown("<h2>Today's Appointments & Pre-Alert Status</h2>", unsafe_allow_html=True)
    
    if not df_scheduled.empty and 'buyFromVendorNo' in df_scheduled.columns:
        df_unique_scheduled = df_scheduled.drop_duplicates(subset=['buyFromVendorNo']).copy()
        pre_alert_list = []
        
        has_returns_count = 0
        has_cn_count = 0
        clean_count = 0
        total_arriving_returns_amt = 0.0
        total_arriving_cn_amt = 0.0
        
        current_month_cn_total = 0.0
        last_month_cn_total = 0.0
        older_cn_total = 0.0
        
        current_date = pd.to_datetime('today')
        current_year_month = current_date.to_period('M')
        last_year_month = (current_date - pd.DateOffset(months=1)).to_period('M')
        
        for idx, row in df_unique_scheduled.iterrows():
            v_no = row['buyFromVendorNo']
            
            v_returns = df_returns[(df_returns['buyFromVendorNo'] == v_no) & (df_returns['Status'] == 'Pending for Collection')]
            ret_amt, ret_count = calculate_target_amounts(v_returns)
            
            v_collected = df_returns[(df_returns['buyFromVendorNo'] == v_no) & (df_returns['Status'] == 'Collected')].copy()
            v_linked = df_linked[df_linked['buyFromVendorNo'] == v_no].copy() if not df_linked.empty else pd.DataFrame()
            v_damage = df_supplier_damage[df_supplier_damage['buyFromVendorNo'] == v_no].copy() if not df_supplier_damage.empty else pd.DataFrame()
            
            amt_coll, count_coll = calculate_target_amounts(v_collected)
            amt_link, count_link = calculate_target_amounts(v_linked)
            amt_damage, count_damage = calculate_target_amounts(v_damage)
            
            total_cn_amt = amt_coll + amt_link + amt_damage
            total_cn_count = count_coll + count_link + count_damage
            
            for sub_df in [v_collected, v_linked, v_damage]:
                if not sub_df.empty:
                    if 'standard_amount' not in sub_df.columns:
                        a_cols = [c for c in sub_df.columns if 'amount' in c.lower() and 'vat' in c.lower()]
                        if not a_cols:
                            a_cols = [c for c in sub_df.columns if 'amount' in c.lower()]
                        if a_cols:
                            sub_df['standard_amount'] = pd.to_numeric(sub_df[a_cols[0]], errors='coerce').fillna(0.0)
                        else:
                            sub_df['standard_amount'] = 0.0
                    
                    if 'standard_date' not in sub_df.columns:
                        d_cols = [c for c in sub_df.columns if any(k in c.lower() for k in ['postingdate', 'documentdate', 'date'])]
                        if d_cols:
                            sub_df['standard_date'] = pd.to_datetime(sub_df[d_cols[0]], errors='coerce')
                        else:
                            sub_df['standard_date'] = pd.NaT

            dfs_cn_all = [d for d in [v_collected, v_linked, v_damage] if not d.empty]
            combined_cn = pd.concat(dfs_cn_all, ignore_index=True) if dfs_cn_all else pd.DataFrame()
            
            max_cn_status_priority = 0 
            
            if not combined_cn.empty and 'standard_amount' in combined_cn.columns:
                for _, cn_row in combined_cn.iterrows():
                    c_date = cn_row.get('standard_date', pd.NaT)
                    c_amt = cn_row.get('standard_amount', 0.0)
                    
                    if pd.isna(c_date):
                        older_cn_total += c_amt
                        max_cn_status_priority = max(max_cn_status_priority, 3)
                    else:
                        p = pd.Period(c_date, freq='M')
                        if p == current_year_month:
                            current_month_cn_total += c_amt
                            max_cn_status_priority = max(max_cn_status_priority, 1)
                        elif p == last_year_month:
                            last_month_cn_total += c_amt
                            max_cn_status_priority = max(max_cn_status_priority, 2)
                        elif p < last_year_month:
                            older_cn_total += c_amt
                            max_cn_status_priority = max(max_cn_status_priority, 3)
                        else:
                            current_month_cn_total += c_amt
                            max_cn_status_priority = max(max_cn_status_priority, 1)
            
            if max_cn_status_priority == 3:
                cn_status_text = "Critical: Claims from Previous Months"
            elif max_cn_status_priority == 2:
                cn_status_text = "Warning: Last Month Claims"
            elif max_cn_status_priority == 1:
                cn_status_text = "Normal: Current Month Claims"
            else:
                cn_status_text = "Clear"

            is_blocked = (ret_amt > 0 or total_cn_amt > 0)
            if ret_amt > 0:
                has_returns_count += 1
                total_arriving_returns_amt += ret_amt
            if total_cn_amt > 0:
                has_cn_count += 1
                total_arriving_cn_amt += total_cn_amt
            if not is_blocked:
                clean_count += 1
            
            pre_alert_list.append({
                "Vendor Code": v_no,
                "Vendor": row.get('buyFromVendorName', 'Unknown Vendor'),
                "Returns Value (SAR)": ret_amt,
                "PRO Count": ret_count,
                "CN Balance (SAR)": total_cn_amt,
                "CN Count": total_cn_count,
                "CN Aging Status": cn_status_text,
                "Decision": "Block" if is_blocked else "Clear Pass"
            })

        tot_vendors = len(df_unique_scheduled)
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Arriving Vendors", f"{tot_vendors}")
        m2.metric("Vendors with Returns", f"{has_returns_count}", f"{total_arriving_returns_amt:,.2f} SAR")
        m3.metric("Vendors with CNs / Damage", f"{has_cn_count}", f"{total_arriving_cn_amt:,.2f} SAR")
        m4.metric("Clear Pass Vendors", f"{clean_count}")
        
        st.markdown(f"""
        <div class="exec-banner-card" style="margin-top: 10px;">
            <div class="exec-banner-label">Credit Notes & Supplier Damage Aging Summary</div>
            <div class="exec-banner-val" style="font-size: 14px;">
                Current Month: <b>{current_month_cn_total:,.2f} SAR</b> | 
                Last Month: <b>{last_month_cn_total:,.2f} SAR</b> | 
                Older: <b style="color: #ef4444;">{older_cn_total:,.2f} SAR</b>
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("---")

        df_final = pd.DataFrame(pre_alert_list)
        st.dataframe(
            df_final, 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "Returns Value (SAR)": st.column_config.NumberColumn(format="%,.2f"),
                "CN Balance (SAR)": st.column_config.NumberColumn(format="%,.2f")
            }
        )
        
        csv_pre_alerts = convert_df_to_csv(df_final)
        st.download_button(
            label="Download Pre-Alerts Report (CSV)",
            data=csv_pre_alerts,
            file_name=f"Pre_Alerts_Report_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
    else:
        st.write("No scheduled appointments available in the current dataset.")

# ==========================================
# PAGE 3: DATA ANALYTICS & INSIGHTS
# ==========================================
elif page == "Data Analytics & Insights":
    st.markdown("<h2>Data Analytics & Insights Hub</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    loc_options = set()
    bin_options = set()
    for df_temp in [df_returns, df_linked, df_pending_pros, df_supplier_damage]:
        if df_temp is not None and not df_temp.empty:
            if 'locationCode' in df_temp.columns:
                loc_options.update(df_temp['locationCode'].dropna().astype(str).unique())
            elif 'location_code' in df_temp.columns:
                loc_options.update(df_temp['location_code'].dropna().astype(str).unique())
            if 'binCode' in df_temp.columns:
                bin_options.update(df_temp['binCode'].dropna().astype(str).unique())
    
    loc_options_list = sorted([str(loc) for loc in loc_options if str(loc).strip() and str(loc).lower() != 'nan'])
    bin_options_list = sorted([str(b) for b in bin_options if str(b).strip() and str(b).lower() != 'nan'])

    with st.sidebar.expander("Filter Analytics Data", expanded=True):
        selected_vendors = st.multiselect(
            "Search / Select Vendor(s):", 
            options=vendor_options_list,
            placeholder="Select one or multiple vendors...",
            key="analytics_vendor_multiselect"
        )

        selected_locations = st.multiselect(
            "Filter by Location Code:",
            options=loc_options_list,
            placeholder="Select Location Code(s)...",
            key="analytics_location_multiselect"
        )

        selected_bins = st.multiselect(
            "Filter by Bin Code:",
            options=bin_options_list,
            placeholder="Select Bin Code(s)...",
            key="analytics_bin_multiselect"
        )

        date_range = st.date_input("Select Period (Document/Posting Date):", [], key="analytics_date_range")
        
    filtered_df = df_returns.copy()
    filtered_linked = df_linked.copy() if not df_linked.empty else pd.DataFrame()
    filtered_pros = df_pending_pros.copy() if not df_pending_pros.empty else pd.DataFrame()
    filtered_damage = df_supplier_damage.copy() if not df_supplier_damage.empty else pd.DataFrame()
    
    display_title = "All Open Records"
    
    if selected_vendors:
        selected_codes = [v.split(" - ")[0] for v in selected_vendors]
        vendor_display_list = [f"{code} - {vendor_lookup.get(code, 'Unknown Vendor')}" for code in selected_codes]
        display_title = f"Vendors: {', '.join(vendor_display_list)}"
        
        filtered_df = filtered_df[filtered_df['buyFromVendorNo'].isin(selected_codes)]
        if not filtered_linked.empty:
            filtered_linked = filtered_linked[filtered_linked['buyFromVendorNo'].isin(selected_codes)]
        if not filtered_pros.empty and 'buyFromVendorNo' in filtered_pros.columns:
            filtered_pros = filtered_pros[filtered_pros['buyFromVendorNo'].isin(selected_codes)]
        if not filtered_damage.empty and 'buyFromVendorNo' in filtered_damage.columns:
            filtered_damage = filtered_damage[filtered_damage['buyFromVendorNo'].isin(selected_codes)]

    if selected_locations:
        if not filtered_df.empty and 'locationCode' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['locationCode'].astype(str).isin(selected_locations)]
        if not filtered_linked.empty and 'locationCode' in filtered_linked.columns:
            filtered_linked = filtered_linked[filtered_linked['locationCode'].astype(str).isin(selected_locations)]
        if not filtered_pros.empty and 'locationCode' in filtered_pros.columns:
            filtered_pros = filtered_pros[filtered_pros['locationCode'].astype(str).isin(selected_locations)]
        if not filtered_damage.empty:
            loc_col = 'locationCode' if 'locationCode' in filtered_damage.columns else ('location_code' if 'location_code' in filtered_damage.columns else None)
            if loc_col:
                filtered_damage = filtered_damage[filtered_damage[loc_col].astype(str).isin(selected_locations)]

    if selected_bins:
        if not filtered_df.empty and 'binCode' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['binCode'].astype(str).isin(selected_bins)]
        if not filtered_linked.empty and 'binCode' in filtered_linked.columns:
            filtered_linked = filtered_linked[filtered_linked['binCode'].astype(str).isin(selected_bins)]
        if not filtered_pros.empty and 'binCode' in filtered_pros.columns:
            filtered_pros = filtered_pros[filtered_pros['binCode'].astype(str).isin(selected_bins)]
        if not filtered_damage.empty and 'binCode' in filtered_damage.columns:
            filtered_damage = filtered_damage[filtered_damage['binCode'].astype(str).isin(selected_bins)]

    if len(date_range) == 2:
        start_date, end_date = date_range
        
        if not filtered_df.empty:
            if 'postingDate' in filtered_df.columns: 
                filtered_df['postingDate'] = pd.to_datetime(filtered_df['postingDate'], errors='coerce')
            if 'documentDate' in filtered_df.columns: 
                filtered_df['documentDate'] = pd.to_datetime(filtered_df['documentDate'], errors='coerce')
            
            pending_mask = (filtered_df['Status'] == 'Pending for Collection') & (filtered_df['documentDate'].dt.date >= start_date) & (filtered_df['documentDate'].dt.date <= end_date)
            collected_mask = (filtered_df['Status'] == 'Collected') & (filtered_df['postingDate'].dt.date >= start_date) & (filtered_df['postingDate'].dt.date <= end_date)
            filtered_df = filtered_df[pending_mask | collected_mask]
        
        if not filtered_linked.empty and 'documentDate' in filtered_linked.columns:
            filtered_linked['documentDate'] = pd.to_datetime(filtered_linked['documentDate'], errors='coerce')
            mask_link = (filtered_linked['documentDate'].dt.date >= start_date) & (filtered_linked['documentDate'].dt.date <= end_date)
            filtered_linked = filtered_linked[mask_link.fillna(False)]
            
        if not filtered_pros.empty and 'documentDate' in filtered_pros.columns:
            filtered_pros['documentDate'] = pd.to_datetime(filtered_pros['documentDate'], errors='coerce')
            mask_pros = (filtered_pros['documentDate'].dt.date >= start_date) & (filtered_pros['documentDate'].dt.date <= end_date)
            filtered_pros = filtered_pros[mask_pros.fillna(False)]

        if not filtered_damage.empty:
            d_col = 'documentDate' if 'documentDate' in filtered_damage.columns else ('postingDate' if 'postingDate' in filtered_damage.columns else None)
            if d_col:
                filtered_damage[d_col] = pd.to_datetime(filtered_damage[d_col], errors='coerce')
                mask_dmg = (filtered_damage[d_col].dt.date >= start_date) & (filtered_damage[d_col].dt.date <= end_date)
                filtered_damage = filtered_damage[mask_dmg.fillna(False)]

    returns_df = filtered_df[filtered_df['Status'] == 'Pending for Collection'].copy()
    collected_df = filtered_df[filtered_df['Status'] == 'Collected'].copy()
    
    total_ret_amt, count_ret = calculate_target_amounts(returns_df)
    total_coll_amt, count_coll = calculate_target_amounts(collected_df)
    total_link_amt, count_link = calculate_target_amounts(filtered_linked)
    total_dmg_amt, count_dmg = calculate_target_amounts(filtered_damage)
    
    total_cn_exposure = total_coll_amt + total_link_amt + total_dmg_amt

    st.subheader(f"Results scope: {display_title}")
    c1, c2 = st.columns(2)
    c1.metric("Pending Returns Value", f"{total_ret_amt:,.2f} SAR")
    c1.caption(f"Document Count: {count_ret}")
    c2.metric("Total CN & Damage Exposure", f"{total_cn_exposure:,.2f} SAR")
    c2.caption(f"Collected: {total_coll_amt:,.2f} ({count_coll}) | Linked PO: {total_link_amt:,.2f} ({count_link}) | Supplier Damage: {total_dmg_amt:,.2f} ({count_dmg})")

    st.markdown("### Detailed Record Analysis")

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Pending Returns", "Collected CNs", "Linked PO CNs", "Supplier Damage (Awaiting CN)", "Orphaned PROs"])
    
    with tab1:
        if not returns_df.empty:
            df_t1 = returns_df.copy()
            df_t1['VendorName'] = df_t1['buyFromVendorNo'].map(vendor_mapping).fillna("Unknown Vendor")
            
            if 'documentDate' in df_t1.columns:
                df_t1['documentDate'] = pd.to_datetime(df_t1['documentDate'], errors='coerce')
                df_t1['aging_days'] = (pd.to_datetime('today') - df_t1['documentDate']).dt.days
            else:
                df_t1['documentDate'] = pd.NaT
                df_t1['aging_days'] = 0

            rec_cols = [c for c in df_t1.columns if 'lastreceiving' in str(c).lower().replace(' ', '').replace('_', '')]
            if rec_cols:
                df_t1['days_since_last_receiving'] = df_t1[rec_cols[0]]
            elif 'days_since_last_receiving' not in df_t1.columns:
                df_t1['days_since_last_receiving'] = None

            if 'amount' not in df_t1.columns: 
                df_t1['amount'] = 0.0
            if 'amountIncludingVAT' not in df_t1.columns:
                vat_cols = [c for c in df_t1.columns if 'vat' in str(c).lower()]
                df_t1['amountIncludingVAT'] = df_t1[vat_cols[0]] if vat_cols else 0.0

            if 'binCode' not in df_t1.columns: 
                df_t1['binCode'] = "N/A"
            if 'locationCode' not in df_t1.columns: 
                df_t1['locationCode'] = "N/A"
            if 'Status' not in df_t1.columns: 
                df_t1['Status'] = 'Pending for Collection'

            cols_t1 = ['buyFromVendorNo', 'VendorName', 'documentDate', 'binCode', 'locationCode', 'amount', 'amountIncludingVAT', 'aging_days', 'no', 'days_since_last_receiving', 'Status']
            view_df1 = df_t1.reindex(columns=cols_t1)
            
            st.dataframe(
                view_df1, 
                use_container_width=True, hide_index=True,
                column_config={
                    "amount": st.column_config.NumberColumn(format="%,.2f"), 
                    "amountIncludingVAT": st.column_config.NumberColumn(format="%,.2f"), 
                    "documentDate": st.column_config.DateColumn(format="YYYY-MM-DD")
                }
            )
            st.download_button("Download Pending Returns (CSV)", data=convert_df_to_csv(view_df1), file_name="Pending_Returns.csv", mime="text/csv", key="dl_t1")
        else: 
            st.write("No pending returns identified for the selected criteria.")
            
    with tab2:
        if not collected_df.empty:
            df_t2 = collected_df.copy()
            df_t2['VendorName'] = df_t2['buyFromVendorNo'].map(vendor_mapping).fillna("Unknown Vendor")
            
            if 'postingDate' in df_t2.columns:
                df_t2['postingDate'] = pd.to_datetime(df_t2['postingDate'], errors='coerce')
                df_t2['aging_days'] = (pd.to_datetime('today') - df_t2['postingDate']).dt.days
            else:
                df_t2['postingDate'] = pd.NaT
                df_t2['aging_days'] = 0

            rec_cols = [c for c in df_t2.columns if 'lastreceiving' in str(c).lower().replace(' ', '').replace('_', '')]
            if rec_cols: 
                df_t2['days_since_last_receiving'] = df_t2[rec_cols[0]]
            elif 'days_since_last_receiving' not in df_t2.columns: 
                df_t2['days_since_last_receiving'] = None

            if 'amount' not in df_t2.columns: 
                df_t2['amount'] = 0.0
            if 'amountIncludingVAT' not in df_t2.columns:
                vat_cols = [c for c in df_t2.columns if 'vat' in str(c).lower()]
                df_t2['amountIncludingVAT'] = df_t2[vat_cols[0]] if vat_cols else 0.0

            if 'Status' not in df_t2.columns: 
                df_t2['Status'] = 'Collected'

            cols_t2 = ['buyFromVendorNo', 'VendorName', 'amount', 'amountIncludingVAT', 'postingDate', 'aging_days', 'no', 'days_since_last_receiving', 'Status']
            view_df2 = df_t2.reindex(columns=cols_t2)

            st.dataframe(
                view_df2, 
                use_container_width=True, hide_index=True,
                column_config={
                    "amount": st.column_config.NumberColumn(format="%,.2f"), 
                    "amountIncludingVAT": st.column_config.NumberColumn(format="%,.2f"), 
                    "postingDate": st.column_config.DateColumn(format="YYYY-MM-DD")
                }
            )
            st.download_button("Download Collected CNs (CSV)", data=convert_df_to_csv(view_df2), file_name="Collected_CNs.csv", mime="text/csv", key="dl_t2")
        else: 
            st.write("No collected records identified for the selected criteria.")
            
    with tab3:
        if not filtered_linked.empty:
            df_t3 = filtered_linked.copy()
            df_t3['VendorName'] = df_t3['buyFromVendorNo'].map(vendor_mapping).fillna("Unknown Vendor")
            
            if 'documentDate' in df_t3.columns:
                df_t3['documentDate'] = pd.to_datetime(df_t3['documentDate'], errors='coerce')
                df_t3['Aging'] = (pd.to_datetime('today') - df_t3['documentDate']).dt.days
            else:
                df_t3['documentDate'] = pd.NaT
                df_t3['Aging'] = 0

            if 'amount' not in df_t3.columns: 
                df_t3['amount'] = 0.0
            vat_cols = [c for c in df_t3.columns if 'vat' in str(c).lower()]
            if 'AmountWithVAT' in df_t3.columns: 
                pass
            elif 'amountIncludingVAT' in df_t3.columns: 
                df_t3['AmountWithVAT'] = df_t3['amountIncludingVAT']
            elif vat_cols: 
                df_t3['AmountWithVAT'] = df_t3[vat_cols[0]]
            else: 
                df_t3['AmountWithVAT'] = 0.0

            if 'locationCode' not in df_t3.columns: 
                df_t3['locationCode'] = "N/A"
            if 'Status' not in df_t3.columns: 
                df_t3['Status'] = 'Linked with PO'

            cols_t3 = ['buyFromVendorNo', 'VendorName', 'documentDate', 'no', 'locationCode', 'amount', 'AmountWithVAT', 'Status', 'Aging']
            view_df3 = df_t3.reindex(columns=cols_t3)

            st.dataframe(
                view_df3, 
                use_container_width=True, hide_index=True,
                column_config={
                    "amount": st.column_config.NumberColumn(format="%,.2f"), 
                    "AmountWithVAT": st.column_config.NumberColumn(format="%,.2f"), 
                    "documentDate": st.column_config.DateColumn(format="YYYY-MM-DD")
                }
            )
            st.download_button("Download Linked PO CNs (CSV)", data=convert_df_to_csv(view_df3), file_name="Linked_PO_CNs.csv", mime="text/csv", key="dl_t3")
        else: 
            st.write("No Linked PO discrepancies found.")

    with tab4:
        if not filtered_damage.empty:
            df_t4_dmg = filtered_damage.copy()
            df_t4_dmg['VendorName'] = df_t4_dmg['buyFromVendorNo'].map(vendor_mapping).fillna("Unknown Vendor")
            
            d_col = 'documentDate' if 'documentDate' in df_t4_dmg.columns else ('postingDate' if 'postingDate' in df_t4_dmg.columns else None)
            if d_col:
                df_t4_dmg['documentDate'] = pd.to_datetime(df_t4_dmg[d_col], errors='coerce')
                df_t4_dmg['Aging'] = (pd.to_datetime('today') - df_t4_dmg['documentDate']).dt.days
            else:
                df_t4_dmg['documentDate'] = pd.NaT
                df_t4_dmg['Aging'] = 0

            if 'amount' not in df_t4_dmg.columns: 
                df_t4_dmg['amount'] = 0.0
            vat_cols = [c for c in df_t4_dmg.columns if 'vat' in str(c).lower()]
            if 'AmountWithVAT' in df_t4_dmg.columns: 
                pass
            elif 'amountIncludingVAT' in df_t4_dmg.columns: 
                df_t4_dmg['AmountWithVAT'] = df_t4_dmg['amountIncludingVAT']
            elif vat_cols: 
                df_t4_dmg['AmountWithVAT'] = df_t4_dmg[vat_cols[0]]
            else: 
                df_t4_dmg['AmountWithVAT'] = 0.0

            loc_col = 'locationCode' if 'locationCode' in df_t4_dmg.columns else ('location_code' if 'location_code' in df_t4_dmg.columns else None)
            if not loc_col or loc_col not in df_t4_dmg.columns:
                df_t4_dmg['locationCode'] = "N/A"
            else:
                df_t4_dmg['locationCode'] = df_t4_dmg[loc_col]

            if 'Status' not in df_t4_dmg.columns: 
                df_t4_dmg['Status'] = 'Supplier Damage (Awaiting CN)'

            cols_t4_dmg = ['buyFromVendorNo', 'VendorName', 'documentDate', 'no', 'locationCode', 'amount', 'AmountWithVAT', 'Status', 'Aging']
            cols_t4_dmg = [c for c in cols_t4_dmg if c in df_t4_dmg.columns]
            view_df4_dmg = df_t4_dmg.reindex(columns=cols_t4_dmg)

            st.dataframe(
                view_df4_dmg, 
                use_container_width=True, hide_index=True,
                column_config={
                    "amount": st.column_config.NumberColumn(format="%,.2f"), 
                    "AmountWithVAT": st.column_config.NumberColumn(format="%,.2f"), 
                    "documentDate": st.column_config.DateColumn(format="YYYY-MM-DD")
                }
            )
            st.download_button("Download Supplier Damage (CSV)", data=convert_df_to_csv(view_df4_dmg), file_name="Supplier_Damage_CN_Awaiting.csv", mime="text/csv", key="dl_t4_dmg")
        else:
            st.write("No Supplier Damage records identified for the selected criteria.")
        
    with tab5:
        global_linked_nos = df_linked['no'].unique() if ('no' in df_linked.columns and not df_linked.empty) else []
        global_returns_nos = df_returns['no'].unique() if ('no' in df_returns.columns and not df_returns.empty) else []
        global_damage_nos = df_supplier_damage['no'].unique() if ('no' in df_supplier_damage.columns and not df_supplier_damage.empty) else []

        if not filtered_pros.empty and 'no' in filtered_pros.columns:
            df_orphans = filtered_pros[
                (~filtered_pros['no'].isin(global_linked_nos)) & 
                (~filtered_pros['no'].isin(global_returns_nos)) &
                (~filtered_pros['no'].isin(global_damage_nos))
            ]
        else:
            df_orphans = pd.DataFrame()

        if not df_orphans.empty:
            orphan_amt, orphan_count = calculate_target_amounts(df_orphans)
            st.write(f"Alert: {orphan_count} Orphaned PROs Detected | Total Value: {orphan_amt:,.2f} SAR")
            
            df_t5 = df_orphans.copy()
            df_t5['VendorName'] = df_t5['buyFromVendorNo'].map(vendor_mapping).fillna("Unknown Vendor")
            cols_t5 = ['buyFromVendorNo', 'VendorName'] + [c for c in df_t5.columns if c not in ['buyFromVendorNo', 'VendorName']]
            view_df5 = df_t5[cols_t5]
            
            st.dataframe(
                view_df5, 
                use_container_width=True, hide_index=True,
                column_config={
                    "amount": st.column_config.NumberColumn(format="%,.2f"), 
                    "amountIncludingVAT": st.column_config.NumberColumn(format="%,.2f")
                }
            )
            st.download_button("Download Orphaned PROs (CSV)", data=convert_df_to_csv(view_df5), file_name="Orphaned_PROs.csv", mime="text/csv", key="dl_t5")
        else:
            st.write("System Status Normal: No Orphaned PROs detected.")

# ==========================================
# PAGE 4: EXECUTIVE ANALYTICS
# ==========================================
elif page == "Executive Analytics":
    st.markdown("<h2>Executive Analytics Dashboard</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    exec_df = df_returns.copy()
    if not exec_df.empty:
        vat_cols = [c for c in exec_df.columns if 'amount' in str(c).lower() and 'vat' in str(c).lower()]
        if not vat_cols:
            vat_cols = [c for c in exec_df.columns if 'amount' in str(c).lower()]
        target_val_col = vat_cols[0] if vat_cols else 'amount'
        exec_df['amountIncludingVAT'] = pd.to_numeric(exec_df[target_val_col], errors='coerce').fillna(0.0)

    st.sidebar.header("Configuration Parameters")
    
    all_dates = pd.Series(dtype='datetime64[ns]')
    if 'documentDate' in exec_df.columns: 
        all_dates = pd.concat([all_dates, pd.to_datetime(exec_df['documentDate'], errors='coerce')])
    if 'postingDate' in exec_df.columns: 
        all_dates = pd.concat([all_dates, pd.to_datetime(exec_df['postingDate'], errors='coerce')])
        
    valid_dates = all_dates.dropna()
    default_start = valid_dates.min().date() if not valid_dates.empty else pd.to_datetime('2025-01-01').date()
    default_end = valid_dates.max().date() if not valid_dates.empty else pd.to_datetime('today').date()

    date_range_exec = st.sidebar.date_input(
        "Observation Period:",
        value=(default_start, default_end),
        min_value=pd.to_datetime('2020-01-01').date(),
        max_value=pd.to_datetime('2030-12-31').date()
    )
    
    st.sidebar.markdown("---")
    top_cn_n = st.sidebar.number_input("Top Variables Filter (CN & Damage):", min_value=1, max_value=100, value=10)
    top_ret_n = st.sidebar.number_input("Top Variables Filter (Returns):", min_value=1, max_value=100, value=10)
    top_aging_n = st.sidebar.number_input("Aging Summary Display Limit:", min_value=1, max_value=200, value=50)

    def get_vendor_label(code):
        return f"{code} - {vendor_mapping.get(code, 'Unknown')}"

    df_exec_filtered = exec_df.copy()
    df_linked_exec = df_linked.copy() if not df_linked.empty else pd.DataFrame()
    df_damage_exec = df_supplier_damage.copy() if not df_supplier_damage.empty else pd.DataFrame()
    
    if not df_damage_exec.empty:
        vat_cols_dmg = [c for c in df_damage_exec.columns if 'amount' in str(c).lower() and 'vat' in str(c).lower()]
        if not vat_cols_dmg:
            vat_cols_dmg = [c for c in df_damage_exec.columns if 'amount' in str(c).lower()]
        t_col_dmg = vat_cols_dmg[0] if vat_cols_dmg else 'amount'
        df_damage_exec['amountIncludingVAT'] = pd.to_numeric(df_damage_exec[t_col_dmg], errors='coerce').fillna(0.0)
        df_damage_exec['Status'] = 'Supplier Damage'

    if isinstance(date_range_exec, tuple) and len(date_range_exec) == 2:
        start_date, end_date = date_range_exec
        if 'documentDate' in df_exec_filtered.columns: 
            df_exec_filtered['documentDate'] = pd.to_datetime(df_exec_filtered['documentDate'], errors='coerce')
        if 'postingDate' in df_exec_filtered.columns: 
            df_exec_filtered['postingDate'] = pd.to_datetime(df_exec_filtered['postingDate'], errors='coerce')
            
        mask_pending = (df_exec_filtered['Status'] == 'Pending for Collection') & (
            (df_exec_filtered['documentDate'].dt.date >= start_date) & (df_exec_filtered['documentDate'].dt.date <= end_date) | df_exec_filtered['documentDate'].isna()
        )
        mask_collected = (df_exec_filtered['Status'] == 'Collected') & (
            (df_exec_filtered.get('postingDate', df_exec_filtered['documentDate']).dt.date >= start_date) & 
            (df_exec_filtered.get('postingDate', df_exec_filtered['documentDate']).dt.date <= end_date)
        )
        
        df_exec_filtered = df_exec_filtered[mask_pending | mask_collected]

        if not df_linked_exec.empty and 'documentDate' in df_linked_exec.columns:
            df_linked_exec['documentDate'] = pd.to_datetime(df_linked_exec['documentDate'], errors='coerce')
            mask_link = (df_linked_exec['documentDate'].dt.date >= start_date) & (df_linked_exec['documentDate'].dt.date <= end_date)
            df_linked_exec = df_linked_exec[mask_link | df_linked_exec['documentDate'].isna()]

        if not df_damage_exec.empty:
            d_col_dmg = 'documentDate' if 'documentDate' in df_damage_exec.columns else ('postingDate' if 'postingDate' in df_damage_exec.columns else None)
            if d_col_dmg:
                df_damage_exec[d_col_dmg] = pd.to_datetime(df_damage_exec[d_col_dmg], errors='coerce')
                mask_dmg_exec = (df_damage_exec[d_col_dmg].dt.date >= start_date) & (df_damage_exec[d_col_dmg].dt.date <= end_date)
                df_damage_exec = df_damage_exec[mask_dmg_exec | df_damage_exec[d_col_dmg].isna()]

        st.write(f"Analytics Scope: Data spanning **{start_date}** to **{end_date}**.")

    # Chart 1: Outstanding CN & Supplier Damage Liability
    st.subheader(f"Top {top_cn_n} Vendors: Outstanding CN & Supplier Damage Liability")
    cn_collected_df = df_exec_filtered[df_exec_filtered['Status'] == 'Collected'].copy()
    
    cn_combined_all = pd.concat([d for d in [cn_collected_df, df_damage_exec] if not d.empty], ignore_index=True) if any(not d.empty for d in [cn_collected_df, df_damage_exec]) else pd.DataFrame()

    if not cn_combined_all.empty and 'buyFromVendorNo' in cn_combined_all.columns:
        cn_data = cn_combined_all.groupby('buyFromVendorNo')['amountIncludingVAT'].sum().nlargest(top_cn_n).reset_index()
    else:
        cn_data = pd.DataFrame()
    
    if not cn_data.empty:
        cn_data['Vendor_Label'] = cn_data['buyFromVendorNo'].apply(get_vendor_label)
        fig1 = px.bar(
            cn_data, 
            x='amountIncludingVAT', 
            y='Vendor_Label', 
            orientation='h', 
            color='amountIncludingVAT', 
            color_continuous_scale=[[0, "#334155"], [1, "#00c9b1"]]
        )
        fig1.update_layout(
            yaxis_title="", 
            xaxis_title="Total Exposure (SAR)", 
            plot_bgcolor='rgba(0,0,0,0)', 
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Inter, sans-serif"),
            margin=dict(l=0, r=20, t=10, b=10)
        )
        fig1.update_traces(hovertemplate='Vendor: %{y}<br>Amount: %{x:,.2f} SAR<extra></extra>')
        st.plotly_chart(fig1, use_container_width=True)

        cn_shown_val = cn_data['amountIncludingVAT'].sum()
        cn_total_system_val = cn_combined_all['amountIncludingVAT'].sum() if not cn_combined_all.empty else 0.0
        cn_share_pct = (cn_shown_val / cn_total_system_val * 100) if cn_total_system_val > 0 else 0.0

        st.markdown("<p style='font-size: 13px; font-weight: 700; margin-bottom: 6px;'> Performance Context </p>", unsafe_allow_html=True)
        b1, b2, b3, b4 = st.columns([1.2, 1.2, 1, 0.8])
        
        with b1:
            st.markdown(f"""
            <div class="exec-banner-card">
                <div class="exec-banner-label">Selected Category Total (Top Vendors)</div>
                <div class="exec-banner-val">{cn_shown_val:,.2f} SAR</div>
            </div>
            """, unsafe_allow_html=True)
            
        with b2:
            st.markdown(f"""
            <div class="exec-banner-card">
                <div class="exec-banner-label">System Total</div>
                <div class="exec-banner-val">{cn_total_system_val:,.2f} SAR</div>
            </div>
            """, unsafe_allow_html=True)
            
        with b3:
            st.markdown(f"""
            <div class="exec-banner-card">
                <div class="exec-banner-label">Actual Representation Ratio</div>
                <div class="exec-banner-val">{cn_share_pct:.2f}%</div>
            </div>
            """, unsafe_allow_html=True)
            
        with b4:
            st.write("")
            if st.button("Focus Mode", key="focus_cn_mode", help="Direct focus on this group's data"):
                st.session_state['focus_cn_list'] = cn_data['buyFromVendorNo'].tolist()
                st.success(f"Focus Mode applied to top {len(cn_data)} vendors.")

        st.progress(min(max(cn_share_pct / 100.0, 0.0), 1.0))
        st.markdown("---")
    else: 
        st.write("Insufficient data to generate chart.")

    # Chart 2: Stagnant Pending Returns
    st.subheader(f"Top {top_ret_n} Vendors: Stagnant Pending Returns")
    ret_pending_df = df_exec_filtered[df_exec_filtered['Status'] == 'Pending for Collection']
    ret_data = ret_pending_df.groupby('buyFromVendorNo')['amountIncludingVAT'].sum().nlargest(top_ret_n).reset_index()
    
    if not ret_data.empty:
        ret_data['Vendor_Label'] = ret_data['buyFromVendorNo'].apply(get_vendor_label)
        fig2 = px.bar(
            ret_data, 
            x='amountIncludingVAT', 
            y='Vendor_Label', 
            orientation='h', 
            color='amountIncludingVAT', 
            color_continuous_scale=[[0, "#1e293b"], [1, "#0ea5e9"]]
        )
        fig2.update_layout(
            yaxis_title="", 
            xaxis_title="Total Value (SAR)", 
            plot_bgcolor='rgba(0,0,0,0)', 
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Inter, sans-serif"),
            margin=dict(l=0, r=20, t=10, b=10)
        )
        fig2.update_traces(hovertemplate='Vendor: %{y}<br>Amount: %{x:,.2f} SAR<extra></extra>')
        st.plotly_chart(fig2, use_container_width=True)

        ret_shown_val = ret_data['amountIncludingVAT'].sum()
        ret_total_system_val = exec_df[exec_df['Status'] == 'Pending for Collection']['amountIncludingVAT'].sum() if not exec_df.empty else 0.0
        ret_share_pct = (ret_shown_val / ret_total_system_val * 100) if ret_total_system_val > 0 else 0.0

        st.markdown("<p style='font-size: 13px; font-weight: 700; margin-bottom: 6px;'> Performance Context </p>", unsafe_allow_html=True)
        rb1, rb2, rb3, rb4 = st.columns([1.2, 1.2, 1, 0.8])
        
        with rb1:
            st.markdown(f"""
            <div class="exec-banner-card">
                <div class="exec-banner-label">Selected Category Total (Top Vendors)</div>
                <div class="exec-banner-val">{ret_shown_val:,.2f} SAR</div>
            </div>
            """, unsafe_allow_html=True)
            
        with rb2:
            st.markdown(f"""
            <div class="exec-banner-card">
                <div class="exec-banner-label">System Total</div>
                <div class="exec-banner-val">{ret_total_system_val:,.2f} SAR</div>
            </div>
            """, unsafe_allow_html=True)
            
        with rb3:
            st.markdown(f"""
            <div class="exec-banner-card">
                <div class="exec-banner-label">Actual Representation Ratio</div>
                <div class="exec-banner-val">{ret_share_pct:.2f}%</div>
            </div>
            """, unsafe_allow_html=True)
            
        with rb4:
            st.write("")
            if st.button("Focus Mode", key="focus_ret_mode", help="Direct focus on this group's data"):
                st.session_state['focus_ret_list'] = ret_data['buyFromVendorNo'].tolist()
                st.success(f"Focus Mode applied to top {len(ret_data)} vendors.")

        st.progress(min(max(ret_share_pct / 100.0, 0.0), 1.0))
        st.markdown("---")
    else: 
        st.write("Insufficient data to generate chart.")

    # MATRIX 1: Pending Returns
    st.subheader("Advanced Aging Summary Matrix - Pending Returns")

    active_vendors = df_exec_filtered[df_exec_filtered['Status'] == 'Collected']['buyFromVendorNo'].unique()
    pending_returns = df_exec_filtered[df_exec_filtered['Status'] == 'Pending for Collection'].copy()
    
    if not pending_returns.empty:
        if 'aging_days' not in pending_returns.columns and 'documentDate' in pending_returns.columns:
            pending_returns['documentDate'] = pd.to_datetime(pending_returns['documentDate'], errors='coerce')
            pending_returns['aging_days'] = (pd.to_datetime('today') - pending_returns['documentDate']).dt.days
        
        amt_col = [c for c in pending_returns.columns if 'amount' in c.lower() and 'vat' in c.lower()]
        target_amt_col = amt_col[0] if amt_col else ('amountIncludingVAT' if 'amountIncludingVAT' in pending_returns.columns else 'amount')
        pending_returns[target_amt_col] = pd.to_numeric(pending_returns[target_amt_col], errors='coerce').fillna(0.0)
        
        grand_total_pending = pending_returns[target_amt_col].sum()
        days_col = next((col for col in pending_returns.columns if str(col).lower().replace(' ', '').replace('_', '') in ['dayssincelastreceiving', 'sincelastreceiving', 'lastreceivingdays', 'dayssincereceiving']), None)

        agg_dict = {
            'buyFromVendorName': lambda x: x.iloc[0] if not x.empty else 'Unknown Vendor',
            target_amt_col: 'sum',
            'aging_days': 'max',
            'no': 'count'
        }
        if days_col: 
            agg_dict[days_col] = 'min'

        aging_summary = pending_returns.groupby('buyFromVendorNo').agg(agg_dict).reset_index()
        
        rename_map = {
            'buyFromVendorNo': 'Vendor Code', 
            'buyFromVendorName': 'Vendor Name', 
            target_amt_col: 'Total Amount', 
            'aging_days': 'Max_Age_Days', 
            'no': 'PRO Count'
        }
        if days_col: 
            rename_map[days_col] = 'Last_Receiving_Raw'
        aging_summary = aging_summary.rename(columns=rename_map)
        
        aging_summary['Vendor Name'] = aging_summary['Vendor Code'].map(vendor_lookup).fillna(aging_summary['Vendor Name'])
        aging_summary['Percentage'] = (aging_summary['Total Amount'] / grand_total_pending * 100) if grand_total_pending > 0 else 0.0
        aging_summary['Is_Active'] = aging_summary['Vendor Code'].isin(active_vendors)
        
        def get_status_alert(row):
            if row['Max_Age_Days'] <= 90: 
                return "Normal"
            elif row['Is_Active']: 
                return "Active Mitigation"
            else: 
                return "Requires Escalation"

        aging_summary['Commitment Status'] = aging_summary.apply(get_status_alert, axis=1)
        
        if 'Last_Receiving_Raw' in aging_summary.columns:
            def format_last_receiving(val):
                if pd.isna(val) or str(val).strip() == "" or str(val).lower() == "nan": 
                    return "N/A"
                try:
                    num = float(val)
                    return "0" if num == 0 else str(int(num))
                except: 
                    return "N/A"
            aging_summary['Last Receiving (Days)'] = aging_summary['Last_Receiving_Raw'].apply(format_last_receiving)
        else:
            aging_summary['Last Receiving (Days)'] = "N/A"

        sort_c1, sort_c2 = st.columns(2)
        with sort_c1:
            sort_choice = st.selectbox("Data Sort Order:", ["Total Amount (High to Low)", "Oldest Return Days (High to Low)", "PRO Count (High to Low)", "Vendor Name (A-Z)", "Commitment Status"], key="exec_sort_option")
        with sort_c2:
            display_limit = st.number_input("Records Limit:", min_value=5, max_value=200, value=top_aging_n, key="exec_limit_option")

        if "Total Amount" in sort_choice: 
            aging_summary = aging_summary.sort_values(by='Total Amount', ascending=False)
        elif "Oldest Return Days" in sort_choice: 
            aging_summary = aging_summary.sort_values(by='Max_Age_Days', ascending=False)
        elif "PRO Count" in sort_choice: 
            aging_summary = aging_summary.sort_values(by='PRO Count', ascending=False)
        elif "Vendor Name" in sort_choice: 
            aging_summary = aging_summary.sort_values(by='Vendor Name', ascending=True)
        elif "Commitment Status" in sort_choice: 
            aging_summary = aging_summary.sort_values(by=['Commitment Status', 'Total Amount'], ascending=[True, False])

        display_aging_table = aging_summary.head(display_limit).copy()
        table_view = display_aging_table[['Vendor Code', 'Vendor Name', 'PRO Count', 'Total Amount', 'Percentage', 'Last Receiving (Days)', 'Max_Age_Days', 'Commitment Status']].copy()
        
        def highlight_aging_cells(row):
            age = row['Max_Age_Days']
            if age > 45: 
                return ['background-color: rgba(239, 68, 68, 0.2); color: #ef4444; font-weight: 600'] * len(row)
            elif age > 20: 
                return ['background-color: rgba(245, 158, 11, 0.2); color: #f59e0b; font-weight: 600'] * len(row)
            else: 
                return ['background-color: rgba(16, 185, 129, 0.2); color: #10b981; font-weight: 600'] * len(row)

        st.dataframe(
            table_view.style.apply(highlight_aging_cells, axis=1),
            use_container_width=True, hide_index=True,
            column_config={
                "Total Amount": st.column_config.NumberColumn("Total Amount (SAR)", format="%,.2f SAR"),
                "Percentage": st.column_config.NumberColumn("Liability Share", format="%.2f%%"),
                "Max_Age_Days": st.column_config.NumberColumn("Peak Age (Days)"),
                "PRO Count": st.column_config.NumberColumn("Volume (PROs)")
            }
        )
        
        st.download_button(
            label="Download Aging Summary Matrix (CSV)",
            data=convert_df_to_csv(table_view),
            file_name=f"Executive_Aging_Summary_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
    else:
        st.write("No aging data applicable for the specified parameters.")

# ==========================================
# PAGE 5: VENDOR SLA & ESCALATION HUB (NEW)
# ==========================================
elif page == "Vendor SLA & Escalation Hub":
    st.markdown("<h2>Vendor SLA & Speed Scorecard Hub</h2>", unsafe_allow_html=True)
    st.write("Executive scorecard analyzing closed return TAT, problem root causes, and automated PO offsetting templates.")
    st.markdown("---")

    # ==========================================
    # Build the real closed-PRO analytical base.
    # The "Closed" sheet only tells us WHICH PROs are closed, WHEN they
    # were posted (postingDate) and their value. It has no reason or
    # opened-date column. The actual discrepancy reason
    # (NEAR_EXPIRY / MISSED_ITEM / QUALITY_ISSUE / NOT_LISTED /
    # NOT_ORDERED / PRICE_ISSUE) and the original issue-logged timestamp
    # live in 'pro_with_issues_linked_with_po_' (line-item level, keyed by
    # PRO number). This page's whole premise depends on joining the two.
    # ==========================================
    if not df_closed.empty:
        closed_base = df_closed.copy()
        closed_base['no'] = closed_base['no'].astype(str)

        if not df_issues.empty and 'no' in df_issues.columns:
            issue_detail = df_issues.copy()
            issue_detail['no'] = issue_detail['no'].astype(str)

            # Primary reason per PRO = the most frequent reason among its line items
            reason_per_pro = (
                issue_detail.groupby('no')['reason']
                .agg(lambda s: s.value_counts().idxmax() if s.notna().any() else pd.NA)
                .rename('reason')
            )

            # Earliest logged item-issue timestamp per PRO = when the case was opened
            if 'mi.created_at' in issue_detail.columns:
                issue_detail['mi.created_at'] = pd.to_datetime(issue_detail['mi.created_at'], errors='coerce')
                opened_at_per_pro = issue_detail.groupby('no')['mi.created_at'].min().rename('opened_at')
            else:
                opened_at_per_pro = pd.Series(dtype='datetime64[ns]', name='opened_at')

            issue_summary = pd.concat([reason_per_pro, opened_at_per_pro], axis=1).reset_index()
            closed_base = closed_base.merge(issue_summary, on='no', how='left')
        else:
            closed_base['reason'] = pd.NA
            closed_base['opened_at'] = pd.NaT

        # Real TAT = closing date (postingDate) - case opened date (from the issues log).
        if 'postingDate' in closed_base.columns:
            closed_base['postingDate'] = pd.to_datetime(closed_base['postingDate'], errors='coerce')
            closed_base['tat_days'] = (closed_base['postingDate'] - closed_base['opened_at']).dt.days
            closed_base.loc[closed_base['tat_days'] < 0, 'tat_days'] = 0  # guard bad/backdated entries
        else:
            closed_base['tat_days'] = pd.NA

        # For PROs we couldn't match to an issue record, fall back to the
        # portfolio's own average TAT (computed live, not a fixed constant)
        # rather than inventing a number.
        known_tat_mean = closed_base['tat_days'].mean()
        known_tat_mean = 0.0 if pd.isna(known_tat_mean) else known_tat_mean
        match_rate = closed_base['reason'].notna().mean() if len(closed_base) else 0.0
        closed_base['tat_days'] = closed_base['tat_days'].fillna(known_tat_mean)

        # A closed PRO with no matching line item in the issues log is NOT a
        # data-quality mystery — it falls into exactly one of two explainable
        # buckets, so we label it as such instead of a generic "UNSPECIFIED":
        #   1) LEGACY_PRO_PRE_TRACKING — its PRO number belongs to the old
        #      "PRO25.." numbering series, which predates the item-level
        #      issue-tracking sheet entirely (the sheet is ~99% "PRO26..").
        #   2) CLOSED_NO_ITEM_ISSUE_LOGGED — a normal "PRO26.." PRO that was
        #      closed without any discrepancy line item ever being recorded
        #      against it (e.g. a clean/full return with nothing missing).
        unresolved_reason_mask = closed_base['reason'].isna()
        legacy_mask = unresolved_reason_mask & closed_base['no'].str.match(r'^PRO25', na=False)
        closed_base.loc[legacy_mask, 'reason'] = 'LEGACY_PRO_PRE_TRACKING'
        closed_base.loc[unresolved_reason_mask & ~legacy_mask, 'reason'] = 'CLOSED_NO_ITEM_ISSUE_LOGGED'
    else:
        closed_base = pd.DataFrame(columns=['buyFromVendorNo', 'buyFromVendorName', 'no', 'tat_days',
                                             'amountIncludingVAT', 'reason', 'location_code'])
        match_rate = 0.0

    KNOWN_ITEM_REASONS = ['NEAR_EXPIRY', 'MISSED_ITEM', 'QUALITY_ISSUE', 'NOT_LISTED', 'NOT_ORDERED', 'PRICE_ISSUE']
    NO_REASON_LABELS = ['LEGACY_PRO_PRE_TRACKING', 'CLOSED_NO_ITEM_ISSUE_LOGGED',
                         'NOT_YET_LINKED_TO_PO', 'LINKED_PENDING_CLASSIFICATION']

    if 'amountIncludingVAT' not in closed_base.columns:
        amt_cols = [c for c in closed_base.columns if 'amount' in str(c).lower()]
        closed_base['amountIncludingVAT'] = pd.to_numeric(closed_base[amt_cols[0]], errors='coerce').fillna(0.0) if amt_cols else 0.0
    elif not closed_base.empty:
        closed_base['amountIncludingVAT'] = pd.to_numeric(closed_base['amountIncludingVAT'], errors='coerce').fillna(0.0)

    if 'location_code' not in closed_base.columns:
        closed_base['location_code'] = 'UNKNOWN'
    elif not closed_base.empty:
        closed_base['location_code'] = closed_base['location_code'].fillna('UNKNOWN')

    def assign_speed_tier(tat):
        if tat <= 7:
            return "Fast (<= 7 days)"
        elif tat <= 14:
            return "Moderate (8-14 days)"
        elif tat <= 30:
            return "Slow (15-30 days)"
        else:
            return "Critical (> 30 days)"

    if not closed_base.empty:
        closed_base['Speed_Tier'] = closed_base['tat_days'].apply(assign_speed_tier)
    else:
        closed_base['Speed_Tier'] = pd.Series(dtype='object')

    if closed_base.empty:
        st.warning("No closed-PRO data is currently available in the 'Closed' sheet, so the scorecard "
                   "and root-cause views below cannot be computed. The Smart Offsetting tab is unaffected.")

    # ==========================================
    # Build the OPEN / UNRESOLVED PRO base — PROs that were created
    # (a discrepancy was raised) but have NOT closed yet. This is the
    # backlog: it is not just an SLA statistic, it is money and vendor
    # accountability actively stuck in the pipeline right now.
    #
    # Source: "Pending PROs" (the full open backlog, current in-progress
    # age already given in 'Aging'). "Linked_With_PO" is confirmed to be a
    # strict subset of "Pending PROs" (every linked PRO is still pending) —
    # it marks PROs that have already been matched to an offsetting PO but
    # not yet fully closed, so we treat it as a status flag, not a
    # separate bucket.
    # ==========================================
    if not df_pending_pros.empty:
        open_base = df_pending_pros.copy()
        open_base['no'] = open_base['no'].astype(str)

        # Normalize the value column (source calls it AmountWithVAT here)
        if 'amountIncludingVAT' not in open_base.columns:
            amt_cols = [c for c in open_base.columns if 'amount' in str(c).lower() and 'vat' in str(c).lower()]
            open_base['amountIncludingVAT'] = pd.to_numeric(open_base[amt_cols[0]], errors='coerce').fillna(0.0) if amt_cols else 0.0
        else:
            open_base['amountIncludingVAT'] = pd.to_numeric(open_base['amountIncludingVAT'], errors='coerce').fillna(0.0)

        if 'location_code' not in open_base.columns:
            open_base['location_code'] = 'UNKNOWN'
        else:
            open_base['location_code'] = open_base['location_code'].fillna('UNKNOWN')

        # Current age = days elapsed since the PRO was opened and STILL not closed.
        open_base['age_days'] = pd.to_numeric(open_base['Aging'], errors='coerce').fillna(0) if 'Aging' in open_base.columns else 0

        linked_pro_nos = set(df_linked['no'].astype(str)) if (not df_linked.empty and 'no' in df_linked.columns) else set()
        open_base['Is_Linked_To_PO'] = open_base['no'].isin(linked_pro_nos)

        # Same real reason join as closed_base — reason only gets logged once
        # a PRO reaches the "linked with PO" stage, so most still-open,
        # not-yet-linked PROs genuinely have no reason logged yet.
        if not df_issues.empty and 'no' in df_issues.columns:
            issue_detail2 = df_issues.copy()
            issue_detail2['no'] = issue_detail2['no'].astype(str)
            reason_per_pro_open = (
                issue_detail2.groupby('no')['reason']
                .agg(lambda s: s.value_counts().idxmax() if s.notna().any() else pd.NA)
                .rename('reason')
            )
            open_base = open_base.merge(reason_per_pro_open, on='no', how='left')
        else:
            open_base['reason'] = pd.NA

        no_reason_open_mask = open_base['reason'].isna()
        open_base.loc[no_reason_open_mask & open_base['Is_Linked_To_PO'], 'reason'] = 'LINKED_PENDING_CLASSIFICATION'
        open_base.loc[no_reason_open_mask & ~open_base['Is_Linked_To_PO'], 'reason'] = 'NOT_YET_LINKED_TO_PO'

        # ------------------------------------------------------------------
        # SPLIT CREDIT-NOTE VERIFICATION RULE (per Ops team):
        # When a vendor splits one credit note into several partial notes
        # (one per item), the original PRO is deleted and replaced with a
        # new PRO per item. These replacement PROs have no record in the
        # item-issue log at all — they are "sourceless". To tell whether a
        # sourceless PRO is genuinely still open or was actually already
        # resolved this way, cross-check it against 'Linked_With_PO':
        #   - present in Linked_With_PO  -> genuinely still open (mid-process)
        #   - absent from Linked_With_PO -> presumed already closed; this
        #     'Pending PROs' row is a stale leftover, not real open work
        # This rule ONLY applies to sourceless PROs — a PRO with a real
        # matched reason went through the normal single-item path and its
        # pending status is trusted as-is.
        # ------------------------------------------------------------------
        open_base['Verified_Status'] = 'Still Open (Reason Logged)'
        open_base.loc[no_reason_open_mask & open_base['Is_Linked_To_PO'], 'Verified_Status'] = 'Still Open (Linked, Awaiting Completion)'
        open_base.loc[no_reason_open_mask & ~open_base['Is_Linked_To_PO'], 'Verified_Status'] = 'Presumed Closed (Stale Record - Split Credit Note)'

        open_base['Speed_Tier'] = open_base['age_days'].apply(assign_speed_tier)
    else:
        open_base = pd.DataFrame(columns=['no', 'buyFromVendorNo', 'buyFromVendorName', 'location_code',
                                           'amountIncludingVAT', 'age_days', 'Is_Linked_To_PO', 'reason',
                                           'Speed_Tier', 'Verified_Status'])

    # Same verification rule applied to the Closed side, as an integrity
    # check: a truly-closed PRO should NEVER still appear in Linked_With_PO
    # (which would mean, per the same rule, that it isn't really closed).
    if not closed_base.empty:
        linked_pro_nos_check = set(df_linked['no'].astype(str)) if (not df_linked.empty and 'no' in df_linked.columns) else set()
        closed_conflict_mask = closed_base['no'].isin(linked_pro_nos_check)
        closed_status_conflicts = int(closed_conflict_mask.sum())
    else:
        closed_status_conflicts = 0

    # ==========================================
    # Total PRO Creation universe = Closed + still-Open (verified disjoint —
    # a PRO number never appears in both sheets). This is "everything that
    # was ever opened against a vendor", used for the creation-side analysis.
    # ==========================================
    created_base = pd.concat([
        closed_base.assign(Current_Status='Closed')[['no', 'buyFromVendorNo', 'buyFromVendorName', 'reason', 'amountIncludingVAT', 'location_code', 'Current_Status']],
        open_base.assign(Current_Status='Open / Unresolved')[['no', 'buyFromVendorNo', 'buyFromVendorName', 'reason', 'amountIncludingVAT', 'location_code', 'Current_Status']],
    ], ignore_index=True) if not closed_base.empty or not open_base.empty else pd.DataFrame(
        columns=['no', 'buyFromVendorNo', 'buyFromVendorName', 'reason', 'amountIncludingVAT', 'location_code', 'Current_Status']
    )

    # Global Performance Executive Banner — computed live from the joined data, not hardcoded.
    total_closed = len(closed_base)
    total_open_raw = len(open_base)

    if total_open_raw:
        stale_mask = open_base['Verified_Status'] == 'Presumed Closed (Stale Record - Split Credit Note)'
    else:
        stale_mask = pd.Series(dtype=bool)
    stale_count = int(stale_mask.sum()) if total_open_raw else 0
    stale_value = open_base.loc[stale_mask, 'amountIncludingVAT'].sum() if total_open_raw else 0.0
    open_verified_base = open_base[~stale_mask].copy() if total_open_raw else open_base
    total_open = len(open_verified_base)

    total_created = total_closed + total_open
    avg_tat = closed_base['tat_days'].mean() if total_closed else 0.0
    critical_overdue_closed = int((closed_base['tat_days'] > 30).sum()) if total_closed else 0
    critical_overdue_open = int((open_verified_base['age_days'] > 30).sum()) if total_open else 0
    trapped_capital_closed = closed_base[closed_base['reason'].isin(['NEAR_EXPIRY', 'MISSED_ITEM'])]['amountIncludingVAT'].sum() if total_closed else 0.0
    open_value_at_risk = open_verified_base['amountIncludingVAT'].sum() if total_open else 0.0

    st.markdown("##### A. Closed PROs — resolution speed (this is history: already done)")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Closed PROs Analyzed", f"{total_closed:,}", f"{match_rate*100:.1f}% matched to a known item-level reason")
    k2.metric("Overall Average SLA (TAT)", f"{avg_tat:.1f} Days", "Avg Response Speed")
    k3.metric("Was Overdue (> 30 Days to Close)", f"{critical_overdue_closed:,} PROs", delta_color="inverse")
    k4.metric("Trapped Capital (Closed)", f"{trapped_capital_closed:,.2f} SAR", "NEAR_EXPIRY + MISSED_ITEM only")

    if closed_status_conflicts:
        st.warning(
            f"⚠️ Data integrity flag: {closed_status_conflicts:,} PRO(s) appear in **both** 'Closed' and "
            f"'Linked_With_PO'. Per the split-credit-note rule, presence in Linked_With_PO means a PRO is "
            f"NOT actually closed yet — review these before trusting the Closed-PRO totals above."
        )
    else:
        st.caption("✅ Closure-status cross-check passed: no PRO marked 'Closed' currently also appears in 'Linked_With_PO'.")

    st.markdown("##### B. Open & Unresolved PROs — the live backlog (this is now: still costing us)")
    o1, o2, o3, o4 = st.columns(4)
    o1.metric("Verified Open PROs", f"{total_open:,}", f"{total_open/total_created*100:.1f}% of everything ever created" if total_created else None)
    o2.metric("Value Stuck in Backlog", f"{open_value_at_risk:,.2f} SAR")
    o3.metric("Open > 30 Days (Critical)", f"{critical_overdue_open:,} PROs", delta_color="inverse")
    o4.metric("Oldest Open PRO", f"{int(open_verified_base['age_days'].max()) if total_open else 0} Days", "Still unresolved right now", delta_color="inverse")

    if stale_count:
        st.markdown(f"""
        <div class="alert-box" style="border-left: 4px solid #f59e0b;">
        <b>🔍 Source & Closure Verification (Split Credit-Note Rule):</b> The <code>Pending PROs</code> sheet lists
        <b>{total_open_raw:,}</b> raw records. Of these, <b>{stale_count:,}</b> have no matching item-level reason in
        <code>pro_with_issues_linked_with_po_</code> <i>and</i> are not present in <code>Linked_With_PO</code>. Per the
        vendor split-credit-note workflow (the original PRO is deleted and replaced by new, sourceless PROs once the
        vendor sends partial credit notes), these <b>{stale_count:,}</b> records — worth <b>{stale_value:,.2f} SAR</b> —
        are presumed <b>already resolved</b> and are stale leftovers in the sheet rather than real open work.
        The metrics above already exclude them (verified open = {total_open:,}, vs {total_open_raw:,} raw sheet rows).
        </div>
        """, unsafe_allow_html=True)
        with st.expander(f"View / export the {stale_count:,} presumed-closed stale records"):
            stale_cols = [c for c in ['no', 'buyFromVendorNo', 'buyFromVendorName', 'location_code', 'amountIncludingVAT', 'age_days'] if c in open_base.columns]
            st.dataframe(open_base.loc[stale_mask, stale_cols], use_container_width=True, hide_index=True)
            st.download_button(
                "Download Presumed-Closed Stale Records (CSV)",
                data=convert_df_to_csv(open_base.loc[stale_mask, stale_cols]),
                file_name=f"Presumed_Closed_Stale_PROs_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv",
                key="dl_stale_presumed_closed"
            )

    st.caption(
        f"**Total PROs ever created (Closed + verified Open): {total_created:,}** — a true, non-duplicated count of "
        f"the full workload after applying the closure-verification rule above."
    )

    if total_closed and match_rate < 1.0:
        st.caption(
            f"Reason coverage on closed PROs: {match_rate*100:.1f}% matched one of the 6 known discrepancy reasons "
            f"({', '.join(KNOWN_ITEM_REASONS)}). The remaining {(1-match_rate)*100:.1f}% is **not an unknown reason** — "
            f"it splits into two explainable, non-mysterious cases: a PRO closed with **no item-level issue ever logged** "
            f"against it (e.g. a clean/full return), or a PRO from the **old 'PRO25..' numbering series** that predates "
            f"the item-issue tracking sheet entirely. See Tab 2 for the exact breakdown."
        )

    st.markdown("---")

    sla_tab1, sla_tab2, sla_tab3, sla_tab4 = st.tabs([
        "1. Vendor SLA & Speed Scorecard", 
        "2. Root Cause & Value Exposure", 
        "3. PRO Creation Deep-Dive (Who & Why)",
        "4. Smart Offsetting & Escalation"
    ])

    # ------------------------------------------
    # TAB 1: VENDOR SLA & SPEED SCORECARD
    # ------------------------------------------
    with sla_tab1:
        st.subheader("Vendor SLA Performance & Turnaround Time (TAT) Scorecard")
        
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            tier_filter = st.multiselect(
                "Filter Speed Tier Category:", 
                options=["Fast (<= 7 days)", "Moderate (8-14 days)", "Slow (15-30 days)", "Critical (> 30 days)"],
                default=["Fast (<= 7 days)", "Moderate (8-14 days)", "Slow (15-30 days)", "Critical (> 30 days)"],
                key="sla_tier_filter"
            )
        with col_f2:
            search_vendor_sla = st.text_input("Search Vendor Name/Code in Scorecard:", key="sla_vendor_search").strip().lower()

        filtered_scorecard = closed_base[closed_base['Speed_Tier'].isin(tier_filter)].copy()
        if search_vendor_sla:
            filtered_scorecard = filtered_scorecard[
                filtered_scorecard['buyFromVendorNo'].str.lower().str.contains(search_vendor_sla) | 
                filtered_scorecard['buyFromVendorName'].str.lower().str.contains(search_vendor_sla)
            ]

        # Scorecard Aggregation
        if not filtered_scorecard.empty:
            scorecard_summary = filtered_scorecard.groupby(['buyFromVendorNo', 'buyFromVendorName']).agg({
                'tat_days': 'mean',
                'amountIncludingVAT': 'sum',
                'reason': 'count'
            }).reset_index()
            scorecard_summary.rename(columns={'reason': 'Closed_PRO_Count', 'tat_days': 'Avg_TAT_Days'}, inplace=True)
            scorecard_summary['Speed_Tier'] = scorecard_summary['Avg_TAT_Days'].apply(assign_speed_tier)
            scorecard_summary['Vendor_Label'] = scorecard_summary['buyFromVendorNo'] + " - " + scorecard_summary['buyFromVendorName']
            scorecard_summary = scorecard_summary.sort_values(by='Avg_TAT_Days', ascending=True)

            # Chart: Average TAT per Vendor
            fig_sla = px.bar(
                scorecard_summary,
                x='Avg_TAT_Days',
                y='Vendor_Label',
                orientation='h',
                color='Speed_Tier',
                color_discrete_map={
                    "Fast (<= 7 days)": "#10b981",
                    "Moderate (8-14 days)": "#0ea5e9",
                    "Slow (15-30 days)": "#f59e0b",
                    "Critical (> 30 days)": "#ef4444"
                },
                title="Average SLA Closure Speed (Days) by Vendor"
            )
            fig_sla.update_layout(
                yaxis_title="", 
                xaxis_title="Average Turnaround Time (Days)", 
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font=dict(family="Inter, sans-serif")
            )
            st.plotly_chart(fig_sla, use_container_width=True)

            # Table View
            st.markdown("### Vendor Scorecard Detailed Matrix")
            scorecard_view = scorecard_summary[['buyFromVendorNo', 'buyFromVendorName', 'Closed_PRO_Count', 'Avg_TAT_Days', 'Speed_Tier', 'amountIncludingVAT']].copy()
            
            def highlight_sla_tier(row):
                tier = row['Speed_Tier']
                if 'Fast' in tier:
                    return ['background-color: rgba(16, 185, 129, 0.15); color: #10b981; font-weight: 600'] * len(row)
                elif 'Moderate' in tier:
                    return ['background-color: rgba(14, 165, 233, 0.15); color: #0ea5e9; font-weight: 600'] * len(row)
                elif 'Slow' in tier:
                    return ['background-color: rgba(245, 158, 11, 0.15); color: #f59e0b; font-weight: 600'] * len(row)
                else:
                    return ['background-color: rgba(239, 68, 68, 0.15); color: #ef4444; font-weight: 600'] * len(row)

            st.dataframe(
                scorecard_view.style.apply(highlight_sla_tier, axis=1),
                use_container_width=True, hide_index=True,
                column_config={
                    "buyFromVendorNo": "Vendor Code",
                    "buyFromVendorName": "Vendor Name",
                    "Closed_PRO_Count": "Closed PROs",
                    "Avg_TAT_Days": st.column_config.NumberColumn("Avg TAT (Days)", format="%.2f Days"),
                    "amountIncludingVAT": st.column_config.NumberColumn("Total Closed Value", format="%,.2f SAR")
                }
            )

            st.download_button(
                label="Download Vendor SLA Scorecard (CSV)",
                data=convert_df_to_csv(scorecard_view),
                file_name=f"Vendor_SLA_Scorecard_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
        else:
            st.info("No scorecard data available matching the selected filters.")

    # ------------------------------------------
    # TAB 2: ROOT CAUSE & VALUE EXPOSURE
    # ------------------------------------------
    with sla_tab2:
        st.subheader("Discrepancy Root Causes & Location Value Exposure")
        st.write("Financial impact analysis linking discrepancy reasons (sourced from `pro_with_issues_linked_with_po_`) to warehouse locations.")

        if closed_base.empty:
            st.info("No closed-PRO data available to analyze root causes.")
        else:
            # Compute the real per-reason exposure first, so the highlight
            # box below reflects whatever is actually in the data instead
            # of two fixed reasons/amounts.
            reason_summary_raw = closed_base.groupby('reason').agg(
                Total_Exposure=('amountIncludingVAT', 'sum'),
                Affected_Vendors=('buyFromVendorNo', 'nunique'),
                Closed_PROs=('reason', 'count')
            ).reset_index().sort_values('Total_Exposure', ascending=False)

            grand_total = reason_summary_raw['Total_Exposure'].sum()
            top_causes = reason_summary_raw.head(2)
            top_causes_share = (top_causes['Total_Exposure'].sum() / grand_total * 100) if grand_total else 0.0

            rc_c1, rc_c2 = st.columns([1.2, 1])

            with rc_c1:
                rc_grouped = closed_base.groupby(['reason', 'location_code'])['amountIncludingVAT'].sum().reset_index()
                fig_rc = px.treemap(
                    rc_grouped,
                    path=['reason', 'location_code'],
                    values='amountIncludingVAT',
                    color='amountIncludingVAT',
                    color_continuous_scale='Reds',
                    title="Financial Exposure Breakdown by Problem Reason & Warehouse Location"
                )
                fig_rc.update_layout(margin=dict(l=0, r=0, t=30, b=0))
                st.plotly_chart(fig_rc, use_container_width=True)

            with rc_c2:
                st.markdown("#### Primary Exposure Key Highlights")
                highlight_rows = "".join(
                    f'<p style="margin:2px 0;"><b>{row.reason} Impact:</b> {row.Total_Exposure:,.2f} SAR</p>'
                    for row in top_causes.itertuples()
                )
                st.markdown(f"""
                <div class="alert-box" style="border-left: 4px solid #ef4444;">
                    {highlight_rows}
                    <p style="margin:6px 0; font-size:12px; color:#64748b;">
                    These top {len(top_causes)} root cause(s) constitute <b>{top_causes_share:.1f}% of total financial discrepancies</b> across warehouse operations.
                    </p>
                </div>
                """, unsafe_allow_html=True)

                reason_summary = reason_summary_raw.rename(columns={
                    'reason': 'Discrepancy Cause',
                    'Total_Exposure': 'Total Exposure (SAR)',
                    'Affected_Vendors': 'Affected Vendors',
                    'Closed_PROs': 'Closed PROs'
                })

                st.dataframe(
                    reason_summary,
                    use_container_width=True, hide_index=True,
                    column_config={"Total Exposure (SAR)": st.column_config.NumberColumn(format="%,.2f SAR")}
                )

    # ------------------------------------------
    # TAB 3: PRO CREATION DEEP-DIVE (WHO & WHY)
    # ------------------------------------------
    # This tab answers the operational questions directly:
    #   - Which vendor is CAUSING us the most PROs (by count)?
    #   - Which vendor is causing us the most financial exposure?
    #   - What is the single most common reason overall?
    #   - Is it getting better or worse over time?
    #   - Which warehouse locations are affected the most?
    # It works off the FULL item-level issues log (both closed AND still-open
    # PROs), not just the closed subset — this is "everything that was ever
    # opened", which is what "created" means.
    # ------------------------------------------
    with sla_tab3:
        st.subheader("PRO Creation Analysis — Who Causes Them, and Why")
        st.write(
            "Full breakdown of every discrepancy line item ever logged in `pro_with_issues_linked_with_po_` "
            "(covers PROs regardless of whether they are closed or still open), answering: which vendor "
            "creates the most PROs, which reason is most common, and where the exposure sits."
        )

        if df_issues.empty or 'no' not in df_issues.columns or 'reason' not in df_issues.columns:
            st.info("The 'pro_with_issues_linked_with_po_' sheet is not available, so this analysis cannot be computed.")
        else:
            issues_all = df_issues.copy()
            issues_all['no'] = issues_all['no'].astype(str)

            value_col = 'cost_inc_vat' if 'cost_inc_vat' in issues_all.columns else (
                'cost' if 'cost' in issues_all.columns else None
            )
            if value_col:
                issues_all[value_col] = pd.to_numeric(issues_all[value_col], errors='coerce').fillna(0.0)
            if 'mi.created_at' in issues_all.columns:
                issues_all['mi.created_at'] = pd.to_datetime(issues_all['mi.created_at'], errors='coerce')

            # ---- Filters ----
            f1, f2, f3 = st.columns(3)
            with f1:
                reason_opts = sorted(issues_all['reason'].dropna().unique().tolist())
                reason_pick = st.multiselect("Filter by Reason:", options=reason_opts, default=reason_opts, key="creation_reason_filter")
            with f2:
                vendor_search_creation = st.text_input("Search Vendor Name/Code:", key="creation_vendor_search").strip().lower()
            with f3:
                if 'location_code' in issues_all.columns:
                    loc_opts = sorted(issues_all['location_code'].dropna().unique().tolist())
                    loc_pick = st.multiselect("Filter by Warehouse:", options=loc_opts, default=loc_opts, key="creation_loc_filter")
                else:
                    loc_pick = None

            issues_f = issues_all[issues_all['reason'].isin(reason_pick)].copy()
            if loc_pick is not None:
                issues_f = issues_f[issues_f['location_code'].isin(loc_pick)]
            if vendor_search_creation:
                name_col = 'buyFromVendorName' if 'buyFromVendorName' in issues_f.columns else None
                code_mask = issues_f['buyFromVendorNo'].str.lower().str.contains(vendor_search_creation, na=False) if 'buyFromVendorNo' in issues_f.columns else False
                name_mask = issues_f[name_col].str.lower().str.contains(vendor_search_creation, na=False) if name_col else False
                issues_f = issues_f[code_mask | name_mask]

            if issues_f.empty:
                st.warning("No records match the selected filters.")
            else:
                distinct_pros = issues_f['no'].nunique()
                distinct_vendors = issues_f['buyFromVendorNo'].nunique() if 'buyFromVendorNo' in issues_f.columns else 0
                total_lines = len(issues_f)
                total_value = issues_f[value_col].sum() if value_col else 0.0
                top_reason_row = issues_f['reason'].value_counts()
                top_reason_name = top_reason_row.index[0] if len(top_reason_row) else "N/A"
                top_reason_share = (top_reason_row.iloc[0] / total_lines * 100) if total_lines else 0.0

                c1, c2, c3, c4 = st.columns(4)
                c1.metric("PROs Represented", f"{distinct_pros:,}", f"{total_lines:,} discrepancy line items")
                c2.metric("Vendors Involved", f"{distinct_vendors:,}")
                c3.metric("Most Common Reason", top_reason_name, f"{top_reason_share:.1f}% of all line items")
                c4.metric("Total Logged Exposure", f"{total_value:,.2f} SAR" if value_col else "N/A")

                st.markdown("---")

                dd1, dd2 = st.columns(2)

                with dd1:
                    st.markdown("#### Which reason is most common? (all 6 causes)")
                    reason_counts = issues_f['reason'].value_counts().reset_index()
                    reason_counts.columns = ['Reason', 'Line Items']
                    fig_reason = px.bar(
                        reason_counts, x='Line Items', y='Reason', orientation='h',
                        color='Line Items', color_continuous_scale='Reds',
                        title="Discrepancy Reason Frequency (Item-Level)"
                    )
                    fig_reason.update_layout(
                        yaxis_title="", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                        font=dict(family="Inter, sans-serif"), coloraxis_showscale=False
                    )
                    st.plotly_chart(fig_reason, use_container_width=True)

                with dd2:
                    st.markdown("#### Which vendor creates the most PROs?")
                    vendor_pro_counts = (
                        issues_f.groupby(['buyFromVendorNo', 'buyFromVendorName'])['no']
                        .nunique().reset_index(name='PRO_Count')
                        .sort_values('PRO_Count', ascending=False).head(10)
                    )
                    vendor_pro_counts['Vendor_Label'] = vendor_pro_counts['buyFromVendorNo'] + " - " + vendor_pro_counts['buyFromVendorName']
                    fig_vendor_count = px.bar(
                        vendor_pro_counts.sort_values('PRO_Count'), x='PRO_Count', y='Vendor_Label', orientation='h',
                        color='PRO_Count', color_continuous_scale='Oranges',
                        title="Top 10 Vendors by Number of PROs Caused"
                    )
                    fig_vendor_count.update_layout(
                        yaxis_title="", xaxis_title="Distinct PROs", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                        font=dict(family="Inter, sans-serif"), coloraxis_showscale=False
                    )
                    st.plotly_chart(fig_vendor_count, use_container_width=True)

                dd3, dd4 = st.columns(2)

                with dd3:
                    if value_col:
                        st.markdown("#### Which vendor has the highest financial exposure?")
                        vendor_value = (
                            issues_f.groupby(['buyFromVendorNo', 'buyFromVendorName'])[value_col]
                            .sum().reset_index(name='Exposure')
                            .sort_values('Exposure', ascending=False).head(10)
                        )
                        vendor_value['Vendor_Label'] = vendor_value['buyFromVendorNo'] + " - " + vendor_value['buyFromVendorName']
                        fig_vendor_value = px.bar(
                            vendor_value.sort_values('Exposure'), x='Exposure', y='Vendor_Label', orientation='h',
                            color='Exposure', color_continuous_scale='Purples',
                            title="Top 10 Vendors by Logged Financial Exposure (SAR)"
                        )
                        fig_vendor_value.update_layout(
                            yaxis_title="", xaxis_title="Exposure (SAR)", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                            font=dict(family="Inter, sans-serif"), coloraxis_showscale=False
                        )
                        st.plotly_chart(fig_vendor_value, use_container_width=True)

                with dd4:
                    if 'location_code' in issues_f.columns:
                        st.markdown("#### Which warehouse locations are affected most?")
                        loc_counts = issues_f['location_code'].value_counts().head(10).reset_index()
                        loc_counts.columns = ['Warehouse', 'Line Items']
                        fig_loc = px.bar(
                            loc_counts, x='Line Items', y='Warehouse', orientation='h',
                            color='Line Items', color_continuous_scale='Blues',
                            title="Top 10 Warehouse Locations by Issue Count"
                        )
                        fig_loc.update_layout(
                            yaxis_title="", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                            font=dict(family="Inter, sans-serif"), coloraxis_showscale=False
                        )
                        st.plotly_chart(fig_loc, use_container_width=True)

                if 'mi.created_at' in issues_f.columns and issues_f['mi.created_at'].notna().any():
                    st.markdown("#### Is this trending up or down over time?")
                    trend_src = issues_f.dropna(subset=['mi.created_at']).copy()
                    trend_src['Month'] = trend_src['mi.created_at'].dt.to_period('M').dt.to_timestamp()
                    trend = trend_src.groupby(['Month', 'reason']).size().reset_index(name='Line Items')
                    fig_trend = px.line(
                        trend, x='Month', y='Line Items', color='reason', markers=True,
                        title="Monthly Discrepancy Volume by Reason"
                    )
                    fig_trend.update_layout(
                        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                        font=dict(family="Inter, sans-serif")
                    )
                    st.plotly_chart(fig_trend, use_container_width=True)

                st.markdown("#### Vendor Root-Cause Matrix (top 15 vendors by PRO count × reason)")
                top15_vendor_codes = vendor_pro_counts.head(15)['buyFromVendorNo'] if not vendor_pro_counts.empty else pd.Series(dtype='object')
                matrix_src = issues_f[issues_f['buyFromVendorNo'].isin(top15_vendor_codes)]
                if not matrix_src.empty:
                    vendor_reason_matrix = (
                        matrix_src.groupby(['buyFromVendorNo', 'buyFromVendorName', 'reason'])['no']
                        .nunique().reset_index(name='PRO_Count')
                    )
                    pivot = vendor_reason_matrix.pivot_table(
                        index=['buyFromVendorNo', 'buyFromVendorName'], columns='reason', values='PRO_Count', fill_value=0
                    ).reset_index()
                    st.dataframe(pivot, use_container_width=True, hide_index=True)

                st.download_button(
                    label="Download Filtered PRO Creation Data (CSV)",
                    data=convert_df_to_csv(issues_f),
                    file_name=f"PRO_Creation_Analysis_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    key="dl_creation_analysis"
                )

    # ------------------------------------------
    # TAB 4: SMART OFFSETTING & ESCALATION
    # ------------------------------------------
    with sla_tab4:
        st.subheader("Smart Offsetting & Automated PO Escalation")
        st.write("Deduct unresolved pending return balances exceeding the 30-day threshold directly from new PO shipments.")

        pending_source = df_returns[df_returns['Status'] == 'Pending for Collection'].copy() if not df_returns.empty else pd.DataFrame()

        if not pending_source.empty:
            if 'documentDate' in pending_source.columns:
                pending_source['documentDate'] = pd.to_datetime(pending_source['documentDate'], errors='coerce')
                pending_source['age_days'] = (pd.to_datetime('today') - pending_source['documentDate']).dt.days
            else:
                pending_source['age_days'] = pd.NA

            overdue_pros = pending_source[pending_source['age_days'] > 30].copy()

            # 'Pending Returns' has no discrepancy-reason column of its own —
            # enrich it from 'pro_with_issues_linked_with_po_' (the same
            # source the Vendor SLA scorecard joins against), keyed by PRO
            # number, instead of leaving/faking a 'reason' field.
            if not overdue_pros.empty:
                if not df_issues.empty and 'no' in df_issues.columns and 'reason' in df_issues.columns:
                    reason_per_pro = (
                        df_issues.groupby('no')['reason']
                        .agg(lambda s: s.value_counts().idxmax() if s.notna().any() else pd.NA)
                        .rename('reason')
                    )
                    overdue_pros = overdue_pros.merge(reason_per_pro, on='no', how='left')
                if 'reason' not in overdue_pros.columns:
                    overdue_pros['reason'] = pd.NA
                overdue_pros['reason'] = overdue_pros['reason'].fillna('UNSPECIFIED')

                if 'location_code' not in overdue_pros.columns:
                    overdue_pros['location_code'] = 'UNKNOWN'
                else:
                    overdue_pros['location_code'] = overdue_pros['location_code'].fillna('UNKNOWN')

                if 'amountIncludingVAT' not in overdue_pros.columns:
                    amt_cols = [c for c in overdue_pros.columns if 'amount' in str(c).lower()]
                    overdue_pros['amountIncludingVAT'] = pd.to_numeric(overdue_pros[amt_cols[0]], errors='coerce').fillna(0.0) if amt_cols else 0.0
        else:
            overdue_pros = pd.DataFrame(columns=['no', 'buyFromVendorNo', 'buyFromVendorName', 'location_code',
                                                  'reason', 'amountIncludingVAT', 'age_days'])

        if overdue_pros.empty:
            st.success("No pending-return PROs currently exceed the 30-day resolution threshold.")
        else:
            st.warning(f"Critical Alert: Identified {len(overdue_pros)} PRO records exceeding the 30-day resolution threshold.")

            st.markdown("Select critical PRO items to generate direct deduction debit notes for Accounts & Procurement:")

            overdue_pros['Select_For_Deduction'] = True

            display_cols = ['Select_For_Deduction', 'no', 'buyFromVendorNo', 'buyFromVendorName',
                             'location_code', 'reason', 'amountIncludingVAT', 'age_days']
            display_cols = [c for c in display_cols if c in overdue_pros.columns]

            edited_df = st.data_editor(
                overdue_pros[display_cols],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Select_For_Deduction": st.column_config.CheckboxColumn("Deduct from PO?", default=True),
                    "no": "PRO Number",
                    "buyFromVendorNo": "Vendor Code",
                    "buyFromVendorName": "Vendor Name",
                    "location_code": "Warehouse",
                    "reason": "Discrepancy Cause",
                    "amountIncludingVAT": st.column_config.NumberColumn("Offset Amount (SAR)", format="%,.2f SAR"),
                    "age_days": st.column_config.NumberColumn("Aging (Days)")
                },
                key="deduction_editor"
            )

            selected_deductions = edited_df[edited_df['Select_For_Deduction'] == True].copy()

            if not selected_deductions.empty:
                total_offset_val = selected_deductions['amountIncludingVAT'].sum()

                e_col1, e_col2 = st.columns([2, 1])
                with e_col1:
                    st.success(f"Total Offsetting Value Selected for Direct Deduction: **{total_offset_val:,.2f} SAR** across **{len(selected_deductions)} PROs**.")

                with e_col2:
                    export_cols = ['no', 'buyFromVendorNo', 'buyFromVendorName', 'location_code', 'reason', 'amountIncludingVAT', 'age_days']
                    export_cols = [c for c in export_cols if c in selected_deductions.columns]
                    offset_export_df = selected_deductions[export_cols].copy()
                    offset_export_df['Action_Type'] = "Direct PO Deduction"
                    offset_export_df['Escalation_Status'] = "Critical Overdue > 30 Days"
                    offset_export_df['Processed_By'] = st.session_state.logged_in_user
                    offset_export_df['Timestamp'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                    csv_offset = convert_df_to_csv(offset_export_df)
                    st.download_button(
                        label="Export Direct PO Deduction Template (CSV)",
                        data=csv_offset,
                        file_name=f"PO_Direct_Deduction_Template_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv"
                    )

                if st.button("Register Offsetting Execution in Audit Trail", key="btn_register_offset"):
                    record_audit(
                        "Smart Offsetting Escalation",
                        f"Direct PO deduction executed for {len(selected_deductions)} PROs totaling {total_offset_val:,.2f} SAR",
                        st.session_state.logged_in_user
                    )
                    st.success("Offsetting action logged into system audit trail.")

# ==========================================
# PAGE 6: AUDIT TRAIL & LOGS
# ==========================================
elif page == "Audit Trail & Logs":
    st.markdown("<h2>System Audit Trail & Operations Log</h2>", unsafe_allow_html=True)
    st.write("Historical record of all gate stampings, attendance updates, and offsetting actions.")
    
    if len(st.session_state.audit_log) > 0:
        df_audit = pd.DataFrame(st.session_state.audit_log)
        st.dataframe(df_audit, use_container_width=True, hide_index=True)
        
        st.download_button(
            label="Download System Audit Trail (CSV)",
            data=convert_df_to_csv(df_audit),
            file_name=f"Audit_Trail_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )
    else:
        st.info("No system activity recorded in current session.")