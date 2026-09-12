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
@st.cache_data(ttl=1800, show_spinner="Loading workbook (this can take a while — 'Receiving Report per items' alone has 500K+ rows)...")
def load_all_sheets_live(file_path):
    if not os.path.exists(file_path): 
        return None, None, None, None, None, None, None, None
    
    # 'calamine' (python-calamine) reads large .xlsx sheets dramatically faster
    # than the default openpyxl engine — worth installing given the 500K+ row
    # 'Receiving Report per items' sheet. Falls back cleanly if not installed.
    try:
        xls = pd.ExcelFile(file_path, engine='calamine')
    except (ImportError, ValueError):
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
    # Full item-level receiving log (what actually arrived at the warehouse,
    # regardless of whether it was later flagged as a problem). Used to turn
    # raw "times returned" counts into a real rejection RATE (returned ÷
    # received), instead of an unanchored count.
    df_receiving = read_sheet(['Receiving Report per items', 'Receiving Report per Items', 'ReceivingReportPerItems', 'Receiving Report'])
    if not df_receiving.empty:
        df_receiving = df_receiving.rename(columns={
            'AL.item_no': 'item_no',
            'AH.location_code': 'location_code',
            'vendor_no': 'buyFromVendorNo',
            'AH.source_no': 'po_no',
            'completed_date': 'received_date',
        })

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
            elif clean_col in ['no', 'documentno', 'prono', 'number', 'purchasereturnorderno', 'documentnumber']:
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

    for df in [df_scheduled, df_returns, df_linked, df_pending_pros, df_closed, df_supplier_damage, df_issues, df_receiving]:
        if not df.empty and 'buyFromVendorNo' in df.columns:
            df['buyFromVendorNo'] = df['buyFromVendorNo'].apply(clean_vendor_code)
        if not df.empty and 'buyFromVendorName' in df.columns:
            df['buyFromVendorName'] = df['buyFromVendorName'].apply(clean_vendor_name)
            
    if not df_scheduled.empty and 'PO_Number' in df_scheduled.columns:
        df_scheduled['PO_Number'] = df_scheduled['PO_Number'].astype(str).str.strip().str.upper()
        
    for df in [df_returns, df_linked, df_pending_pros, df_closed, df_supplier_damage, df_issues]:
        if not df.empty and 'no' in df.columns:
            df['no'] = df['no'].astype(str).str.strip().str.upper()

    if not df_issues.empty and 'item_no' in df_issues.columns:
        df_issues['item_no'] = df_issues['item_no'].astype(str).str.strip()

    if not df_receiving.empty:
        if 'item_no' in df_receiving.columns:
            df_receiving['item_no'] = df_receiving['item_no'].astype(str).str.strip()
        if 'QTY' in df_receiving.columns:
            df_receiving['QTY'] = pd.to_numeric(df_receiving['QTY'], errors='coerce').fillna(0.0)
        if 'received_date' in df_receiving.columns:
            df_receiving['received_date'] = pd.to_datetime(df_receiving['received_date'], errors='coerce')

    return df_scheduled, df_returns, df_linked, df_pending_pros, df_closed, df_supplier_damage, df_issues, df_receiving

def calculate_target_amounts(df):
    if df.empty: 
        return 0.0, 0
    amt_col = [c for c in df.columns if 'amount' in c.lower() and 'vat' in c.lower()]
    if not amt_col:
        amt_col = [c for c in df.columns if 'amount' in c.lower()] 
    total_amt = pd.to_numeric(df[amt_col[0]], errors='coerce').sum() if amt_col else 0.0
    return total_amt, len(df)

df_scheduled, df_returns, df_linked, df_pending_pros, df_closed, df_supplier_damage, df_issues, df_receiving = load_all_sheets_live(EXCEL_FILE)

if df_scheduled is None:
    st.error(f"System Error: Target file '{EXCEL_FILE}' not found in current directory.")
    st.stop()

# ==========================================
# Vendor Dictionary Build
# ==========================================
all_codes = set()
for df in [df_scheduled, df_returns, df_linked, df_pending_pros, df_closed, df_supplier_damage, df_issues, df_receiving]:
    if not df.empty and 'buyFromVendorNo' in df.columns:
        all_codes.update(df['buyFromVendorNo'].dropna().unique())

vendor_lookup = {}
for df in [df_scheduled, df_returns, df_linked, df_pending_pros, df_closed, df_supplier_damage, df_issues, df_receiving]:
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
st.sidebar.caption(
    "Data is cached for 30 minutes to avoid re-reading the ~520K-row Receiving Report on every click. "
    "Use the button below to force a fresh read if the source file just changed."
)
if st.sidebar.button("🔄 Refresh Data Now", use_container_width=True):
    st.cache_data.clear()
    st.rerun()

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
    top_cn_n = st.sidebar.number_input("Top Variables Filter (CN Related To RTV):", min_value=1, max_value=100, value=10)
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

    # Chart 2: Outstanding CN related to PROs (Linked_With_PO)
    st.subheader(f"Top {top_cn_n} Vendors: Outstanding CN related to PROs")
    st.caption(
        "Credit notes/deductions linked to the PRO (Purchase Order discrepancy) process — sourced from "
        "`Linked_With_PO`, i.e. claims already matched to a new PO for deduction but not yet fully closed. "
        "This is the PRO-side equivalent of the RTV chart above."
    )

    if not df_linked_exec.empty:
        vat_cols_lnk = [c for c in df_linked_exec.columns if 'amount' in str(c).lower() and 'vat' in str(c).lower()]
        if not vat_cols_lnk:
            vat_cols_lnk = [c for c in df_linked_exec.columns if 'amount' in str(c).lower()]
        t_col_lnk = vat_cols_lnk[0] if vat_cols_lnk else 'amount'
        df_linked_exec['amountIncludingVAT'] = pd.to_numeric(df_linked_exec[t_col_lnk], errors='coerce').fillna(0.0) if t_col_lnk in df_linked_exec.columns else 0.0

    if not df_linked_exec.empty and 'buyFromVendorNo' in df_linked_exec.columns:
        pro_cn_data = df_linked_exec.groupby('buyFromVendorNo')['amountIncludingVAT'].sum().nlargest(top_cn_n).reset_index()
    else:
        pro_cn_data = pd.DataFrame()

    if not pro_cn_data.empty:
        pro_cn_data['Vendor_Label'] = pro_cn_data['buyFromVendorNo'].apply(get_vendor_label)
        fig1b = px.bar(
            pro_cn_data,
            x='amountIncludingVAT',
            y='Vendor_Label',
            orientation='h',
            color='amountIncludingVAT',
            color_continuous_scale=[[0, "#334155"], [1, "#a78bfa"]]
        )
        fig1b.update_layout(
            yaxis_title="",
            xaxis_title="Total Exposure (SAR)",
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Inter, sans-serif"),
            margin=dict(l=0, r=20, t=10, b=10)
        )
        fig1b.update_traces(hovertemplate='Vendor: %{y}<br>Amount: %{x:,.2f} SAR<extra></extra>')
        st.plotly_chart(fig1b, use_container_width=True)

        pro_cn_shown_val = pro_cn_data['amountIncludingVAT'].sum()
        pro_cn_total_system_val = df_linked_exec['amountIncludingVAT'].sum() if not df_linked_exec.empty else 0.0
        pro_cn_share_pct = (pro_cn_shown_val / pro_cn_total_system_val * 100) if pro_cn_total_system_val > 0 else 0.0

        st.markdown("<p style='font-size: 13px; font-weight: 700; margin-bottom: 6px;'> Performance Context </p>", unsafe_allow_html=True)
        pb1, pb2, pb3, pb4 = st.columns([1.2, 1.2, 1, 0.8])

        with pb1:
            st.markdown(f"""
            <div class="exec-banner-card">
                <div class="exec-banner-label">Selected Category Total (Top Vendors)</div>
                <div class="exec-banner-val">{pro_cn_shown_val:,.2f} SAR</div>
            </div>
            """, unsafe_allow_html=True)

        with pb2:
            st.markdown(f"""
            <div class="exec-banner-card">
                <div class="exec-banner-label">System Total</div>
                <div class="exec-banner-val">{pro_cn_total_system_val:,.2f} SAR</div>
            </div>
            """, unsafe_allow_html=True)

        with pb3:
            st.markdown(f"""
            <div class="exec-banner-card">
                <div class="exec-banner-label">Actual Representation Ratio</div>
                <div class="exec-banner-val">{pro_cn_share_pct:.2f}%</div>
            </div>
            """, unsafe_allow_html=True)

        with pb4:
            st.write("")
            if st.button("Focus Mode", key="focus_pro_cn_mode", help="Direct focus on this group's data"):
                st.session_state['focus_pro_cn_list'] = pro_cn_data['buyFromVendorNo'].tolist()
                st.success(f"Focus Mode applied to top {len(pro_cn_data)} vendors.")

        st.progress(min(max(pro_cn_share_pct / 100.0, 0.0), 1.0))
        st.markdown("---")
    else:
        st.write("Insufficient data to generate chart.")

    # Chart 3: Stagnant Pending Returns
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
    # MATRIX 2A: RTV & SUPPLIER DAMAGE CREDIT NOTES
    # ==========================================
    st.markdown("---")
    st.subheader("1. RTV Credit Notes Matrix")

    rtv_cn_list = []

    # Collected Returns (Withdrawn by Vendor - Pending CN)
    if not cn_collected_df.empty:
        c_tmp = cn_collected_df.copy()
        c_tmp['Category'] = 'Collected Return'
        d_col = 'postingDate' if 'postingDate' in c_tmp.columns else ('documentDate' if 'documentDate' in c_tmp.columns else None)
        c_tmp['ref_date'] = pd.to_datetime(c_tmp[d_col], errors='coerce') if d_col else pd.NaT
        rtv_cn_list.append(c_tmp)

    # Supplier Damage (Awaiting CN Settlement)
    if not df_damage_exec.empty:
        d_tmp = df_damage_exec.copy()
        d_tmp['Category'] = 'Supplier Damage'
        d_col = 'documentDate' if 'documentDate' in d_tmp.columns else ('postingDate' if 'postingDate' in d_tmp.columns else None)
        d_tmp['ref_date'] = pd.to_datetime(d_tmp[d_col], errors='coerce') if d_col else pd.NaT
        
        vat_cols_d = [c for c in d_tmp.columns if 'amount' in str(c).lower() and 'vat' in str(c).lower()]
        if not vat_cols_d:
            vat_cols_d = [c for c in d_tmp.columns if 'amount' in str(c).lower()]
        t_col_d = vat_cols_d[0] if vat_cols_d else 'amount'
        d_tmp['amountIncludingVAT'] = pd.to_numeric(d_tmp[t_col_d], errors='coerce').fillna(0.0)
        rtv_cn_list.append(d_tmp)

    if rtv_cn_list:
        combined_rtv_cn = pd.concat(rtv_cn_list, ignore_index=True)
        combined_rtv_cn['aging_days'] = (pd.to_datetime('today') - combined_rtv_cn['ref_date']).dt.days.fillna(0)
        
        grand_rtv_cn_total = combined_rtv_cn['amountIncludingVAT'].sum()
        
        rtv_summary = combined_rtv_cn.groupby('buyFromVendorNo').agg({
            'buyFromVendorName': lambda x: x.iloc[0] if not x.empty else 'Unknown Vendor',
            'amountIncludingVAT': 'sum',
            'aging_days': 'max',
            'no': 'count'
        }).reset_index().rename(columns={
            'buyFromVendorNo': 'Vendor Code',
            'buyFromVendorName': 'Vendor Name',
            'amountIncludingVAT': 'RTV CN Exposure',
            'aging_days': 'Max_Age_Days',
            'no': 'PRO / CN Count'
        })
        
        rtv_summary['Vendor Name'] = rtv_summary['Vendor Code'].map(vendor_lookup).fillna(rtv_summary['Vendor Name'])
        rtv_summary['Share %'] = (rtv_summary['RTV CN Exposure'] / grand_rtv_cn_total * 100) if grand_rtv_cn_total > 0 else 0.0
        
        rtv_summary['Status Alert'] = rtv_summary['Max_Age_Days'].apply(
            lambda age: "Normal" if age <= 20 else ("Pending Vendor Action" if age <= 45 else "Escalation Required")
        )

        r_c1, r_c2 = st.columns(2)
        with r_c1:
            sort_rtv = st.selectbox(
                "RTV Sort Order:", 
                ["RTV CN Exposure (High to Low)", "Oldest Days (High to Low)", "Record Count (High to Low)", "Vendor Name (A-Z)"], 
                key="exec_rtv_cn_sort_option"
            )
        with r_c2:
            limit_rtv = st.number_input("Records Limit:", min_value=5, max_value=200, value=top_aging_n, key="exec_rtv_cn_limit_option")

        if "RTV CN Exposure" in sort_rtv:
            rtv_summary = rtv_summary.sort_values(by='RTV CN Exposure', ascending=False)
        elif "Oldest Days" in sort_rtv:
            rtv_summary = rtv_summary.sort_values(by='Max_Age_Days', ascending=False)
        elif "Record Count" in sort_rtv:
            rtv_summary = rtv_summary.sort_values(by='PRO / CN Count', ascending=False)
        elif "Vendor Name" in sort_rtv:
            rtv_summary = rtv_summary.sort_values(by='Vendor Name', ascending=True)

        display_rtv_table = rtv_summary.head(limit_rtv)[['Vendor Code', 'Vendor Name', 'PRO / CN Count', 'RTV CN Exposure', 'Share %', 'Max_Age_Days', 'Status Alert']]

        def style_cells(row):
            age = row['Max_Age_Days']
            if age > 45:
                return ['background-color: rgba(239, 68, 68, 0.2); color: #ef4444; font-weight: 600'] * len(row)
            elif age > 20:
                return ['background-color: rgba(245, 158, 11, 0.2); color: #f59e0b; font-weight: 600'] * len(row)
            else:
                return ['background-color: rgba(16, 185, 129, 0.2); color: #10b981; font-weight: 600'] * len(row)

        st.dataframe(
            display_rtv_table.style.apply(style_cells, axis=1),
            use_container_width=True, hide_index=True,
            column_config={
                "RTV CN Exposure": st.column_config.NumberColumn("Total RTV Exposure (SAR)", format="%,.2f SAR"),
                "Share %": st.column_config.NumberColumn("Liability Share", format="%.2f%%"),
                "Max_Age_Days": st.column_config.NumberColumn("Peak Age (Days)")
            }
        )

        st.download_button(
            label="Download RTV Credit Notes Matrix (CSV)",
            data=convert_df_to_csv(display_rtv_table),
            file_name=f"Executive_RTV_Credit_Notes_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            key="dl_exec_rtv_cn_matrix"
        )
    else:
        st.info("No active RTV or Supplier Damage Credit Note liabilities found for the selected period.")

    # ==========================================
    # MATRIX 2B: PO DISCREPANCIES CREDIT NOTES (LINKED WITH PO)
    # ==========================================
    st.markdown("---")
    st.subheader("2.Credit Note Linked PO(Receiving under adjust)")

    if not df_linked_exec.empty:
        po_tmp = df_linked_exec.copy()
        
        d_col_po = 'documentDate' if 'documentDate' in po_tmp.columns else ('postingDate' if 'postingDate' in po_tmp.columns else None)
        po_tmp['ref_date'] = pd.to_datetime(po_tmp[d_col_po], errors='coerce') if d_col_po else pd.NaT
        po_tmp['aging_days'] = (pd.to_datetime('today') - po_tmp['ref_date']).dt.days.fillna(0)
        
        vat_cols_p = [c for c in po_tmp.columns if 'amount' in str(c).lower() and 'vat' in str(c).lower()]
        if not vat_cols_p:
            vat_cols_p = [c for c in po_tmp.columns if 'amount' in str(c).lower()]
        t_col_p = vat_cols_p[0] if vat_cols_p else 'amount'
        po_tmp['amountIncludingVAT'] = pd.to_numeric(po_tmp[t_col_p], errors='coerce').fillna(0.0)
        
        grand_po_cn_total = po_tmp['amountIncludingVAT'].sum()
        
        po_summary = po_tmp.groupby('buyFromVendorNo').agg({
            'buyFromVendorName': lambda x: x.iloc[0] if not x.empty else 'Unknown Vendor',
            'amountIncludingVAT': 'sum',
            'aging_days': 'max',
            'no': 'count'
        }).reset_index().rename(columns={
            'buyFromVendorNo': 'Vendor Code',
            'buyFromVendorName': 'Vendor Name',
            'amountIncludingVAT': 'PO Discrepancy Exposure',
            'aging_days': 'Max_Age_Days',
            'no': 'Discrepancy Records'
        })
        
        po_summary['Vendor Name'] = po_summary['Vendor Code'].map(vendor_lookup).fillna(po_summary['Vendor Name'])
        po_summary['Share %'] = (po_summary['PO Discrepancy Exposure'] / grand_po_cn_total * 100) if grand_po_cn_total > 0 else 0.0
        
        po_summary['Action Status'] = po_summary['Max_Age_Days'].apply(
            lambda age: "Recent Issue" if age <= 15 else ("Pending PO Deduction" if age <= 30 else "Critical Unsettled PO Discrepancy")
        )

        p_c1, p_c2 = st.columns(2)
        with p_c1:
            sort_po = st.selectbox(
                "PO Discrepancy Sort Order:", 
                ["PO Discrepancy Exposure (High to Low)", "Oldest Days (High to Low)", "Record Count (High to Low)", "Vendor Name (A-Z)"], 
                key="exec_po_cn_sort_option"
            )
        with p_c2:
            limit_po = st.number_input("Records Limit:", min_value=5, max_value=200, value=top_aging_n, key="exec_po_cn_limit_option")

        if "PO Discrepancy Exposure" in sort_po:
            po_summary = po_summary.sort_values(by='PO Discrepancy Exposure', ascending=False)
        elif "Oldest Days" in sort_po:
            po_summary = po_summary.sort_values(by='Max_Age_Days', ascending=False)
        elif "Record Count" in sort_po:
            po_summary = po_summary.sort_values(by='Discrepancy Records', ascending=False)
        elif "Vendor Name" in sort_po:
            po_summary = po_summary.sort_values(by='Vendor Name', ascending=True)

        display_po_table = po_summary.head(limit_po)[['Vendor Code', 'Vendor Name', 'Discrepancy Records', 'PO Discrepancy Exposure', 'Share %', 'Max_Age_Days', 'Action Status']]

        st.dataframe(
            display_po_table.style.apply(style_cells, axis=1),
            use_container_width=True, hide_index=True,
            column_config={
                "PO Discrepancy Exposure": st.column_config.NumberColumn("Total PO Discrepancy (SAR)", format="%,.2f SAR"),
                "Share %": st.column_config.NumberColumn("Liability Share", format="%.2f%%"),
                "Max_Age_Days": st.column_config.NumberColumn("Peak Age (Days)")
            }
        )

        st.download_button(
            label="Download PO Discrepancy Matrix (CSV)",
            data=convert_df_to_csv(display_po_table),
            file_name=f"Executive_PO_Discrepancy_Credit_Notes_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            key="dl_exec_po_cn_matrix"
        )
    else:
        st.info("No active Purchase Order discrepancy liabilities found for the selected period.")

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
            f"the item-issue tracking sheet entirely. See Tab 2 (PRO Creation Deep-Dive) for the full reason breakdown."
        )

    st.markdown("---")

    sla_tab1, sla_tab3, sla_tab5, sla_tab6, sla_tab7, sla_tab8 = st.tabs([
        "1. Vendor SLA & Speed Scorecard", 
        "2. PRO Creation Deep-Dive (Who & Why)",
        "3. Recurring Item Analysis",
        "4. Rejection Rate vs. Receiving",
        "5. Live Vendor Pickup Responsiveness",
        "6. Vendor 360° Reliability Scorecard"
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
            fig_sla.update_traces(hovertemplate='%{y}: %{x:,.1f} Days<extra></extra>')
            fig_sla.update_layout(
                yaxis_title="", 
                xaxis_title="Average Turnaround Time (Days)", 
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                font=dict(family="Inter, sans-serif"),
                xaxis=dict(tickformat=',')
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
    # TAB 2: PRO CREATION DEEP-DIVE (WHO & WHY)
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

            # 'cost' / 'cost_inc_vat' are PER-UNIT prices, not the line
            # total — the real financial exposure of a discrepancy line is
            # unit price x missing_quantity. Using cost_inc_vat alone
            # understates true exposure by ~60x on this data.
            unit_price_col = 'cost_inc_vat' if 'cost_inc_vat' in issues_all.columns else (
                'cost' if 'cost' in issues_all.columns else None
            )
            value_col = None
            if unit_price_col and 'missing_quantity' in issues_all.columns:
                issues_all[unit_price_col] = pd.to_numeric(issues_all[unit_price_col], errors='coerce').fillna(0.0)
                issues_all['missing_quantity'] = pd.to_numeric(issues_all['missing_quantity'], errors='coerce').fillna(0.0)
                issues_all['Line_Exposure'] = issues_all[unit_price_col] * issues_all['missing_quantity']
                value_col = 'Line_Exposure'
            elif unit_price_col:
                issues_all[unit_price_col] = pd.to_numeric(issues_all[unit_price_col], errors='coerce').fillna(0.0)
                value_col = unit_price_col
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
                c4.metric("Total Logged Exposure", f"{total_value:,.2f} SAR" if value_col else "N/A",
                          "Unit price x missing qty" if value_col == 'Line_Exposure' else None)

                st.markdown("---")

                dd1, dd2 = st.columns(2)

                with dd1:
                    st.markdown("#### Which reason is most common? (all 6 causes)")
                    reason_agg = issues_f.groupby('reason').agg(
                        Line_Items=('reason', 'count'),
                        PRO_Count=('no', 'nunique'),
                        Total_Value=(value_col, 'sum') if value_col else ('reason', 'count')
                    ).reset_index().sort_values('Line_Items', ascending=False)
                    fig_reason = px.bar(
                        reason_agg, x='Line_Items', y='reason', orientation='h',
                        color='Line_Items', color_continuous_scale='Reds',
                        title="Discrepancy Reason Frequency (Item-Level)",
                        custom_data=['PRO_Count', 'Total_Value'] if value_col else ['PRO_Count']
                    )
                    if value_col:
                        fig_reason.update_traces(
                            texttemplate='%{x:,} items | %{customdata[0]:,} PROs | %{customdata[1]:,.0f} SAR',
                            textposition='outside',
                            hovertemplate='%{y}<br>%{x:,} line items<br>%{customdata[0]:,} PROs<br>%{customdata[1]:,.2f} SAR<extra></extra>'
                        )
                    else:
                        fig_reason.update_traces(
                            texttemplate='%{x:,} items | %{customdata[0]:,} PROs',
                            textposition='outside',
                            hovertemplate='%{y}<br>%{x:,} line items<br>%{customdata[0]:,} PROs<extra></extra>'
                        )
                    fig_reason.update_layout(
                        yaxis_title="", xaxis_title="Line Items", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                        font=dict(family="Inter, sans-serif"), coloraxis_showscale=False,
                        xaxis=dict(tickformat=',', range=[0, reason_agg['Line_Items'].max() * 1.35])
                    )
                    st.plotly_chart(fig_reason, use_container_width=True)

                with dd2:
                    h2a, h2b = st.columns([3, 1])
                    with h2a:
                        st.markdown("#### Which vendor creates the most PROs?")
                    with h2b:
                        top_n_vc = st.number_input("Top N", min_value=3, max_value=50, value=10, step=1,
                                                    key="topn_vendor_count", label_visibility="collapsed")
                    vendor_pro_counts_full = (
                        issues_f.groupby(['buyFromVendorNo', 'buyFromVendorName'])
                        .agg(PRO_Count=('no', 'nunique'), Total_Value=(value_col, 'sum') if value_col else ('no', 'nunique'))
                        .reset_index().sort_values('PRO_Count', ascending=False)
                    )
                    vendor_pro_counts = vendor_pro_counts_full.head(top_n_vc).copy()
                    vendor_pro_counts['Vendor_Label'] = vendor_pro_counts['buyFromVendorNo'] + " - " + vendor_pro_counts['buyFromVendorName']
                    fig_vendor_count = px.bar(
                        vendor_pro_counts.sort_values('PRO_Count'), x='PRO_Count', y='Vendor_Label', orientation='h',
                        color='PRO_Count', color_continuous_scale='Oranges',
                        title=f"Top {top_n_vc} Vendors by Number of PROs Caused",
                        custom_data=['Total_Value'] if value_col else None
                    )
                    if value_col:
                        fig_vendor_count.update_traces(
                            texttemplate='%{x:,} PROs | %{customdata[0]:,.0f} SAR', textposition='outside',
                            hovertemplate='%{y}<br>%{x:,} PROs<br>%{customdata[0]:,.2f} SAR<extra></extra>'
                        )
                    else:
                        fig_vendor_count.update_traces(texttemplate='%{x:,} PROs', textposition='outside', hovertemplate='%{y}: %{x:,} PROs<extra></extra>')
                    fig_vendor_count.update_layout(
                        yaxis_title="", xaxis_title="Distinct PROs", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                        font=dict(family="Inter, sans-serif"), coloraxis_showscale=False,
                        xaxis=dict(tickformat=',', range=[0, vendor_pro_counts['PRO_Count'].max() * 1.4] if not vendor_pro_counts.empty else None)
                    )
                    st.plotly_chart(fig_vendor_count, use_container_width=True)

                dd3, dd4 = st.columns(2)

                with dd3:
                    if value_col:
                        h3a, h3b = st.columns([3, 1])
                        with h3a:
                            st.markdown("#### Which vendor has the highest financial exposure?")
                        with h3b:
                            top_n_vv = st.number_input("Top N", min_value=3, max_value=50, value=10, step=1,
                                                        key="topn_vendor_value", label_visibility="collapsed")
                        vendor_value_full = (
                            issues_f.groupby(['buyFromVendorNo', 'buyFromVendorName'])
                            .agg(Exposure=(value_col, 'sum'), PRO_Count=('no', 'nunique'))
                            .reset_index().sort_values('Exposure', ascending=False)
                        )
                        vendor_value = vendor_value_full.head(top_n_vv).copy()
                        vendor_value['Vendor_Label'] = vendor_value['buyFromVendorNo'] + " - " + vendor_value['buyFromVendorName']
                        fig_vendor_value = px.bar(
                            vendor_value.sort_values('Exposure'), x='Exposure', y='Vendor_Label', orientation='h',
                            color='Exposure', color_continuous_scale='Purples',
                            title=f"Top {top_n_vv} Vendors by Logged Financial Exposure (SAR)",
                            custom_data=['PRO_Count']
                        )
                        fig_vendor_value.update_traces(
                            texttemplate='%{x:,.0f} SAR | %{customdata[0]:,} PROs', textposition='outside',
                            hovertemplate='%{y}<br>%{x:,.2f} SAR<br>%{customdata[0]:,} PROs<extra></extra>'
                        )
                        fig_vendor_value.update_layout(
                            yaxis_title="", xaxis_title="Exposure (SAR)", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                            font=dict(family="Inter, sans-serif"), coloraxis_showscale=False,
                            xaxis=dict(tickformat=',', range=[0, vendor_value['Exposure'].max() * 1.4] if not vendor_value.empty else None)
                        )
                        st.plotly_chart(fig_vendor_value, use_container_width=True)

                with dd4:
                    if 'location_code' in issues_f.columns:
                        h4a, h4b = st.columns([3, 1])
                        with h4a:
                            st.markdown("#### Which warehouse locations are affected most?")
                        with h4b:
                            top_n_loc = st.number_input("Top N", min_value=3, max_value=50, value=10, step=1,
                                                         key="topn_location", label_visibility="collapsed")
                        loc_agg_full = (
                            issues_f.groupby('location_code')
                            .agg(Line_Items=('location_code', 'count'), PRO_Count=('no', 'nunique'),
                                 Total_Value=(value_col, 'sum') if value_col else ('location_code', 'count'))
                            .reset_index().sort_values('Line_Items', ascending=False)
                        )
                        loc_agg = loc_agg_full.head(top_n_loc).copy()
                        fig_loc = px.bar(
                            loc_agg.sort_values('Line_Items'), x='Line_Items', y='location_code', orientation='h',
                            color='Line_Items', color_continuous_scale='Blues',
                            title=f"Top {top_n_loc} Warehouse Locations by Issue Count",
                            custom_data=['PRO_Count', 'Total_Value'] if value_col else ['PRO_Count']
                        )
                        if value_col:
                            fig_loc.update_traces(
                                texttemplate='%{x:,} items | %{customdata[0]:,} PROs | %{customdata[1]:,.0f} SAR',
                                textposition='outside',
                                hovertemplate='%{y}<br>%{x:,} items<br>%{customdata[0]:,} PROs<br>%{customdata[1]:,.2f} SAR<extra></extra>'
                            )
                        else:
                            fig_loc.update_traces(
                                texttemplate='%{x:,} items | %{customdata[0]:,} PROs', textposition='outside',
                                hovertemplate='%{y}<br>%{x:,} items<br>%{customdata[0]:,} PROs<extra></extra>'
                            )
                        fig_loc.update_layout(
                            yaxis_title="", xaxis_title="Line Items", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                            font=dict(family="Inter, sans-serif"), coloraxis_showscale=False,
                            xaxis=dict(tickformat=',', range=[0, loc_agg['Line_Items'].max() * 1.5] if not loc_agg.empty else None)
                        )
                        st.plotly_chart(fig_loc, use_container_width=True)

                st.markdown("#### Warehouse Rejection Profile — % of Each Warehouse's Issues, by Reason")
                st.write(
                    "Each row sums to 100% — this shows the *mix* of reasons at each warehouse, not raw volume. "
                    "A warehouse with an unusually high share of one reason (compared to the others) often points "
                    "to a local process issue at that specific site, not a vendor-wide problem."
                )
                if 'location_code' in issues_f.columns and not issues_f.empty:
                    wh_reason_counts = pd.crosstab(issues_f['location_code'], issues_f['reason'])
                    wh_reason_pct = wh_reason_counts.div(wh_reason_counts.sum(axis=1), axis=0) * 100
                    wh_reason_pct = wh_reason_pct.round(1).reset_index().rename(columns={'location_code': 'Warehouse'})

                    pct_cols = [c for c in wh_reason_pct.columns if c != 'Warehouse']
                    st.dataframe(
                        wh_reason_pct, use_container_width=True, hide_index=True,
                        column_config={c: st.column_config.NumberColumn(c, format="%.1f%%") for c in pct_cols}
                    )

                    # Flag the single biggest outlier automatically: the
                    # warehouse+reason cell furthest above that reason's
                    # average share across all other warehouses.
                    overall_share = wh_reason_counts.sum(axis=0) / wh_reason_counts.sum(axis=0).sum() * 100
                    diffs = wh_reason_pct.set_index('Warehouse')[pct_cols].subtract(overall_share, axis=1)
                    if not diffs.empty and diffs.to_numpy().size:
                        worst_wh, worst_reason = diffs.stack().idxmax()
                        worst_val = wh_reason_pct.set_index('Warehouse').loc[worst_wh, worst_reason]
                        baseline_val = overall_share[worst_reason]
                        st.info(
                            f"📍 **Biggest local anomaly**: at **{worst_wh}**, **{worst_reason}** makes up "
                            f"**{worst_val:.1f}%** of all issues there, vs **{baseline_val:.1f}%** on average across "
                            f"other warehouses — worth a site-specific review rather than a vendor-wide one."
                        )

                    st.download_button(
                        "Download Warehouse Rejection Profile (CSV)",
                        data=convert_df_to_csv(wh_reason_pct),
                        file_name=f"Warehouse_Rejection_Profile_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv",
                        key="dl_warehouse_profile"
                    )

                if 'mi.created_at' in issues_f.columns and issues_f['mi.created_at'].notna().any():
                    st.markdown("#### Is this trending up or down over time?")
                    trend_src = issues_f.dropna(subset=['mi.created_at']).copy()
                    trend_src['Month'] = trend_src['mi.created_at'].dt.to_period('M').dt.to_timestamp()
                    trend = trend_src.groupby(['Month', 'reason']).size().reset_index(name='Line Items')
                    fig_trend = px.line(
                        trend, x='Month', y='Line Items', color='reason', markers=True,
                        title="Monthly Discrepancy Volume by Reason"
                    )
                    fig_trend.update_traces(hovertemplate='%{x}: %{y:,}<extra></extra>')
                    fig_trend.update_layout(
                        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                        font=dict(family="Inter, sans-serif"),
                        yaxis=dict(tickformat=',')
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

                    # -- Still-open PROs per vendor, per 'Linked_With_PO' --
                    linked_pro_nos_matrix = (
                        set(df_linked['no'].astype(str).str.strip().str.upper())
                        if (not df_linked.empty and 'no' in df_linked.columns) else set()
                    )
                    open_pending_per_vendor = (
                        matrix_src[matrix_src['no'].isin(linked_pro_nos_matrix)]
                        .groupby('buyFromVendorNo')['no'].nunique()
                    )

                    # -- Total closed amount per vendor (from the Closed sheet) --
                    closed_amount_per_vendor = (
                        closed_base.groupby('buyFromVendorNo')['amountIncludingVAT'].sum()
                        if not closed_base.empty else pd.Series(dtype=float)
                    )

                    # -- Total logged exposure per vendor (unit price x missing qty), full population --
                    exposure_per_vendor = (
                        issues_f.groupby('buyFromVendorNo')[value_col].sum()
                        if value_col else pd.Series(dtype=float)
                    )

                    pivot['Open_Pending_PROs (Linked_With_PO)'] = pivot['buyFromVendorNo'].map(open_pending_per_vendor).fillna(0).astype(int)
                    pivot['Closed_Amount_SAR'] = pivot['buyFromVendorNo'].map(closed_amount_per_vendor).fillna(0.0)
                    pivot['Total_Logged_Exposure_SAR'] = pivot['buyFromVendorNo'].map(exposure_per_vendor).fillna(0.0)
                    pivot['Closed_%_of_Exposure'] = pivot.apply(
                        lambda r: (r['Closed_Amount_SAR'] / r['Total_Logged_Exposure_SAR'] * 100) if r['Total_Logged_Exposure_SAR'] > 0 else 0.0,
                        axis=1
                    )

                    st.caption(
                        "`Open_Pending_PROs`: distinct PRO numbers for this vendor currently sitting in `Linked_With_PO` "
                        "(matched to a PO for offsetting but not yet closed). `Closed_Amount_SAR`: this vendor's total "
                        "closed value from the `Closed` sheet (all-time, not limited to the reason filter above). "
                        "`Closed_%_of_Exposure`: that closed amount as a share of the vendor's Total Logged Exposure "
                        "(unit price × missing quantity, from the filtered discrepancy log). Note this ratio can exceed "
                        "100% — `Closed_Amount_SAR` covers the vendor's full closed-PRO value (including PROs that were "
                        "never logged with an item-level reason), while `Total_Logged_Exposure_SAR` only reflects line "
                        "items that were specifically logged in `pro_with_issues_linked_with_po_`, so it will often "
                        "under-cover the true closed value — a ratio above 100% is expected, not an error."
                    )

                    st.dataframe(
                        pivot, use_container_width=True, hide_index=True,
                        column_config={
                            "Open_Pending_PROs (Linked_With_PO)": st.column_config.NumberColumn(format="%,d"),
                            "Closed_Amount_SAR": st.column_config.NumberColumn("Closed Amount (SAR)", format="%,.2f"),
                            "Total_Logged_Exposure_SAR": st.column_config.NumberColumn("Total Logged Exposure (SAR)", format="%,.2f"),
                            "Closed_%_of_Exposure": st.column_config.NumberColumn("Closed % of Exposure", format="%.1f%%"),
                        }
                    )

                st.download_button(
                    label="Download Filtered PRO Creation Data (CSV)",
                    data=convert_df_to_csv(issues_f),
                    file_name=f"PRO_Creation_Analysis_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    key="dl_creation_analysis"
                )

    # ------------------------------------------
    # TAB 3: RECURRING ITEM ANALYSIS
    # ------------------------------------------
    # Answers: does a specific ITEM keep coming back as a return? For which
    # reason(s)? Does the same VENDOR repeat the same problem with that item,
    # and how many times? Is there a detectable timing pattern (e.g. this
    # item recurs roughly every N days from this vendor)?
    # ------------------------------------------
    with sla_tab5:
        st.subheader("Recurring Item Analysis — Which SKUs Keep Coming Back, and Why")
        st.write(
            "Item-level recurrence analysis from `pro_with_issues_linked_with_po_`: which specific items are "
            "returned again and again, their dominant reason, whether the same vendor repeats the same problem "
            "with that item, and whether there's a visible timing pattern."
        )

        if df_issues.empty or 'item_no' not in df_issues.columns:
            st.info("The 'pro_with_issues_linked_with_po_' sheet has no item-level detail, so recurrence analysis cannot be computed.")
        else:
            item_src = df_issues.copy()
            item_src['no'] = item_src['no'].astype(str)
            item_src['item_no'] = item_src['item_no'].astype(str).replace('nan', pd.NA)
            item_src = item_src.dropna(subset=['item_no'])

            value_col_ri = 'cost_inc_vat' if 'cost_inc_vat' in item_src.columns else None
            qty_col_ri = 'missing_quantity' if 'missing_quantity' in item_src.columns else None
            if value_col_ri:
                item_src[value_col_ri] = pd.to_numeric(item_src[value_col_ri], errors='coerce').fillna(0.0)
            if qty_col_ri:
                item_src[qty_col_ri] = pd.to_numeric(item_src[qty_col_ri], errors='coerce').fillna(0.0)
            item_src['Line_Exposure'] = (item_src[value_col_ri] * item_src[qty_col_ri]) if (value_col_ri and qty_col_ri) else 0.0
            if 'mi.created_at' in item_src.columns:
                item_src['mi.created_at'] = pd.to_datetime(item_src['mi.created_at'], errors='coerce')
            item_name_col = 'mi.name' if 'mi.name' in item_src.columns else None

            # ---- Filters ----
            rf1, rf2, rf3 = st.columns(3)
            with rf1:
                min_recur = st.number_input("Minimum times returned:", min_value=2, max_value=50, value=3, step=1, key="ri_min_recur")
            with rf2:
                ri_reason_opts = sorted(item_src['reason'].dropna().unique().tolist())
                ri_reason_pick = st.multiselect("Filter by Reason:", options=ri_reason_opts, default=ri_reason_opts, key="ri_reason_filter")
            with rf3:
                ri_vendor_search = st.text_input("Search Vendor Name/Code:", key="ri_vendor_search").strip().lower()

            item_f = item_src[item_src['reason'].isin(ri_reason_pick)].copy()
            if ri_vendor_search:
                code_m = item_f['buyFromVendorNo'].str.lower().str.contains(ri_vendor_search, na=False) if 'buyFromVendorNo' in item_f.columns else False
                name_m = item_f['buyFromVendorName'].str.lower().str.contains(ri_vendor_search, na=False) if 'buyFromVendorName' in item_f.columns else False
                item_f = item_f[code_m | name_m]

            if item_f.empty:
                st.warning("No records match the selected filters.")
            else:
                # ---- Per-item recurrence summary ----
                def _mode_or_na(s):
                    return s.value_counts().idxmax() if s.notna().any() and len(s) else pd.NA

                agg_dict = {
                    'Occurrences': ('item_no', 'count'),
                    'Distinct_PROs': ('no', 'nunique'),
                    'Distinct_Vendors': ('buyFromVendorNo', 'nunique'),
                    'Dominant_Reason': ('reason', _mode_or_na),
                    'Dominant_Vendor_Code': ('buyFromVendorNo', _mode_or_na),
                    'Dominant_Vendor_Name': ('buyFromVendorName', _mode_or_na),
                    'Total_Exposure_SAR': ('Line_Exposure', 'sum'),
                }
                if item_name_col:
                    agg_dict['Item_Name'] = (item_name_col, _mode_or_na)
                if 'mi.created_at' in item_f.columns:
                    agg_dict['First_Seen'] = ('mi.created_at', 'min')
                    agg_dict['Last_Seen'] = ('mi.created_at', 'max')

                item_summary = item_f.groupby('item_no').agg(**agg_dict).reset_index()

                if 'First_Seen' in item_summary.columns and 'Last_Seen' in item_summary.columns:
                    item_summary['Span_Days'] = (item_summary['Last_Seen'] - item_summary['First_Seen']).dt.days
                    item_summary['Avg_Days_Between'] = item_summary.apply(
                        lambda r: (r['Span_Days'] / (r['Occurrences'] - 1)) if r['Occurrences'] > 1 else pd.NA, axis=1
                    )

                # ---- Pull in the vendor's closure-speed classification from the
                # Vendor SLA Scorecard (Tab 1), so we can see whether a vendor with
                # a recurring-item problem is also slow to close the credit note
                # (compounding issue) or fast (contained issue). ----
                if not closed_base.empty:
                    vendor_speed_lookup = closed_base.groupby('buyFromVendorNo')['tat_days'].mean()
                    item_summary['Vendor_Avg_TAT_Days'] = item_summary['Dominant_Vendor_Code'].map(vendor_speed_lookup)
                    item_summary['Vendor_Speed_Tier'] = item_summary['Vendor_Avg_TAT_Days'].apply(
                        lambda v: assign_speed_tier(v) if pd.notna(v) else 'No Closed-PRO History'
                    )
                else:
                    item_summary['Vendor_Avg_TAT_Days'] = pd.NA
                    item_summary['Vendor_Speed_Tier'] = 'No Closed-PRO History'

                recurring_items = item_summary[item_summary['Occurrences'] >= min_recur].sort_values('Occurrences', ascending=False)

                total_items = item_summary['item_no'].nunique()
                recurring_count = len(recurring_items)
                recurring_exposure = recurring_items['Total_Exposure_SAR'].sum()
                top_item_occ = int(recurring_items['Occurrences'].iloc[0]) if not recurring_items.empty else 0

                rc1, rc2, rc3, rc4 = st.columns(4)
                rc1.metric("Distinct Items Involved", f"{total_items:,}")
                rc2.metric(f"Items Returned ≥ {min_recur} Times", f"{recurring_count:,}",
                           f"{recurring_count/total_items*100:.1f}% of all items" if total_items else None)
                rc3.metric("Most-Returned Item — Times", f"{top_item_occ:,}")
                rc4.metric("Exposure Tied to Recurring Items", f"{recurring_exposure:,.2f} SAR")

                st.markdown("---")

                if recurring_items.empty:
                    st.info(f"No item was returned {min_recur}+ times under the current filters. Try lowering the threshold.")
                else:
                    ch1, ch2 = st.columns([3, 1])
                    with ch1:
                        st.markdown("#### Most Frequently Returned Items")
                    with ch2:
                        top_n_ri = st.number_input("Top N", min_value=3, max_value=50, value=10, step=1,
                                                    key="topn_recurring_items", label_visibility="collapsed")

                    chart_items = recurring_items.head(top_n_ri).copy()
                    label_col = 'Item_Name' if 'Item_Name' in chart_items.columns else 'item_no'
                    chart_items['Item_Label'] = chart_items['item_no'].astype(str) + " - " + chart_items[label_col].astype(str).str.slice(0, 40)

                    fig_recur = px.bar(
                        chart_items.sort_values('Occurrences'), x='Occurrences', y='Item_Label', orientation='h',
                        color='Occurrences', color_continuous_scale='Reds',
                        title=f"Top {top_n_ri} Most-Returned Items (≥ {min_recur} times)",
                        custom_data=['Distinct_PROs', 'Total_Exposure_SAR', 'Dominant_Reason', 'Dominant_Vendor_Name']
                    )
                    fig_recur.update_traces(
                        texttemplate='%{x:,} times | %{customdata[1]:,.0f} SAR',
                        textposition='outside',
                        hovertemplate='%{y}<br>Returned %{x:,} times (%{customdata[0]:,} PROs)<br>Exposure: %{customdata[1]:,.2f} SAR'
                                      '<br>Dominant reason: %{customdata[2]}<br>Dominant vendor: %{customdata[3]}<extra></extra>'
                    )
                    fig_recur.update_layout(
                        yaxis_title="", xaxis_title="Times Returned", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                        font=dict(family="Inter, sans-serif"), coloraxis_showscale=False,
                        xaxis=dict(tickformat=',', range=[0, chart_items['Occurrences'].max() * 1.4])
                    )
                    st.plotly_chart(fig_recur, use_container_width=True)

                    st.markdown("#### Full Recurring-Item Detail")
                    st.caption(
                        "`Occurrences` and `Total_Exposure_SAR` here count **every reason combined** for this item. "
                        "`Dominant_Reason` is just the *most common* reason, not the only one — so these totals can be "
                        "slightly higher than the same item's row in the 'Repeat Offenders' table below, which counts "
                        "one specific reason only. A small gap between the two usually means this item occasionally "
                        "recurs for a different reason than its usual one."
                    )
                    st.caption(
                        "`Avg_Days_Between`: average gap between consecutive returns of this item — a small, "
                        "consistent number suggests a regular, systemic pattern rather than a one-off coincidence. "
                        "`Vendor_Speed_Tier` / `Vendor_Avg_TAT_Days`: pulled from the Vendor SLA Scorecard (Tab 1) for "
                        "this item's dominant vendor — this tells you whether a vendor with a recurring-item problem "
                        "is at least resolving it fast (Fast/Moderate), or is *also* slow to close the credit note on "
                        "top of the recurring issue (Slow/Critical) — a compounding problem worth escalating first."
                    )
                    detail_cols = [c for c in ['item_no', 'Item_Name', 'Occurrences', 'Distinct_PROs', 'Distinct_Vendors',
                                                'Dominant_Reason', 'Dominant_Vendor_Code', 'Dominant_Vendor_Name',
                                                'Vendor_Speed_Tier', 'Vendor_Avg_TAT_Days', 'Total_Exposure_SAR',
                                                'First_Seen', 'Last_Seen', 'Span_Days', 'Avg_Days_Between'] if c in recurring_items.columns]

                    def _highlight_vendor_speed(row):
                        tier = row.get('Vendor_Speed_Tier', '')
                        if 'Fast' in tier:
                            style = 'background-color: rgba(16, 185, 129, 0.15); color: #10b981; font-weight: 600'
                        elif 'Moderate' in tier:
                            style = 'background-color: rgba(14, 165, 233, 0.15); color: #0ea5e9; font-weight: 600'
                        elif 'Slow' in tier:
                            style = 'background-color: rgba(245, 158, 11, 0.15); color: #f59e0b; font-weight: 600'
                        elif 'Critical' in tier:
                            style = 'background-color: rgba(239, 68, 68, 0.15); color: #ef4444; font-weight: 600'
                        else:
                            style = ''
                        return [style if col == 'Vendor_Speed_Tier' else '' for col in row.index]

                    st.dataframe(
                        recurring_items[detail_cols].style.apply(_highlight_vendor_speed, axis=1),
                        use_container_width=True, hide_index=True,
                        column_config={
                            "Total_Exposure_SAR": st.column_config.NumberColumn("Total Exposure (SAR)", format="%,.2f"),
                            "First_Seen": st.column_config.DateColumn(format="YYYY-MM-DD"),
                            "Last_Seen": st.column_config.DateColumn(format="YYYY-MM-DD"),
                            "Avg_Days_Between": st.column_config.NumberColumn("Avg Days Between", format="%.1f"),
                            "Vendor_Avg_TAT_Days": st.column_config.NumberColumn("Vendor Avg TAT (Days)", format="%.1f"),
                            "Dominant_Vendor_Code": "Vendor Code",
                            "Dominant_Vendor_Name": "Vendor Name",
                        }
                    )

                    worst_combo = recurring_items[recurring_items['Vendor_Speed_Tier'].isin(['Slow (15-30 days)', 'Critical (> 30 days)'])]
                    if not worst_combo.empty:
                        st.warning(
                            f"⚠️ **{len(worst_combo):,} recurring item(s)** are tied to vendors who are *also* Slow/Critical "
                            f"on closure speed — a recurring problem that also takes a long time to resolve. These are the "
                            f"top escalation priority."
                        )

                    st.markdown("---")
                    st.markdown("#### Repeat Offenders: Vendor × Item × Reason")
                    st.write("Same vendor, same item, same reason — repeated. This is the clearest signal of a systemic (not random) problem.")

                    combo = (
                        item_f.groupby(['buyFromVendorNo', 'buyFromVendorName', 'item_no'] + ([item_name_col] if item_name_col else []) + ['reason'])
                        .agg(Times_Repeated=('item_no', 'count'), Total_Exposure_SAR=('Line_Exposure', 'sum'))
                        .reset_index()
                    )
                    combo = combo[combo['Times_Repeated'] >= min_recur].sort_values('Times_Repeated', ascending=False)

                    if combo.empty:
                        st.info(f"No single vendor repeated the same reason for the same item {min_recur}+ times under the current filters.")
                    else:
                        combo_cols = [c for c in ['buyFromVendorNo', 'buyFromVendorName', 'item_no', item_name_col, 'reason',
                                                   'Times_Repeated', 'Total_Exposure_SAR'] if c and c in combo.columns]
                        st.dataframe(
                            combo[combo_cols].head(50), use_container_width=True, hide_index=True,
                            column_config={
                                "Total_Exposure_SAR": st.column_config.NumberColumn("Total Exposure (SAR)", format="%,.2f"),
                                "Times_Repeated": st.column_config.NumberColumn(format="%,d"),
                            }
                        )

                    st.markdown("---")
                    st.markdown("#### Timing Pattern — When Does a Specific Item Recur?")
                    pick_options = (chart_items['item_no'] + " - " + chart_items[label_col].astype(str).str.slice(0, 50)).tolist()
                    picked = st.selectbox("Pick an item to see its return timeline:", options=pick_options, key="ri_timeline_pick")
                    picked_item_no = picked.split(" - ")[0]

                    timeline_src = item_f[(item_f['item_no'] == picked_item_no) & item_f['mi.created_at'].notna()].sort_values('mi.created_at')
                    if timeline_src.empty:
                        st.info("No dated records available for this item.")
                    else:
                        fig_timeline = px.scatter(
                            timeline_src, x='mi.created_at', y='buyFromVendorName', color='reason',
                            size=[10] * len(timeline_src),
                            title=f"Return Timeline for Item {picked_item_no}",
                            hover_data={'reason': True, 'Line_Exposure': ':,.2f'}
                        )
                        fig_timeline.update_layout(
                            yaxis_title="Vendor", xaxis_title="Date Returned",
                            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                            font=dict(family="Inter, sans-serif")
                        )
                        st.plotly_chart(fig_timeline, use_container_width=True)

                        gaps = timeline_src['mi.created_at'].diff().dt.days.dropna()
                        if len(gaps) >= 2:
                            st.caption(
                                f"This item was returned {len(timeline_src)} times between "
                                f"{timeline_src['mi.created_at'].min().date()} and {timeline_src['mi.created_at'].max().date()}, "
                                f"averaging one return every **{gaps.mean():.1f} days** "
                                f"(most consistent gap: {gaps.mode().iloc[0]:.0f} days; range: {gaps.min():.0f}-{gaps.max():.0f} days)."
                            )

                    st.download_button(
                        "Download Full Recurring-Item Data (CSV)",
                        data=convert_df_to_csv(recurring_items[detail_cols]),
                        file_name=f"Recurring_Items_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv",
                        key="dl_recurring_items"
                    )

    # ------------------------------------------
    # TAB 4: REJECTION RATE vs. RECEIVING
    # ------------------------------------------
    # Business rule (per Ops): a vendor's delivery is only ACCEPTED into the
    # warehouse (and therefore appears in 'Receiving Report per items') if
    # the item's remaining shelf life is >= 70%. If it's below that, the
    # delivery is rejected right at the gate and NEVER enters inventory —
    # but it IS logged as a discrepancy record (e.g. NEAR_EXPIRY) in
    # 'pro_with_issues_linked_with_po_'.
    #
    # This means "Receiving Report" only captures ACCEPTED attempts, not
    # every time the vendor showed up. So the correct rejection rate is:
    #
    #   Rejection Rate = Rejected Attempts / (Rejected Attempts + Accepted Attempts)
    #
    # NOT flagged-quantity / received-quantity — that under-counts total
    # attempts, since rejected ones are invisible to the receiving report by
    # design. This version is always between 0% and 100%.
    # ------------------------------------------
    with sla_tab6:
        st.subheader("Rejection Rate — Accepted vs. Rejected at the Gate")
        st.write(
            "Every time a vendor brings an item, it's either **accepted** (appears in `Receiving Report per items`) "
            "or **rejected at the gate** — e.g. shelf life below the 70% threshold — which never enters inventory "
            "but is logged as a discrepancy in `pro_with_issues_linked_with_po_`. This tab computes the real "
            "rejection rate: rejected attempts ÷ (rejected + accepted) attempts."
        )

        if df_receiving.empty:
            st.info("The 'Receiving Report per items' sheet is not available, so a rejection rate cannot be computed.")
        elif df_issues.empty or 'item_no' not in df_issues.columns:
            st.info("The 'pro_with_issues_linked_with_po_' sheet is not available, so a rejection rate cannot be computed.")
        else:
            issue_src = df_issues.copy()
            issue_src['missing_quantity'] = pd.to_numeric(issue_src.get('missing_quantity', 0), errors='coerce').fillna(0.0)
            issue_src['cost_inc_vat'] = pd.to_numeric(issue_src.get('cost_inc_vat', 0), errors='coerce').fillna(0.0)
            issue_src['Line_Exposure'] = issue_src['cost_inc_vat'] * issue_src['missing_quantity']

            # Cached separately: this groupby scans all 500K+ receiving rows,
            # and Streamlit reruns this whole block on every widget interaction
            # (typing in the search box, changing Top N, etc.) — caching keeps
            # those interactions instant instead of re-scanning every time.
            @st.cache_data(show_spinner=False)
            def _aggregate_receiving(recv_df):
                return recv_df.groupby(['item_no', 'buyFromVendorNo']).agg(
                    Times_Accepted=('item_no', 'count'),
                    Qty_Accepted=('QTY', 'sum')
                ).reset_index()

            recv_agg = _aggregate_receiving(df_receiving)

            # ---- All 6 reasons are treated identically ----
            # Confirmed by Ops: 'Receiving Report per items' contains ONLY
            # items that the receiving team physically accepted and posted
            # into the system. ALL 6 reasons (including PRICE_ISSUE) are
            # "Receiving under adjustment" — never actually received, never
            # entered Receiving Report, and get a PRO the same way a rejected
            # item does. So there's no special-casing needed: every reason's
            # quantity is simply ADDED on top of Receiving Report to get the
            # true total invoiced quantity, and the rate is Rejected ÷ Total.
            all_reason_opts = sorted(issue_src['reason'].dropna().unique().tolist())
            st.markdown("##### All 6 reasons are included — none of them were ever physically received:")
            st.caption(
                "`Receiving Report per items` only contains items the warehouse team actually accepted and posted "
                "into the system. All 6 discrepancy reasons (NEAR_EXPIRY, MISSED_ITEM, QUALITY_ISSUE, NOT_LISTED, "
                "NOT_ORDERED, and PRICE_ISSUE) are 'Receiving under adjustment' — never received, never in this "
                "report, and handled as a rejected item with a PRO either way."
            )
            gate_reasons = st.multiselect(
                "Reasons included in this rate (deselect to exclude one):",
                options=all_reason_opts, default=all_reason_opts, key="rr_gate_reasons"
            )

            if not gate_reasons:
                st.warning("Select at least one reason above to compute a rejection rate.")
                st.stop()

            issue_src_gate = issue_src[issue_src['reason'].isin(gate_reasons)]
            issue_agg = issue_src_gate.groupby(['item_no', 'buyFromVendorNo']).agg(
                Times_Rejected=('item_no', 'count'), Qty_Rejected=('missing_quantity', 'sum'),
                Total_Exposure_SAR=('Line_Exposure', 'sum'),
                Dominant_Reason=('reason', lambda s: s.value_counts().idxmax() if s.notna().any() else pd.NA)
            ).reset_index()

            # Outer merge: keep item+vendor combos that only have accepted
            # deliveries too (0% rejection rate) so the overall baseline is
            # accurate, not just skewed toward problem items.
            rate_df = recv_agg.merge(issue_agg, on=['item_no', 'buyFromVendorNo'], how='outer')
            rate_df['buyFromVendorName'] = rate_df['buyFromVendorNo'].map(vendor_lookup)

            for c in ['Times_Accepted', 'Qty_Accepted', 'Times_Rejected', 'Qty_Rejected', 'Total_Exposure_SAR']:
                rate_df[c] = rate_df[c].fillna(0.0)

            # Total invoiced quantity = accepted (Receiving Report) + rejected
            # (never received, any of the 6 reasons above).
            rate_df['Total_Attempts'] = rate_df['Times_Accepted'] + rate_df['Times_Rejected']
            rate_df['Total_Qty_Attempted'] = rate_df['Qty_Accepted'] + rate_df['Qty_Rejected']

            rate_df = rate_df[rate_df['Total_Attempts'] > 0].copy()

            rate_df['Rejection_Rate_%'] = rate_df['Times_Rejected'] / rate_df['Total_Attempts'] * 100
            rate_df['Rejection_Rate_Qty_%'] = (
                rate_df['Qty_Rejected'] / rate_df['Total_Qty_Attempted'].replace(0, pd.NA) * 100
            ).fillna(0.0)

            total_combos = len(rate_df)
            never_accepted = int((rate_df['Times_Accepted'] == 0).sum())
            always_accepted = int((rate_df['Times_Rejected'] == 0).sum())
            baseline_rate = rate_df['Qty_Rejected'].sum() / rate_df['Total_Qty_Attempted'].sum() * 100 if rate_df['Total_Qty_Attempted'].sum() else 0.0

            kk1, kk2, kk3, kk4 = st.columns(4)
            kk1.metric("Item+Vendor Combos Analyzed", f"{total_combos:,}")
            kk2.metric("Baseline Rejection Rate", f"{baseline_rate:.2f}%", f"Based on all {len(gate_reasons)} selected reasons, by quantity")
            kk3.metric("Always Needs Credit Note (100%)", f"{never_accepted:,}", delta_color="inverse")
            kk4.metric("Never Needed One (0%)", f"{always_accepted:,}", delta_color="off")

            st.info(
                f"📊 **Baseline: {baseline_rate:.2f}%** of all invoiced quantity (across every item+vendor combo, "
                f"across all {len(gate_reasons)} selected reasons) ended up needing a credit note. Use this as the "
                f"'normal' benchmark — combos well above this are the real outliers, not just the ones with the "
                f"highest raw count."
            )

            if not gate_reasons == all_reason_opts:
                with st.expander("Reasons currently excluded from this rate"):
                    excluded_reasons = [r for r in all_reason_opts if r not in gate_reasons]
                    excl_summary = (
                        issue_src[issue_src['reason'].isin(excluded_reasons)]
                        .groupby('reason').agg(Line_Items=('reason', 'count'), Total_Exposure_SAR=('Line_Exposure', 'sum'))
                        .reset_index().sort_values('Total_Exposure_SAR', ascending=False)
                    )
                    st.dataframe(excl_summary, use_container_width=True, hide_index=True,
                                 column_config={"Total_Exposure_SAR": st.column_config.NumberColumn("Total Exposure (SAR)", format="%,.2f")})

            st.markdown("---")

            rf1, rf2 = st.columns(2)
            with rf1:
                min_attempts = st.number_input("Minimum total attempts (filters out noise):", min_value=1, value=3, step=1, key="rr_min_attempts")
            with rf2:
                rr_vendor_search = st.text_input("Search Vendor Name/Code:", key="rr_vendor_search").strip().lower()

            view_df = rate_df[rate_df['Total_Attempts'] >= min_attempts].copy()
            if rr_vendor_search:
                code_m = view_df['buyFromVendorNo'].str.lower().str.contains(rr_vendor_search, na=False)
                name_m = view_df['buyFromVendorName'].astype(str).str.lower().str.contains(rr_vendor_search, na=False)
                view_df = view_df[code_m | name_m]

            if view_df.empty:
                st.warning("No item+vendor combos match the current filters.")
            else:
                ch1, ch2 = st.columns([3, 1])
                with ch1:
                    st.markdown("#### Highest Rejection Rates (accepted vs. rejected attempts)")
                with ch2:
                    top_n_rr = st.number_input("Top N", min_value=3, max_value=50, value=10, step=1, key="topn_rejection_rate", label_visibility="collapsed")

                item_names_map = (
                    df_issues.dropna(subset=['item_no']).drop_duplicates('item_no').set_index('item_no')['mi.name'].to_dict()
                    if 'mi.name' in df_issues.columns else {}
                )
                view_df['Item_Name'] = view_df['item_no'].map(item_names_map).fillna('')
                chart_rr = view_df.sort_values(['Rejection_Rate_%', 'Total_Attempts'], ascending=[False, False]).head(top_n_rr).copy()
                chart_rr['Label'] = chart_rr['item_no'] + " (" + chart_rr['buyFromVendorName'].astype(str) + ")"

                fig_rr = px.bar(
                    chart_rr.sort_values('Rejection_Rate_%'), x='Rejection_Rate_%', y='Label', orientation='h',
                    color='Rejection_Rate_%', color_continuous_scale='Reds',
                    title=f"Top {top_n_rr} Highest Rejection Rates (min {min_attempts} attempts)",
                    custom_data=['Times_Rejected', 'Times_Accepted', 'Total_Attempts', 'Total_Exposure_SAR']
                )
                fig_rr.update_traces(
                    texttemplate='%{x:,.1f}%', textposition='outside',
                    hovertemplate='%{y}<br>Rejected %{customdata[0]:,.0f} of %{customdata[2]:,.0f} attempts '
                                  '(%{customdata[1]:,.0f} accepted)<br>Rate: %{x:,.1f}%'
                                  '<br>Exposure: %{customdata[3]:,.2f} SAR<extra></extra>'
                )
                fig_rr.update_layout(
                    yaxis_title="", xaxis_title="Rejection Rate (%)", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                    font=dict(family="Inter, sans-serif"), coloraxis_showscale=False,
                    xaxis=dict(ticksuffix='%', range=[0, 105])
                )
                st.plotly_chart(fig_rr, use_container_width=True)

                st.markdown("#### Full Rejection-Rate Detail")
                st.caption(
                    "`Rejection_Rate_%`: rejected attempts ÷ total attempts (accepted + rejected) — the primary, "
                    "event-based rate. `Rejection_Rate_Qty_%`: the same idea by quantity instead of event count, "
                    "shown for context (a rejected event with a huge quantity matters more than a small one)."
                )
                detail_rr_cols = ['item_no', 'Item_Name', 'buyFromVendorNo', 'buyFromVendorName', 'Dominant_Reason',
                                   'Times_Accepted', 'Times_Rejected', 'Total_Attempts', 'Rejection_Rate_%',
                                   'Qty_Accepted', 'Qty_Rejected', 'Rejection_Rate_Qty_%', 'Total_Exposure_SAR']
                detail_rr_cols = [c for c in detail_rr_cols if c in view_df.columns]
                st.dataframe(
                    view_df[detail_rr_cols].sort_values('Rejection_Rate_%', ascending=False),
                    use_container_width=True, hide_index=True,
                    column_config={
                        "buyFromVendorNo": "Vendor Code", "buyFromVendorName": "Vendor Name",
                        "Times_Accepted": st.column_config.NumberColumn("Times Accepted", format="%,.0f"),
                        "Times_Rejected": st.column_config.NumberColumn("Times Rejected", format="%,.0f"),
                        "Total_Attempts": st.column_config.NumberColumn("Total Attempts", format="%,.0f"),
                        "Rejection_Rate_%": st.column_config.NumberColumn("Rejection Rate", format="%.1f%%"),
                        "Qty_Accepted": st.column_config.NumberColumn("Qty Accepted", format="%,.0f"),
                        "Qty_Rejected": st.column_config.NumberColumn("Qty Rejected", format="%,.0f"),
                        "Rejection_Rate_Qty_%": st.column_config.NumberColumn("Rejection Rate (Qty)", format="%.1f%%"),
                        "Total_Exposure_SAR": st.column_config.NumberColumn("Exposure (SAR)", format="%,.2f"),
                    }
                )

                st.download_button(
                    "Download Rejection Rate Data (CSV)",
                    data=convert_df_to_csv(view_df[detail_rr_cols]),
                    file_name=f"Rejection_Rates_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    key="dl_rejection_rates"
                )

    # ------------------------------------------
    # TAB 5: LIVE VENDOR PICKUP RESPONSIVENESS
    # ------------------------------------------
    # This is a SEPARATE process from the PRO discrepancy workflow tracked in
    # Tabs 1-4 (confirmed by Ops: different document population, RTV vs
    # Purchase Order returns, zero overlap in reference numbers). It answers
    # one specific question: once we've told a vendor "come pick this up",
    # how long are they actually taking to show up?
    #
    # IMPORTANT LIMITATION (per Ops): once a return is Collected or gets its
    # credit note Posted, the date fields get OVERWRITTEN with the closing
    # date — the original "waiting since" date is lost. So a real historical
    # TAT cannot be computed from this sheet. This view is intentionally
    # LIVE-ONLY: it only looks at PROs still sitting in "Pending for
    # Collection" today, where 'aging_days' is still the genuine, unmodified
    # number of days that stock has been waiting for vendor pickup.
    # ------------------------------------------
    with sla_tab7:
        st.subheader("Live Vendor Pickup Responsiveness")
        st.write(
            "How long is stock actually sitting and waiting for **vendor pickup**, right now? This tracks the "
            "RTV (Return To Vendor) process from `Pending Returns` — a separate workflow from the PRO discrepancy "
            "process in the other tabs. Only 'Pending for Collection' records are used, because once a return is "
            "collected or credited, its original wait-start date is overwritten and can't be recovered — so this "
            "is a live snapshot, not a historical average."
        )

        if df_returns.empty or 'Status' not in df_returns.columns:
            st.info("The 'Pending Returns' sheet is not available, so pickup responsiveness cannot be computed.")
        else:
            live_pending = df_returns[df_returns['Status'] == 'Pending for Collection'].copy()

            if live_pending.empty:
                st.success("No returns are currently waiting for vendor pickup.")
            else:
                live_pending['aging_days'] = pd.to_numeric(live_pending.get('aging_days', 0), errors='coerce').fillna(0)
                amt_col = 'amountIncludingVAT' if 'amountIncludingVAT' in live_pending.columns else None
                if amt_col:
                    live_pending[amt_col] = pd.to_numeric(live_pending[amt_col], errors='coerce').fillna(0.0)

                def assign_pickup_tier(days):
                    if days <= 14:
                        return "Fast (<= 14 days)"
                    elif days <= 30:
                        return "Moderate (15-30 days)"
                    elif days <= 60:
                        return "Slow (31-60 days)"
                    else:
                        return "Critical (> 60 days)"

                live_pending['Pickup_Tier'] = live_pending['aging_days'].apply(assign_pickup_tier)

                total_pending = len(live_pending)
                total_value = live_pending[amt_col].sum() if amt_col else 0.0
                avg_wait = live_pending['aging_days'].mean()
                critical_count = int((live_pending['aging_days'] > 60).sum())
                oldest_wait = int(live_pending['aging_days'].max())

                p1, p2, p3, p4 = st.columns(4)
                p1.metric("Currently Awaiting Pickup", f"{total_pending:,}")
                p2.metric("Value Waiting for Pickup", f"{total_value:,.2f} SAR" if amt_col else "N/A")
                p3.metric("Average Wait So Far", f"{avg_wait:.1f} Days")
                p4.metric("Critical (> 60 Days)", f"{critical_count:,}", delta_color="inverse")

                st.caption(f"Oldest item still waiting for pickup: **{oldest_wait} days**.")
                st.markdown("---")

                st.markdown("#### Vendor Pickup Scorecard")
                st.write("Ranked by average wait time — this is a *live* ranking of who's slow to collect their returns right now, not a historical average.")

                vf1, vf2 = st.columns(2)
                with vf1:
                    min_items = st.number_input("Minimum items waiting (filters out one-off noise):", min_value=1, value=2, step=1, key="pickup_min_items")
                with vf2:
                    pickup_vendor_search = st.text_input("Search Vendor Name/Code:", key="pickup_vendor_search").strip().lower()

                vendor_agg = live_pending.groupby(['buyFromVendorNo', 'buyFromVendorName']).agg(
                    Items_Waiting=('buyFromVendorNo', 'count'),
                    Avg_Wait_Days=('aging_days', 'mean'),
                    Oldest_Wait_Days=('aging_days', 'max'),
                    Value_Waiting_SAR=(amt_col, 'sum') if amt_col else ('buyFromVendorNo', 'count')
                ).reset_index()
                vendor_agg = vendor_agg[vendor_agg['Items_Waiting'] >= min_items]
                vendor_agg['Pickup_Tier'] = vendor_agg['Avg_Wait_Days'].apply(assign_pickup_tier)

                if pickup_vendor_search:
                    code_m = vendor_agg['buyFromVendorNo'].str.lower().str.contains(pickup_vendor_search, na=False)
                    name_m = vendor_agg['buyFromVendorName'].astype(str).str.lower().str.contains(pickup_vendor_search, na=False)
                    vendor_agg = vendor_agg[code_m | name_m]

                vendor_agg = vendor_agg.sort_values('Avg_Wait_Days', ascending=False)

                if vendor_agg.empty:
                    st.info("No vendors match the current filters.")
                else:
                    def _highlight_pickup_tier(row):
                        tier = row.get('Pickup_Tier', '')
                        if 'Fast' in tier:
                            style = 'background-color: rgba(16, 185, 129, 0.15); color: #10b981; font-weight: 600'
                        elif 'Moderate' in tier:
                            style = 'background-color: rgba(14, 165, 233, 0.15); color: #0ea5e9; font-weight: 600'
                        elif 'Slow' in tier:
                            style = 'background-color: rgba(245, 158, 11, 0.15); color: #f59e0b; font-weight: 600'
                        else:
                            style = 'background-color: rgba(239, 68, 68, 0.15); color: #ef4444; font-weight: 600'
                        return [style if col == 'Pickup_Tier' else '' for col in row.index]

                    st.dataframe(
                        vendor_agg.style.apply(_highlight_pickup_tier, axis=1),
                        use_container_width=True, hide_index=True,
                        column_config={
                            "buyFromVendorNo": "Vendor Code", "buyFromVendorName": "Vendor Name",
                            "Items_Waiting": st.column_config.NumberColumn(format="%,d"),
                            "Avg_Wait_Days": st.column_config.NumberColumn("Avg Wait (Days)", format="%.1f"),
                            "Oldest_Wait_Days": st.column_config.NumberColumn("Oldest Wait (Days)", format="%.0f"),
                            "Value_Waiting_SAR": st.column_config.NumberColumn("Value Waiting (SAR)", format="%,.2f"),
                        }
                    )

                    critical_vendors = vendor_agg[vendor_agg['Pickup_Tier'] == 'Critical (> 60 days)']
                    if not critical_vendors.empty:
                        st.warning(
                            f"⚠️ **{len(critical_vendors):,} vendor(s)** are averaging over 60 days to collect their "
                            f"returns — worth an immediate pickup-escalation call, separate from any PRO/quality escalation."
                        )

                    st.download_button(
                        "Download Vendor Pickup Scorecard (CSV)",
                        data=convert_df_to_csv(vendor_agg),
                        file_name=f"Vendor_Pickup_Scorecard_{datetime.now().strftime('%Y%m%d')}.csv",
                        mime="text/csv",
                        key="dl_pickup_scorecard"
                    )

    # ------------------------------------------
    # TAB 6: VENDOR 360° RELIABILITY SCORECARD
    # ------------------------------------------
    # The PRO-discrepancy workflow (Tabs 1-4) and the RTV pickup workflow
    # (Tab 5) are two separate document populations with no shared reference
    # number — but they DO share one common key: the vendor code. This tab
    # combines every dimension we track, per vendor, into one view, so a
    # vendor that's fine on 3 metrics but critical on the 4th doesn't stay
    # hidden inside a single-purpose tab.
    # ------------------------------------------
    with sla_tab8:
        st.subheader("Vendor 360° Reliability Scorecard")
        st.write(
            "Every vendor-performance dimension tracked across this Hub, combined into one table: how many "
            "discrepancy PROs they cause, how fast they resolve them, their rejection rate at receiving, and how "
            "fast they pick up RTV returns. A vendor can look fine on three metrics and be critical on the fourth — "
            "this table is built so that doesn't stay hidden in a single-purpose tab."
        )

        # ---- A) PRO creation + resolution speed (from Tabs 1-2 data sources) ----
        pro_caused = pd.DataFrame(columns=['buyFromVendorNo', 'PROs_Caused'])
        if not df_issues.empty and 'no' in df_issues.columns and 'buyFromVendorNo' in df_issues.columns:
            pro_caused = df_issues.groupby('buyFromVendorNo')['no'].nunique().reset_index(name='PROs_Caused')

        resolution = pd.DataFrame(columns=['buyFromVendorNo', 'Closed_PROs', 'Avg_Resolution_TAT_Days'])
        if not closed_base.empty:
            resolution = closed_base.groupby('buyFromVendorNo').agg(
                Closed_PROs=('buyFromVendorNo', 'count'),
                Avg_Resolution_TAT_Days=('tat_days', 'mean')
            ).reset_index()

        def _resolution_tier(days):
            if pd.isna(days):
                return "No Closed History"
            elif days <= 7:
                return "Fast"
            elif days <= 14:
                return "Moderate"
            elif days <= 30:
                return "Slow"
            else:
                return "Critical"

        # ---- B) Rejection rate at receiving (from Tab 4's model) ----
        rejection = pd.DataFrame(columns=['buyFromVendorNo', 'Rejection_Rate_%'])
        if not df_receiving.empty and not df_issues.empty and 'item_no' in df_issues.columns:
            recv_by_vendor = df_receiving.groupby('buyFromVendorNo').agg(
                Times_Accepted=('buyFromVendorNo', 'count')
            ).reset_index()
            issue_by_vendor = df_issues.groupby('buyFromVendorNo').agg(
                Times_Rejected=('buyFromVendorNo', 'count')
            ).reset_index()
            rejection = recv_by_vendor.merge(issue_by_vendor, on='buyFromVendorNo', how='outer')
            rejection[['Times_Accepted', 'Times_Rejected']] = rejection[['Times_Accepted', 'Times_Rejected']].fillna(0.0)
            rejection['Total_Attempts'] = rejection['Times_Accepted'] + rejection['Times_Rejected']
            rejection = rejection[rejection['Total_Attempts'] > 0].copy()
            rejection['Rejection_Rate_%'] = rejection['Times_Rejected'] / rejection['Total_Attempts'] * 100
            rejection = rejection[['buyFromVendorNo', 'Rejection_Rate_%']]

        def _rejection_tier(rate):
            if pd.isna(rate):
                return "No Receiving History"
            elif rate <= 2:
                return "Low"
            elif rate <= 5:
                return "Moderate"
            elif rate <= 10:
                return "High"
            else:
                return "Critical"

        # ---- C) Live RTV pickup responsiveness (from Tab 5's model) ----
        pickup = pd.DataFrame(columns=['buyFromVendorNo', 'RTV_Items_Waiting', 'RTV_Avg_Wait_Days'])
        if not df_returns.empty and 'Status' in df_returns.columns:
            live_pending_360 = df_returns[df_returns['Status'] == 'Pending for Collection'].copy()
            if not live_pending_360.empty:
                live_pending_360['aging_days'] = pd.to_numeric(live_pending_360.get('aging_days', 0), errors='coerce').fillna(0)
                pickup = live_pending_360.groupby('buyFromVendorNo').agg(
                    RTV_Items_Waiting=('buyFromVendorNo', 'count'),
                    RTV_Avg_Wait_Days=('aging_days', 'mean')
                ).reset_index()

        def _pickup_tier(days):
            if pd.isna(days):
                return "Nothing Waiting"
            elif days <= 14:
                return "Fast"
            elif days <= 30:
                return "Moderate"
            elif days <= 60:
                return "Slow"
            else:
                return "Critical"

        # ---- Combine all four into one vendor-level table ----
        scorecard = pro_caused.merge(resolution, on='buyFromVendorNo', how='outer') \
                               .merge(rejection, on='buyFromVendorNo', how='outer') \
                               .merge(pickup, on='buyFromVendorNo', how='outer')

        if scorecard.empty:
            st.info("No vendor data available across any of the tracked dimensions.")
        else:
            scorecard['buyFromVendorName'] = scorecard['buyFromVendorNo'].map(vendor_lookup)
            scorecard['PROs_Caused'] = scorecard['PROs_Caused'].fillna(0).astype(int)
            scorecard['Closed_PROs'] = scorecard['Closed_PROs'].fillna(0).astype(int)
            scorecard['RTV_Items_Waiting'] = scorecard['RTV_Items_Waiting'].fillna(0).astype(int)

            scorecard['Resolution_Tier'] = scorecard['Avg_Resolution_TAT_Days'].apply(_resolution_tier)
            scorecard['Rejection_Tier'] = scorecard['Rejection_Rate_%'].apply(_rejection_tier)
            scorecard['Pickup_Tier'] = scorecard['RTV_Avg_Wait_Days'].apply(_pickup_tier)

            # Red Flags: how many of the 3 speed/quality dimensions are
            # Slow/Critical/High for this vendor — the higher this number,
            # the more this is a vendor-wide reliability problem rather than
            # a one-off issue in a single process.
            def _count_red_flags(row):
                flags = 0
                if row['Resolution_Tier'] in ('Slow', 'Critical'):
                    flags += 1
                if row['Rejection_Tier'] in ('High', 'Critical'):
                    flags += 1
                if row['Pickup_Tier'] in ('Slow', 'Critical'):
                    flags += 1
                return flags

            scorecard['Red_Flags'] = scorecard.apply(_count_red_flags, axis=1)

            f1, f2, f3 = st.columns(3)
            with f1:
                min_pros_360 = st.number_input("Minimum PROs caused (filters out one-off vendors):", min_value=0, value=3, step=1, key="v360_min_pros")
            with f2:
                min_flags_360 = st.number_input("Minimum Red Flags:", min_value=0, max_value=3, value=0, step=1, key="v360_min_flags")
            with f3:
                v360_search = st.text_input("Search Vendor Name/Code:", key="v360_search").strip().lower()

            view_360 = scorecard[(scorecard['PROs_Caused'] >= min_pros_360) & (scorecard['Red_Flags'] >= min_flags_360)].copy()
            if v360_search:
                code_m = view_360['buyFromVendorNo'].str.lower().str.contains(v360_search, na=False)
                name_m = view_360['buyFromVendorName'].astype(str).str.lower().str.contains(v360_search, na=False)
                view_360 = view_360[code_m | name_m]

            view_360 = view_360.sort_values(['Red_Flags', 'PROs_Caused'], ascending=[False, False])

            multi_flag_count = int((scorecard['Red_Flags'] >= 2).sum())
            st.info(
                f"📌 **{multi_flag_count:,} vendor(s)** are flagged as Slow/Critical/High on **2 or more** of the "
                f"three tracked dimensions at once — these are vendor-wide reliability problems, not an isolated "
                f"issue in one process, and are the top escalation priority overall."
            )

            if view_360.empty:
                st.warning("No vendors match the current filters.")
            else:
                def _highlight_flags(row):
                    flags = row.get('Red_Flags', 0)
                    if flags >= 2:
                        style = 'background-color: rgba(239, 68, 68, 0.15); color: #ef4444; font-weight: 700'
                    elif flags == 1:
                        style = 'background-color: rgba(245, 158, 11, 0.15); color: #f59e0b; font-weight: 600'
                    else:
                        style = 'background-color: rgba(16, 185, 129, 0.10); color: #10b981; font-weight: 600'
                    return [style if col == 'Red_Flags' else '' for col in row.index]

                display_cols = ['buyFromVendorNo', 'buyFromVendorName', 'PROs_Caused', 'Closed_PROs',
                                 'Avg_Resolution_TAT_Days', 'Resolution_Tier', 'Rejection_Rate_%', 'Rejection_Tier',
                                 'RTV_Items_Waiting', 'RTV_Avg_Wait_Days', 'Pickup_Tier', 'Red_Flags']
                display_cols = [c for c in display_cols if c in view_360.columns]

                st.dataframe(
                    view_360[display_cols].style.apply(_highlight_flags, axis=1),
                    use_container_width=True, hide_index=True,
                    column_config={
                        "buyFromVendorNo": "Vendor Code", "buyFromVendorName": "Vendor Name",
                        "Avg_Resolution_TAT_Days": st.column_config.NumberColumn("Resolution TAT (Days)", format="%.1f"),
                        "Rejection_Rate_%": st.column_config.NumberColumn("Rejection Rate", format="%.2f%%"),
                        "RTV_Avg_Wait_Days": st.column_config.NumberColumn("RTV Wait (Days)", format="%.1f"),
                        "Red_Flags": st.column_config.NumberColumn("Red Flags", format="%d / 3"),
                    }
                )

                st.download_button(
                    "Download Vendor 360° Scorecard (CSV)",
                    data=convert_df_to_csv(view_360[display_cols]),
                    file_name=f"Vendor_360_Scorecard_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    key="dl_v360_scorecard"
                )

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