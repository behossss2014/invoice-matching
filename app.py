import streamlit as st
import pandas as pd
import os
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go

# ==========================================
# 1. Page Configuration & Adaptive Executive Theme CSS
# ==========================================
st.set_page_config(page_title="Warehouse Intelligence Hub", layout="wide")

st.markdown("""
    <style>
    /* Global Container & Adaptive Executive Theme */
    [data-testid="stHeader"] { visibility: hidden; }
    
    .block-container {
        padding-top: 4.5rem !important;
        font-family: 'Inter', 'Segoe UI', -apple-system, sans-serif;
        max-width: 98%;
    }
    
    h1, h2, h3, h4 { 
        text-align: left; 
        direction: ltr; 
        font-weight: 700;
        letter-spacing: -0.02em;
        margin-bottom: 0.5rem;
    }

    /* Modern Navbar Theme (Adaptive) */
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
        justify-content: center;
        box-shadow: 0 2px 6px rgba(0, 0, 0, 0.15);
        border-bottom: 1px solid rgba(128, 128, 128, 0.2);
    }
    .ninja-logo {
        color: #00c9b1;
        font-size: 26px;
        font-weight: 800;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        letter-spacing: -0.5px;
    }

    /* Adaptive Sidebar Refinement */
    [data-testid="stSidebar"] {
        border-right: 1px solid rgba(128, 128, 128, 0.2);
    }

    /* Card & Alert Adaptive Styles (Light Mode Default) */
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
        font-size: 15px;
        font-weight: 700;
        margin-bottom: 10px;
        color: #00c9b1;
        border-bottom: 1px solid #e2e8f0;
        padding-bottom: 8px;
        text-transform: uppercase;
        letter-spacing: 0.03em;
    }

    /* Executive Compact Metric Banners (Light Mode Default) */
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
    </div>
    """, unsafe_allow_html=True)

EXCEL_FILE = "Returns Sheet.xlsx"

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
    if not os.path.exists(file_path): return None, None, None, None
    
    xls = pd.ExcelFile(file_path)
    df_scheduled = pd.read_excel(xls, sheet_name='Next Day Scheduled') if 'Next Day Scheduled' in xls.sheet_names else pd.DataFrame()
    df_returns = pd.read_excel(xls, sheet_name='Pending Returns') if 'Pending Returns' in xls.sheet_names else pd.DataFrame()
    df_pending_pros = pd.read_excel(xls, sheet_name='Pending PROs') if 'Pending PROs' in xls.sheet_names else pd.DataFrame()
    
    linked_sheet_name = next((name for name in ['Linked_With_PO', 'Linked With PO', 'Linked_with_PO'] if name in xls.sheet_names), None)
    df_linked = pd.read_excel(xls, sheet_name=linked_sheet_name) if linked_sheet_name else pd.DataFrame()
    
    def standardize_columns(df):
        if df.empty: return
        rename_dict = {}
        for col in df.columns:
            clean_col = str(col).lower().replace(' ', '').replace('_', '').replace('-', '').replace('.', '')
            if clean_col in ['supplierno', 'buyfromvendorno', 'vendorno', 'vendorcode', 'suppliercode']:
                rename_dict[col] = 'buyFromVendorNo'
            elif clean_col in ['suppliername', 'buyfromvendorname', 'buyfromvendornam', 'vendorname', 'name']:
                rename_dict[col] = 'buyFromVendorName'
            elif clean_col in ['no', 'documentno', 'prono', 'number']:
                rename_dict[col] = 'no' 
        df.rename(columns=rename_dict, inplace=True)

    standardize_columns(df_scheduled)
    standardize_columns(df_returns)
    standardize_columns(df_linked)
    standardize_columns(df_pending_pros)
            
    def clean_vendor_code(val):
        if pd.isna(val) or str(val).strip() == "": return ""
        val_str = str(val).strip().upper()
        if val_str.startswith('V'): return val_str
        if val_str.replace('.','',1).isdigit(): return f"V{int(float(val_str)):05d}"
        return val_str
        
    for df in [df_scheduled, df_returns, df_linked, df_pending_pros]:
        if not df.empty and 'buyFromVendorNo' in df.columns:
            df['buyFromVendorNo'] = df['buyFromVendorNo'].apply(clean_vendor_code)
            
    if not df_scheduled.empty and 'PO_Number' in df_scheduled.columns:
        df_scheduled['PO_Number'] = df_scheduled['PO_Number'].astype(str).str.strip().str.upper()
        
    for df in [df_returns, df_linked, df_pending_pros]:
        if not df.empty and 'no' in df.columns:
            df['no'] = df['no'].astype(str).str.strip().str.upper()
    
    return df_scheduled, df_returns, df_linked, df_pending_pros

def calculate_target_amounts(df):
    if df.empty: return 0.0, 0
    amt_col = [c for c in df.columns if 'amount' in c.lower() and 'vat' in c.lower()]
    if not amt_col:
        amt_col = [c for c in df.columns if 'amount' in c.lower()] 
    total_amt = pd.to_numeric(df[amt_col[0]], errors='coerce').sum() if amt_col else 0.0
    return total_amt, len(df)

df_scheduled, df_returns, df_linked, df_pending_pros = load_all_sheets_live(EXCEL_FILE)

if df_scheduled is None:
    st.error(f"System Error: Target file '{EXCEL_FILE}' not found in directory.")
    st.stop()

# ==========================================
# Vendor Dictionary Build
# ==========================================
all_codes = set()
for df in [df_scheduled, df_returns, df_linked, df_pending_pros]:
    if not df.empty and 'buyFromVendorNo' in df.columns:
        all_codes.update(df['buyFromVendorNo'].dropna().unique())

vendor_lookup = {}
for df in [df_scheduled, df_returns, df_linked, df_pending_pros]:
    if not df.empty and 'buyFromVendorNo' in df.columns and 'buyFromVendorName' in df.columns:
        temp_dict = df.drop_duplicates(subset=['buyFromVendorNo']).set_index('buyFromVendorNo')['buyFromVendorName'].to_dict()
        vendor_lookup.update(temp_dict)

vendor_options_list = sorted([f"{code} - {vendor_lookup.get(code, 'Unknown Vendor')}" for code in all_codes if code])
vendor_mapping = {opt.split(" - ")[0]: opt.split(" - ")[1] for opt in vendor_options_list}

# ==========================================
# 3. Navigation Menu & User Session Sidebar
# ==========================================
st.sidebar.markdown(f"<p style='font-size:12px; margin-top:5px;'>Logged in as:<br><b>{st.session_state.logged_in_user}</b></p>", unsafe_allow_html=True)

st.sidebar.markdown("<h3 style='margin-top: 15px; margin-bottom: 15px; font-size: 16px; font-weight:700;'>Navigation Menu</h3>", unsafe_allow_html=True)
page = st.sidebar.radio("Select Module:", [
    "Gate Operations", 
    "Today's Pre-Alerts", 
    "Data Analytics & Insights",
    "Executive Analytics",
    "Audit Trail & Logs"
])
st.sidebar.markdown("---")
st.sidebar.write("**System Status:** Active (Live Mode)")

# ==========================================
# PAGE 1: GATE OPERATIONS
# ==========================================
if page == "Gate Operations":
    st.markdown("<h2>Gate Operations Dashboard</h2>", unsafe_allow_html=True)
    st.write("Search scheduled or unscheduled POs, review liabilities, and record arrival with employee tracking.")
    
    search_type = st.radio("Search Criteria:", ["PO Number", "Vendor Name/Code"], horizontal=True, key="gate_search_type")
    
    def display_vendor_status(v_no, v_name, v_time_slot, active_po_num, has_appointment=True):
        status_text = "Scheduled" if has_appointment else "Ad-hoc / Unscheduled Arrival"
        border_color = "#00C9B1" if has_appointment else "#f59e0b"
        
        st.markdown(f"""
        <div class="alert-box" style="border-left: 4px solid {border_color};">
            <div class="card-title">Vendor & Shipment Information</div>
            <p style="margin:5px 0;"><b>Name:</b> {v_name}</p>
            <p style="margin:5px 0;"><b>Code:</b> {v_no}</p>
            <p style="margin:5px 0;"><b>PO Number:</b> {active_po_num}</p>
            <p style="margin:5px 0;"><b>Type / Time Slot:</b> {v_time_slot} ({status_text})</p>
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
            
            amt_coll, _ = calculate_target_amounts(match_collected)
            amt_link, _ = calculate_target_amounts(match_linked)
            total_cn_amount = amt_coll + amt_link
            
            if total_cn_amount > 0:
                combined_cn = pd.concat([match_collected, match_linked], ignore_index=True) if not match_linked.empty else match_collected.copy()
                cn_date_col = 'documentDate' if 'documentDate' in combined_cn.columns else ('postingDate' if 'postingDate' in combined_cn.columns else None)
                
                if cn_date_col:
                    combined_cn[cn_date_col] = pd.to_datetime(combined_cn[cn_date_col], errors='coerce')
                    combined_cn['age'] = (pd.to_datetime('today') - combined_cn[cn_date_col]).dt.days
                    max_cn_age = combined_cn['age'].max()
                else:
                    max_cn_age = 0
                
                if max_cn_age > 20:
                    st.error(f"Critical: Unclosed CN exceeded 20 days (Max Age: {int(max_cn_age)} days)")
                elif max_cn_age >= 10:
                    st.warning(f"Warning: Unclosed CN aging (Max Age: {int(max_cn_age)} days)")
                else:
                    st.success(f"Status Normal (Max Age: {int(max_cn_age)} days)")

                st.metric(label="Total Combined CN Amount (Inc. VAT)", value=f"{total_cn_amount:,.2f} SAR")
            else:
                st.success("Clear: No outstanding Credit Notes found.")

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
                st.warning("PO Number not found in today's scheduled roster. Use Ad-hoc entry below to process un-scheduled arrival.")
                
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
        selected_vendors = st.multiselect("Search / Select Vendor(s):", options=vendor_options_list, placeholder="Select one or multiple vendors to inspect...")
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
                        "Processed By (Email)": u_val,
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
            
            amt_coll, count_coll = calculate_target_amounts(v_collected)
            amt_link, count_link = calculate_target_amounts(v_linked)
            
            total_cn_amt = amt_coll + amt_link
            total_cn_count = count_coll + count_link
            
            for sub_df in [v_collected, v_linked]:
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

            combined_cn = pd.concat([v_collected, v_linked], ignore_index=True) if not (v_collected.empty and v_linked.empty) else (v_collected.copy() if not v_collected.empty else v_linked.copy())
            
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
        m3.metric("Vendors with CNs", f"{has_cn_count}", f"{total_arriving_cn_amt:,.2f} SAR")
        m4.metric("Clear Pass Vendors", f"{clean_count}")
        st.write(f"CN Aging Summary - Current Month: {current_month_cn_total:,.2f} | Last Month: {last_month_cn_total:,.2f} | Older: {older_cn_total:,.2f}")
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
    st.write("---")
    
    loc_options = set()
    bin_options = set()
    for df_temp in [df_returns, df_linked, df_pending_pros]:
        if df_temp is not None and not df_temp.empty:
            if 'locationCode' in df_temp.columns:
                loc_options.update(df_temp['locationCode'].dropna().astype(str).unique())
            if 'binCode' in df_temp.columns:
                bin_options.update(df_temp['binCode'].dropna().astype(str).unique())
    
    loc_options_list = sorted([str(loc) for loc in loc_options if str(loc).strip() and str(loc).lower() != 'nan'])
    bin_options_list = sorted([str(b) for b in bin_options if str(b).strip() and str(b).lower() != 'nan'])

    with st.sidebar.expander("Filter Analytics", expanded=True):
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

    if selected_locations:
        if not filtered_df.empty and 'locationCode' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['locationCode'].astype(str).isin(selected_locations)]
        if not filtered_linked.empty and 'locationCode' in filtered_linked.columns:
            filtered_linked = filtered_linked[filtered_linked['locationCode'].astype(str).isin(selected_locations)]
        if not filtered_pros.empty and 'locationCode' in filtered_pros.columns:
            filtered_pros = filtered_pros[filtered_pros['locationCode'].astype(str).isin(selected_locations)]

    if selected_bins:
        if not filtered_df.empty and 'binCode' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['binCode'].astype(str).isin(selected_bins)]
        if not filtered_linked.empty and 'binCode' in filtered_linked.columns:
            filtered_linked = filtered_linked[filtered_linked['binCode'].astype(str).isin(selected_bins)]
        if not filtered_pros.empty and 'binCode' in filtered_pros.columns:
            filtered_pros = filtered_pros[filtered_pros['binCode'].astype(str).isin(selected_bins)]

    if len(date_range) == 2:
        start_date, end_date = date_range
        
        if not filtered_df.empty:
            if 'postingDate' in filtered_df.columns: filtered_df['postingDate'] = pd.to_datetime(filtered_df['postingDate'], errors='coerce')
            if 'documentDate' in filtered_df.columns: filtered_df['documentDate'] = pd.to_datetime(filtered_df['documentDate'], errors='coerce')
            
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

    returns_df = filtered_df[filtered_df['Status'] == 'Pending for Collection'].copy()
    collected_df = filtered_df[filtered_df['Status'] == 'Collected'].copy()
    
    total_ret_amt, count_ret = calculate_target_amounts(returns_df)
    total_coll_amt, count_coll = calculate_target_amounts(collected_df)
    total_link_amt, count_link = calculate_target_amounts(filtered_linked)
    
    total_cn_exposure = total_coll_amt + total_link_amt

    st.subheader(f"Results scope: {display_title}")
    c1, c2 = st.columns(2)
    c1.metric("Pending Returns Value", f"{total_ret_amt:,.2f} SAR")
    c1.caption(f"Document Count: {count_ret}")
    c2.metric("Total CN Exposure", f"{total_cn_exposure:,.2f} SAR")
    c2.caption(f"Collected: {total_coll_amt:,.2f} ({count_coll}) | Linked PO: {total_link_amt:,.2f} ({count_link})")

    global_linked_nos = df_linked['no'].unique() if 'no' in df_linked.columns else []
    global_returns_nos = df_returns['no'].unique() if 'no' in df_returns.columns else []
    
    if not filtered_pros.empty and 'no' in filtered_pros.columns:
        df_orphans = filtered_pros[
            (~filtered_pros['no'].isin(global_linked_nos)) & 
            (~filtered_pros['no'].isin(global_returns_nos))
        ]
    else:
        df_orphans = pd.DataFrame()

    st.markdown("### Detailed Record Analysis")

    tab1, tab2, tab3, tab4 = st.tabs(["Pending Returns", "Collected CNs", "Linked PO CNs", "Orphaned PROs"])
    
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

            if 'amount' not in df_t1.columns: df_t1['amount'] = 0.0
            if 'amountIncludingVAT' not in df_t1.columns:
                vat_cols = [c for c in df_t1.columns if 'vat' in str(c).lower()]
                df_t1['amountIncludingVAT'] = df_t1[vat_cols[0]] if vat_cols else 0.0

            if 'binCode' not in df_t1.columns: df_t1['binCode'] = "N/A"
            if 'locationCode' not in df_t1.columns: df_t1['locationCode'] = "N/A"
            if 'Status' not in df_t1.columns: df_t1['Status'] = 'Pending for Collection'

            cols_t1 = ['buyFromVendorNo', 'VendorName', 'documentDate', 'binCode', 'locationCode', 'amount', 'amountIncludingVAT', 'aging_days', 'no', 'days_since_last_receiving', 'Status']
            view_df1 = df_t1.reindex(columns=cols_t1)
            
            st.dataframe(
                view_df1, 
                use_container_width=True, hide_index=True,
                column_config={"amount": st.column_config.NumberColumn(format="%,.2f"), "amountIncludingVAT": st.column_config.NumberColumn(format="%,.2f"), "documentDate": st.column_config.DateColumn(format="YYYY-MM-DD")}
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
            if rec_cols: df_t2['days_since_last_receiving'] = df_t2[rec_cols[0]]
            elif 'days_since_last_receiving' not in df_t2.columns: df_t2['days_since_last_receiving'] = None

            if 'amount' not in df_t2.columns: df_t2['amount'] = 0.0
            if 'amountIncludingVAT' not in df_t2.columns:
                vat_cols = [c for c in df_t2.columns if 'vat' in str(c).lower()]
                df_t2['amountIncludingVAT'] = df_t2[vat_cols[0]] if vat_cols else 0.0

            if 'Status' not in df_t2.columns: df_t2['Status'] = 'Collected'

            cols_t2 = ['buyFromVendorNo', 'VendorName', 'amount', 'amountIncludingVAT', 'postingDate', 'aging_days', 'no', 'days_since_last_receiving', 'Status']
            view_df2 = df_t2.reindex(columns=cols_t2)

            st.dataframe(
                view_df2, 
                use_container_width=True, hide_index=True,
                column_config={"amount": st.column_config.NumberColumn(format="%,.2f"), "amountIncludingVAT": st.column_config.NumberColumn(format="%,.2f"), "postingDate": st.column_config.DateColumn(format="YYYY-MM-DD")}
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

            if 'amount' not in df_t3.columns: df_t3['amount'] = 0.0
            vat_cols = [c for c in df_t3.columns if 'vat' in str(c).lower()]
            if 'AmountWithVAT' in df_t3.columns: pass
            elif 'amountIncludingVAT' in df_t3.columns: df_t3['AmountWithVAT'] = df_t3['amountIncludingVAT']
            elif vat_cols: df_t3['AmountWithVAT'] = df_t3[vat_cols[0]]
            else: df_t3['AmountWithVAT'] = 0.0

            if 'locationCode' not in df_t3.columns: df_t3['locationCode'] = "N/A"
            if 'Status' not in df_t3.columns: df_t3['Status'] = 'Linked with PO'

            cols_t3 = ['buyFromVendorNo', 'VendorName', 'documentDate', 'no', 'locationCode', 'amount', 'AmountWithVAT', 'Status', 'Aging']
            view_df3 = df_t3.reindex(columns=cols_t3)

            st.dataframe(
                view_df3, 
                use_container_width=True, hide_index=True,
                column_config={"amount": st.column_config.NumberColumn(format="%,.2f"), "AmountWithVAT": st.column_config.NumberColumn(format="%,.2f"), "documentDate": st.column_config.DateColumn(format="YYYY-MM-DD")}
            )
            st.download_button("Download Linked PO CNs (CSV)", data=convert_df_to_csv(view_df3), file_name="Linked_PO_CNs.csv", mime="text/csv", key="dl_t3")
        else: 
            st.write("No Linked PO discrepancies found.")
        
    with tab4:
        if not df_orphans.empty:
            orphan_amt, orphan_count = calculate_target_amounts(df_orphans)
            st.write(f"Alert: {orphan_count} Orphaned PROs Detected | Total Value: {orphan_amt:,.2f} SAR")
            
            df_t4 = df_orphans.copy()
            df_t4['VendorName'] = df_t4['buyFromVendorNo'].map(vendor_mapping).fillna("Unknown Vendor")
            cols_t4 = ['buyFromVendorNo', 'VendorName'] + [c for c in df_t4.columns if c not in ['buyFromVendorNo', 'VendorName']]
            view_df4 = df_t4[cols_t4]
            
            st.dataframe(
                view_df4, 
                use_container_width=True, hide_index=True,
                column_config={"amount": st.column_config.NumberColumn(format="%,.2f"), "amountIncludingVAT": st.column_config.NumberColumn(format="%,.2f")}
            )
            st.download_button("Download Orphaned PROs (CSV)", data=convert_df_to_csv(view_df4), file_name="Orphaned_PROs.csv", mime="text/csv", key="dl_t4")
        else:
            st.write("System Status Normal: No Orphaned PROs detected.")

# ==========================================
# PAGE 4: EXECUTIVE ANALYTICS
# ==========================================
elif page == "Executive Analytics":
    st.markdown("<h2>Executive Analytics Dashboard</h2>", unsafe_allow_html=True)
    st.markdown("---")
    
    # Ensure standard amount including VAT calculation across base dataset
    exec_df = df_returns.copy()
    if not exec_df.empty:
        vat_cols = [c for c in exec_df.columns if 'amount' in str(c).lower() and 'vat' in str(c).lower()]
        if not vat_cols:
            vat_cols = [c for c in exec_df.columns if 'amount' in str(c).lower()]
        target_val_col = vat_cols[0] if vat_cols else 'amount'
        exec_df['amountIncludingVAT'] = pd.to_numeric(exec_df[target_val_col], errors='coerce').fillna(0.0)

    st.sidebar.header("Configuration Parameters")
    
    all_dates = pd.Series(dtype='datetime64[ns]')
    if 'documentDate' in exec_df.columns: all_dates = pd.concat([all_dates, pd.to_datetime(exec_df['documentDate'], errors='coerce')])
    if 'postingDate' in exec_df.columns: all_dates = pd.concat([all_dates, pd.to_datetime(exec_df['postingDate'], errors='coerce')])
        
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
    top_cn_n = st.sidebar.number_input("Top Variables Filter (CN):", min_value=1, max_value=100, value=10)
    top_ret_n = st.sidebar.number_input("Top Variables Filter (Returns):", min_value=1, max_value=100, value=10)
    top_aging_n = st.sidebar.number_input("Aging Summary Display Limit:", min_value=1, max_value=200, value=50)

    def get_vendor_label(code):
        return f"{code} - {vendor_mapping.get(code, 'Unknown')}"

    df_exec_filtered = exec_df.copy()
    df_linked_exec = df_linked.copy() if not df_linked.empty else pd.DataFrame()
    
    if isinstance(date_range_exec, tuple) and len(date_range_exec) == 2:
        start_date, end_date = date_range_exec
        if 'documentDate' in df_exec_filtered.columns: df_exec_filtered['documentDate'] = pd.to_datetime(df_exec_filtered['documentDate'], errors='coerce')
        if 'postingDate' in df_exec_filtered.columns: df_exec_filtered['postingDate'] = pd.to_datetime(df_exec_filtered['postingDate'], errors='coerce')
            
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

        st.write(f"Analytics Scope: Data spanning **{start_date}** to **{end_date}**.")

    st.subheader(f"Top {top_cn_n} Vendors: Outstanding CN Liability")
    cn_collected_df = df_exec_filtered[df_exec_filtered['Status'] == 'Collected']
    cn_data = cn_collected_df.groupby('buyFromVendorNo')['amountIncludingVAT'].sum().nlargest(top_cn_n).reset_index()
    
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
        cn_total_system_val = exec_df[exec_df['Status'] == 'Collected']['amountIncludingVAT'].sum() if not exec_df.empty else 0.0
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

    # ==========================================
    # MATRIX 1: Pending Returns
    # ==========================================
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
        if days_col: agg_dict[days_col] = 'min'

        aging_summary = pending_returns.groupby('buyFromVendorNo').agg(agg_dict).reset_index()
        
        rename_map = {'buyFromVendorNo': 'Vendor Code', 'buyFromVendorName': 'Vendor Name', target_amt_col: 'Total Amount', 'aging_days': 'Max_Age_Days', 'no': 'PRO Count'}
        if days_col: rename_map[days_col] = 'Last_Receiving_Raw'
        aging_summary = aging_summary.rename(columns=rename_map)
        
        aging_summary['Vendor Name'] = aging_summary['Vendor Code'].map(vendor_lookup).fillna(aging_summary['Vendor Name'])
        aging_summary['Percentage'] = (aging_summary['Total Amount'] / grand_total_pending * 100) if grand_total_pending > 0 else 0.0
        aging_summary['Is_Active'] = aging_summary['Vendor Code'].isin(active_vendors)
        
        def get_status_alert(row):
            if row['Max_Age_Days'] <= 90: return "Normal"
            elif row['Is_Active']: return "Active Mitigation"
            else: return "Requires Escalation"

        aging_summary['Commitment Status'] = aging_summary.apply(get_status_alert, axis=1)
        
        if 'Last_Receiving_Raw' in aging_summary.columns:
            def format_last_receiving(val):
                if pd.isna(val) or str(val).strip() == "" or str(val).lower() == "nan": return "N/A"
                try:
                    num = float(val)
                    return "0" if num == 0 else str(int(num))
                except: return "N/A"
            aging_summary['Last Receiving (Days)'] = aging_summary['Last_Receiving_Raw'].apply(format_last_receiving)
        else:
            aging_summary['Last Receiving (Days)'] = "N/A"

        sort_c1, sort_c2 = st.columns(2)
        with sort_c1:
            sort_choice = st.selectbox("Data Sort Order:", ["Total Amount (High to Low)", "Oldest Return Days (High to Low)", "PRO Count (High to Low)", "Vendor Name (A-Z)", "Commitment Status"], key="exec_sort_option")
        with sort_c2:
            display_limit = st.number_input("Records Limit:", min_value=5, max_value=200, value=top_aging_n, key="exec_limit_option")

        if "Total Amount" in sort_choice: aging_summary = aging_summary.sort_values(by='Total Amount', ascending=False)
        elif "Oldest Return Days" in sort_choice: aging_summary = aging_summary.sort_values(by='Max_Age_Days', ascending=False)
        elif "PRO Count" in sort_choice: aging_summary = aging_summary.sort_values(by='PRO Count', ascending=False)
        elif "Vendor Name" in sort_choice: aging_summary = aging_summary.sort_values(by='Vendor Name', ascending=True)
        elif "Commitment Status" in sort_choice: aging_summary = aging_summary.sort_values(by=['Commitment Status', 'Total Amount'], ascending=[True, False])

        display_aging_table = aging_summary.head(display_limit).copy()
        table_view = display_aging_table[['Vendor Code', 'Vendor Name', 'PRO Count', 'Total Amount', 'Percentage', 'Last Receiving (Days)', 'Max_Age_Days', 'Commitment Status']].copy()
        
        def highlight_aging_cells(row):
            age = row['Max_Age_Days']
            if age > 45: return ['background-color: rgba(239, 68, 68, 0.2); color: #ef4444; font-weight: 600'] * len(row)
            elif age > 20: return ['background-color: rgba(245, 158, 11, 0.2); color: #f59e0b; font-weight: 600'] * len(row)
            else: return ['background-color: rgba(16, 185, 129, 0.2); color: #10b981; font-weight: 600'] * len(row)

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
        
        # Interactive Expandable Vendor Breakdown
        with st.expander("Click to view Granular Vendor Document Breakdown (Pending Returns)", expanded=False):
            p_vendor_list = display_aging_table['Vendor Code'].unique().tolist()
            p_vendor_map = {r['Vendor Code']: f"{r['Vendor Code']} - {r['Vendor Name']} | SAR {r['Total Amount']:,.2f} ({r['PRO Count']} PROs)" for _, r in display_aging_table.iterrows()}
            
            selected_p_vcode = st.selectbox(
                "Select Vendor to Inspect Documents:", 
                options=p_vendor_list, 
                format_func=lambda x: p_vendor_map.get(x, x),
                key="sb_pending_vendor"
            )
            if selected_p_vcode:
                vendor_pros_df = pending_returns[pending_returns['buyFromVendorNo'] == selected_p_vcode]
                st.dataframe(
                    vendor_pros_df, 
                    use_container_width=True, hide_index=True,
                    column_config={"amount": st.column_config.NumberColumn(format="%,.2f"), "amountIncludingVAT": st.column_config.NumberColumn(format="%,.2f")}
                )
    else:
        st.write("No aging data applicable for the specified parameters.")

    st.markdown("---")

    # ==========================================
    # MATRIX 2: Collected Credit Notes
    # ==========================================
    st.subheader("Advanced Aging Summary Matrix - Collected Credit Notes")

    cn_collected_df = df_exec_filtered[df_exec_filtered['Status'] == 'Collected'].copy()

    if not cn_collected_df.empty:
        if 'aging_days' not in cn_collected_df.columns:
            date_col = 'postingDate' if 'postingDate' in cn_collected_df.columns else 'documentDate'
            if date_col in cn_collected_df.columns:
                cn_collected_df[date_col] = pd.to_datetime(cn_collected_df[date_col], errors='coerce')
                cn_collected_df['aging_days'] = (pd.to_datetime('today') - cn_collected_df[date_col]).dt.days
            else:
                cn_collected_df['aging_days'] = 0

        amt_col_cn = [c for c in cn_collected_df.columns if 'amount' in c.lower() and 'vat' in c.lower()]
        target_amt_cn = amt_col_cn[0] if amt_col_cn else ('amountIncludingVAT' if 'amountIncludingVAT' in cn_collected_df.columns else 'amount')
        cn_collected_df[target_amt_cn] = pd.to_numeric(cn_collected_df[target_amt_cn], errors='coerce').fillna(0.0)

        grand_total_cn = cn_collected_df[target_amt_cn].sum()
        days_col_cn = next((col for col in cn_collected_df.columns if str(col).lower().replace(' ', '').replace('_', '') in ['dayssincelastreceiving', 'sincelastreceiving', 'lastreceivingdays', 'dayssincereceiving']), None)

        agg_dict_cn = {
            'buyFromVendorName': lambda x: x.iloc[0] if not x.empty else 'Unknown Vendor',
            target_amt_cn: 'sum',
            'aging_days': 'max',
            'no': 'count'
        }
        if days_col_cn: agg_dict_cn[days_col_cn] = 'min'

        aging_cn = cn_collected_df.groupby('buyFromVendorNo').agg(agg_dict_cn).reset_index()
        rename_cn = {'buyFromVendorNo': 'Vendor Code', 'buyFromVendorName': 'Vendor Name', target_amt_cn: 'Total Amount', 'aging_days': 'Max_Age_Days', 'no': 'CN Count'}
        if days_col_cn: rename_cn[days_col_cn] = 'Last_Receiving_Raw'
        aging_cn = aging_cn.rename(columns=rename_cn)

        aging_cn['Vendor Name'] = aging_cn['Vendor Code'].map(vendor_lookup).fillna(aging_cn['Vendor Name'])
        aging_cn['Percentage'] = (aging_cn['Total Amount'] / grand_total_cn * 100) if grand_total_cn > 0 else 0.0

        def get_cn_status(row):
            if row['Max_Age_Days'] <= 10: return "Normal"
            elif row['Max_Age_Days'] <= 20: return "Pending Settlement"
            else: return "Requires Escalation"

        aging_cn['Commitment Status'] = aging_cn.apply(get_cn_status, axis=1)

        if 'Last_Receiving_Raw' in aging_cn.columns:
            aging_cn['Last Receiving (Days)'] = aging_cn['Last_Receiving_Raw'].apply(lambda v: str(int(float(v))) if pd.notna(v) and str(v).strip() != "" else "N/A")
        else:
            aging_cn['Last Receiving (Days)'] = "N/A"

        c1_cn, c2_cn = st.columns(2)
        with c1_cn:
            sort_cn = st.selectbox("Data Sort Order (Collected CNs):", ["Total Amount (High to Low)", "Oldest Days (High to Low)", "CN Count (High to Low)", "Vendor Name (A-Z)"], key="sort_cn_coll")
        with c2_cn:
            limit_cn = st.number_input("Records Limit (Collected CNs):", min_value=5, max_value=200, value=top_aging_n, key="limit_cn_coll")

        if "Total Amount" in sort_cn: aging_cn = aging_cn.sort_values(by='Total Amount', ascending=False)
        elif "Oldest Days" in sort_cn: aging_cn = aging_cn.sort_values(by='Max_Age_Days', ascending=False)
        elif "CN Count" in sort_cn: aging_cn = aging_cn.sort_values(by='CN Count', ascending=False)
        elif "Vendor Name" in sort_cn: aging_cn = aging_cn.sort_values(by='Vendor Name', ascending=True)

        display_cn_table = aging_cn.head(limit_cn).copy()
        view_cn_table = display_cn_table[['Vendor Code', 'Vendor Name', 'CN Count', 'Total Amount', 'Percentage', 'Last Receiving (Days)', 'Max_Age_Days', 'Commitment Status']].copy()

        def highlight_cn(row):
            age = row['Max_Age_Days']
            if age > 20: return ['background-color: rgba(239, 68, 68, 0.2); color: #ef4444; font-weight: 600'] * len(row)
            elif age > 10: return ['background-color: rgba(245, 158, 11, 0.2); color: #f59e0b; font-weight: 600'] * len(row)
            else: return ['background-color: rgba(16, 185, 129, 0.2); color: #10b981; font-weight: 600'] * len(row)

        st.dataframe(
            view_cn_table.style.apply(highlight_cn, axis=1),
            use_container_width=True, hide_index=True,
            column_config={
                "Total Amount": st.column_config.NumberColumn("Total Amount (SAR)", format="%,.2f SAR"),
                "Percentage": st.column_config.NumberColumn("Liability Share", format="%.2f%%"),
                "Max_Age_Days": st.column_config.NumberColumn("Peak Age (Days)"),
                "CN Count": st.column_config.NumberColumn("Volume (CNs)")
            }
        )

        st.download_button(
            label="Download Collected CNs Aging Matrix (CSV)",
            data=convert_df_to_csv(view_cn_table),
            file_name=f"Collected_CNs_Aging_Summary_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )

        # Interactive Expandable Vendor Breakdown
        with st.expander("Click to view Granular Vendor Document Breakdown (Collected CNs)", expanded=False):
            cn_vendor_list = display_cn_table['Vendor Code'].unique().tolist()
            cn_vendor_map = {r['Vendor Code']: f"{r['Vendor Code']} - {r['Vendor Name']} | SAR {r['Total Amount']:,.2f} ({r['CN Count']} CNs)" for _, r in display_cn_table.iterrows()}
            
            selected_cn_vcode = st.selectbox(
                "Select Vendor to Inspect Documents:", 
                options=cn_vendor_list, 
                format_func=lambda x: cn_vendor_map.get(x, x),
                key="sb_cn_vendor"
            )
            if selected_cn_vcode:
                v_df = cn_collected_df[cn_collected_df['buyFromVendorNo'] == selected_cn_vcode]
                st.dataframe(v_df, use_container_width=True, hide_index=True)
    else:
        st.info("No collected Credit Notes data available for the selected range.")

    st.markdown("---")

    # ==========================================
    # MATRIX 3: Linked PO Credit Notes
    # ==========================================
    st.subheader("Advanced Aging Summary Matrix - Linked PO Credit Notes")

    linked_cn_df = df_linked_exec.copy() if not df_linked_exec.empty else pd.DataFrame()

    if not linked_cn_df.empty:
        if 'documentDate' in linked_cn_df.columns:
            linked_cn_df['documentDate'] = pd.to_datetime(linked_cn_df['documentDate'], errors='coerce')
            linked_cn_df['aging_days'] = (pd.to_datetime('today') - linked_cn_df['documentDate']).dt.days
        else:
            linked_cn_df['aging_days'] = 0

        amt_col_link = [c for c in linked_cn_df.columns if 'amount' in c.lower() and 'vat' in c.lower()]
        target_amt_link = amt_col_link[0] if amt_col_link else ('AmountWithVAT' if 'AmountWithVAT' in linked_cn_df.columns else 'amount')
        linked_cn_df[target_amt_link] = pd.to_numeric(linked_cn_df[target_amt_link], errors='coerce').fillna(0.0)

        grand_total_link = linked_cn_df[target_amt_link].sum()

        agg_dict_link = {
            'buyFromVendorName': lambda x: x.iloc[0] if not x.empty else 'Unknown Vendor',
            target_amt_link: 'sum',
            'aging_days': 'max',
            'no': 'count'
        }

        aging_link = linked_cn_df.groupby('buyFromVendorNo').agg(agg_dict_link).reset_index()
        rename_link = {'buyFromVendorNo': 'Vendor Code', 'buyFromVendorName': 'Vendor Name', target_amt_link: 'Total Amount', 'aging_days': 'Max_Age_Days', 'no': 'Linked CN Count'}
        aging_link = aging_link.rename(columns=rename_link)

        aging_link['Vendor Name'] = aging_link['Vendor Code'].map(vendor_lookup).fillna(aging_link['Vendor Name'])
        aging_link['Percentage'] = (aging_link['Total Amount'] / grand_total_link * 100) if grand_total_link > 0 else 0.0

        def get_link_status(row):
            if row['Max_Age_Days'] <= 15: return "Normal"
            elif row['Max_Age_Days'] <= 30: return "Pending PO Matching"
            else: return "Action Required"

        aging_link['Commitment Status'] = aging_link.apply(get_link_status, axis=1)

        c1_lk, c2_lk = st.columns(2)
        with c1_lk:
            sort_lk = st.selectbox("Data Sort Order (Linked PO CNs):", ["Total Amount (High to Low)", "Oldest Days (High to Low)", "Linked CN Count (High to Low)", "Vendor Name (A-Z)"], key="sort_cn_linked")
        with c2_lk:
            limit_lk = st.number_input("Records Limit (Linked PO CNs):", min_value=5, max_value=200, value=top_aging_n, key="limit_cn_linked")

        if "Total Amount" in sort_lk: aging_link = aging_link.sort_values(by='Total Amount', ascending=False)
        elif "Oldest Days" in sort_lk: aging_link = aging_link.sort_values(by='Max_Age_Days', ascending=False)
        elif "Linked CN Count" in sort_lk: aging_link = aging_link.sort_values(by='Linked CN Count', ascending=False)
        elif "Vendor Name" in sort_lk: aging_link = aging_link.sort_values(by='Vendor Name', ascending=True)

        display_link_table = aging_link.head(limit_lk).copy()
        view_link_table = display_link_table[['Vendor Code', 'Vendor Name', 'Linked CN Count', 'Total Amount', 'Percentage', 'Max_Age_Days', 'Commitment Status']].copy()

        def highlight_link(row):
            age = row['Max_Age_Days']
            if age > 30: return ['background-color: rgba(239, 68, 68, 0.2); color: #ef4444; font-weight: 600'] * len(row)
            elif age > 15: return ['background-color: rgba(245, 158, 11, 0.2); color: #f59e0b; font-weight: 600'] * len(row)
            else: return ['background-color: rgba(16, 185, 129, 0.2); color: #10b981; font-weight: 600'] * len(row)

        st.dataframe(
            view_link_table.style.apply(highlight_link, axis=1),
            use_container_width=True, hide_index=True,
            column_config={
                "Total Amount": st.column_config.NumberColumn("Total Amount (SAR)", format="%,.2f SAR"),
                "Percentage": st.column_config.NumberColumn("Liability Share", format="%.2f%%"),
                "Max_Age_Days": st.column_config.NumberColumn("Peak Age (Days)"),
                "Linked CN Count": st.column_config.NumberColumn("Volume (CNs)")
            }
        )

        st.download_button(
            label="Download Linked PO CNs Aging Matrix (CSV)",
            data=convert_df_to_csv(view_link_table),
            file_name=f"Linked_PO_CNs_Aging_Summary_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv"
        )

        # Interactive Expandable Vendor Breakdown
        with st.expander("Click to view Granular Vendor Document Breakdown (Linked PO CNs)", expanded=False):
            lk_vendor_list = display_link_table['Vendor Code'].unique().tolist()
            lk_vendor_map = {r['Vendor Code']: f"{r['Vendor Code']} - {r['Vendor Name']} | SAR {r['Total Amount']:,.2f} ({r['Linked CN Count']} CNs)" for _, r in display_link_table.iterrows()}
            
            selected_lk_vcode = st.selectbox(
                "Select Vendor to Inspect Documents:", 
                options=lk_vendor_list, 
                format_func=lambda x: lk_vendor_map.get(x, x),
                key="sb_link_vendor"
            )
            if selected_lk_vcode:
                v_df = linked_cn_df[linked_cn_df['buyFromVendorNo'] == selected_lk_vcode]
                st.dataframe(
                    v_df, 
                    use_container_width=True, hide_index=True,
                    column_config={
                        "amount": st.column_config.NumberColumn(format="%,.2f"),
                        "AmountWithVAT": st.column_config.NumberColumn(format="%,.2f SAR"),
                        "documentDate": st.column_config.DateColumn(format="YYYY-MM-DD")
                    }
                )
    else:
        st.info("No Linked PO Credit Notes data available for the selected range.")

# ==========================================
# PAGE 5: AUDIT TRAIL & LOGS
# ==========================================
elif page == "Audit Trail & Logs":
    st.markdown("<h2>System Audit Trail & Activity Logs</h2>", unsafe_allow_html=True)
    st.write("Complete historical record of operational actions, gate check-ins, user authentication, and stamp tracking.")
    st.markdown("---")
    
    if len(st.session_state.audit_log) > 0:
        df_audit = pd.DataFrame(st.session_state.audit_log)
        st.dataframe(df_audit, use_container_width=True, hide_index=True)
        
        csv_audit = convert_df_to_csv(df_audit)
        st.download_button(
            label="Download Audit Trail (CSV)",
            data=csv_audit,
            file_name=f"System_Audit_Trail_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )
    else:
        st.write("No recorded audit activities in the current session.")