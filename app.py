import pandas as pd
import os
import sys

# Folder path
folder = r"C:\Users\Admin\Documents\CKYC Python"

# 🔍 Dynamically detect files
def find_file(keyword):
    for f in os.listdir(folder):
        if keyword.lower() in f.lower() and f.lower().endswith('.xlsx'):
            return os.path.join(folder, f)
    return None

# File detection
base_file = find_file("CKYC BASE DATA")
audit_file = find_file("Custom_audit_report")

if not base_file:
    sys.exit("❌ Error: 'CKYC BASE DATA.xlsx' not found in folder.")
if not audit_file:
    sys.exit("❌ Error: 'Custom_audit_report.xlsx' not found in folder.")

# Load Excel files
try:
    base_df = pd.read_excel(base_file)
    audit_df = pd.read_excel(audit_file)
except Exception as e:
    sys.exit(f"❌ Error reading Excel files: {str(e)}")

# Clean columns and strings
base_df.columns = base_df.columns.str.strip()
audit_df.columns = audit_df.columns.str.strip()
audit_df = audit_df.map(lambda x: x.strip() if isinstance(x, str) else x)

# ----------------------------------------------------------------------------
# Approved/Disbursed Date
# ----------------------------------------------------------------------------
date_cols = ["App Form DisbursalDate", "Appform Approval Date", "Recent Status Date"]
for col in date_cols:
    if col in base_df.columns:
        base_df[col] = pd.to_datetime(base_df[col], errors='coerce').dt.date

def get_disbursed_date(row):
    for col in date_cols:
        if pd.notna(row.get(col)):
            return row[col]
    return pd.NaT

base_df["Approved/Disbursed Date"] = base_df.apply(get_disbursed_date, axis=1)

# ----------------------------------------------------------------------------
# Month Column
# ----------------------------------------------------------------------------
base_df["Month"] = pd.to_datetime(base_df["Approved/Disbursed Date"], errors='coerce').dt.strftime("%b'%y")

# ----------------------------------------------------------------------------
# Applicant ID mapping
# ----------------------------------------------------------------------------
if "Los App Id" in audit_df.columns:
    audit_df["Applicant_id"] = audit_df["Los App Id"].astype(str).str.extract(r'_(\d+)$')[0].str.strip()

base_df["Applicant_id"] = base_df["Applicant_id"].astype(str).str.strip()
audit_df["Applicant_id"] = audit_df["Applicant_id"].astype(str).str.strip()

# ----------------------------------------------------------------------------
# Workflow Mapping from Status
# ----------------------------------------------------------------------------
if "Status" in audit_df.columns:
    workflow_map = audit_df.set_index("Applicant_id")["Status"].to_dict()
    base_df["Workflow"] = base_df["Applicant_id"].map(workflow_map)

# ----------------------------------------------------------------------------
# CKYC Status from Workflow
# ----------------------------------------------------------------------------
workflow_status_map = {
    "Auto resolution": [
        "download_auth_failed_notify", "download_initiated", "download_processing_pending",
        "download_submitted", "download_uploaded", "failed_timer_pending", "initiated",
        "operation_decision_failed", "probable_match_submitted", "submitted", "triggered"
    ],
    "CKYC Completed": ["ckyc_number_updated"],
    "Under Resolution with Ops": [
        "download_auth_failed", "manual_review", "post_processing_download_auth_failed",
        "search_and_download_validation_failed"
    ],
    "Pending with CKYC Team": ["probable_match_uploaded", "processed_awaiting_response", "uploaded"],
    "Issue with CKYC": ["download_auth_failed_notified"],
    "Manually Reported by Ops": ["Manually Reported by Ops"],
    "CKYC Upload Pending": ["CKYC Upload Pending"]
}

def map_ckyc_status(workflow):
    for status, workflows in workflow_status_map.items():
        if str(workflow).strip() in workflows:
            return status
    return ""

base_df["CKYC Status"] = base_df["Workflow"].apply(map_ckyc_status)

# ----------------------------------------------------------------------------
# Final Status
# ----------------------------------------------------------------------------
completed_keywords = ["ckyc completed", "ckyc completed - manual", "issue with ckyc", "manually reported by ops", "under resolution with ops"]
pending_keywords = ["auto resolution", "ckyc upload pending", "pending with ckyc team"]

def determine_final_status(status1):
    if pd.isna(status1):
        return ""
    status_lower = str(status1).strip().lower()
    if status_lower in completed_keywords:
        return "Completed"
    elif status_lower in pending_keywords:
        return "Pending"
    return ""

base_df["Final Status"] = base_df["CKYC Status"].apply(determine_final_status)

# ----------------------------------------------------------------------------
# InwardDate Mapping
# ----------------------------------------------------------------------------
if "Triggered Date" in audit_df.columns:
    triggered_date_map = audit_df.set_index("Applicant_id")["Triggered Date"].to_dict()
    base_df["InwardDate"] = pd.to_datetime(base_df["Applicant_id"].map(triggered_date_map), errors='coerce').dt.date

# ----------------------------------------------------------------------------
# Completion Date
# ----------------------------------------------------------------------------
if "CKYC Completion Date" in audit_df.columns:
    completion_map = audit_df.set_index("Applicant_id")["CKYC Completion Date"].to_dict()
    base_df["Completion Date"] = pd.to_datetime(base_df["Applicant_id"].map(completion_map), errors='coerce').dt.date

# ----------------------------------------------------------------------------
# CKYC Number
# ----------------------------------------------------------------------------
if "CKYC Number" in audit_df.columns:
    number_map = audit_df.set_index("Applicant_id")["CKYC Number"].to_dict()
    base_df["CKYC Number"] = base_df["Applicant_id"].map(number_map)

# ----------------------------------------------------------------------------
# CKYC Upload Date
# ----------------------------------------------------------------------------
if "First Batch Upload Date" in audit_df.columns:
    upload_map = audit_df.set_index("Applicant_id")["First Batch Upload Date"].to_dict()
    base_df["CKYC Upload Date"] = pd.to_datetime(base_df["Applicant_id"].map(upload_map), errors='coerce').dt.date

# ----------------------------------------------------------------------------
# CKYC Reporting TAT = CKYC Upload Date - Disbursed Date
# ----------------------------------------------------------------------------
base_df["CKYC Reporting TAT"] = (
    pd.to_datetime(base_df["CKYC Upload Date"], errors='coerce') -
    pd.to_datetime(base_df["Approved/Disbursed Date"], errors='coerce')
).dt.days

# ----------------------------------------------------------------------------
# CKYC Trigger TAT = InwardDate - Disbursed Date
# ----------------------------------------------------------------------------
base_df["CKYC Trigger TAT"] = (
    pd.to_datetime(base_df["InwardDate"], errors='coerce') -
    pd.to_datetime(base_df["Approved/Disbursed Date"], errors='coerce')
).dt.days

# ----------------------------------------------------------------------------
# CKYC ID Length
# ----------------------------------------------------------------------------
base_df["CKYC ID Length"] = base_df["CKYC Number"].apply(lambda x: len(str(x)) if pd.notna(x) else "")

# ----------------------------------------------------------------------------
# Product Name Mapping (insert after "Partner Id")
# ----------------------------------------------------------------------------
product_map = {
    "SEP": "SEP",
    "AIR": "Embedded Finance", "ANG": "Embedded Finance", "CLP": "Embedded Finance", "ETC": "Embedded Finance",
    "GRO": "Embedded Finance", "INC": "Embedded Finance", "JAR": "Embedded Finance", "NBR": "Embedded Finance",
    "NRO": "Embedded Finance", "OLA": "Embedded Finance", "ONL": "Embedded Finance", "PEL": "Embedded Finance",
    "SPM": "Embedded Finance",
    "LAP": "LAP", "LPA": "LAP", "LPD": "LAP", "HLD": "LAP", "PLP": "LAP",
    "PCL": "Fintech Partnership", "AVN": "Fintech Partnership", "BPT": "Fintech Partnership", "CRC": "Fintech Partnership",
    "ESC": "Fintech Partnership", "JPT": "Fintech Partnership", "KBL": "Fintech Partnership", "MTC": "Fintech Partnership",
    "MVL": "Fintech Partnership", "PRL": "Fintech Partnership", "PSE": "Fintech Partnership", "UNC": "Fintech Partnership",
    "ZM": "Fintech Partnership",
    "SBA": "SBL", "SBD": "SBL", "SBL": "SBL",
    "UBL": "UBL", "UPL": "UPL", "NVI": "NVI", "LKB": "LKB", "AFB": "AFB", "WSL": "WSL"
}

base_df["Product Name"] = base_df["Loan Product"].map(product_map)

# Move 'Product Name' next to 'Partner Id'
if "Partner Id" in base_df.columns:
    cols = list(base_df.columns)
    idx = cols.index("Partner Id") + 1
    cols.insert(idx, cols.pop(cols.index("Product Name")))
    base_df = base_df[cols]

# ----------------------------------------------------------------------------
# Save updated Excel
# ----------------------------------------------------------------------------
try:
    base_df.to_excel(base_file, index=False)
    print("\n✅ CKYC BASE DATA.xlsx updated with:")
    print(" - Approved/Disbursed Date")
    print(" - Month")
    print(" - Workflow")
    print(" - CKYC Status")
    print(" - Final Status")
    print(" - InwardDate")
    print(" - Completion Date")
    print(" - CKYC Number")
    print(" - CKYC Upload Date")
    print(" - CKYC Reporting TAT")
    print(" - CKYC Trigger TAT")
    print(" - CKYC ID Length")
    print(" - Product Name")
except Exception as e:
    print(f"❌ Failed to save Excel file: {str(e)}")
