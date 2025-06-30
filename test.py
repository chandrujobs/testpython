import pandas as pd
import os

# 📁 Folder path and filenames
folder = r"C:\Users\Admin\Documents\CKYC Python"
base_file = os.path.join(folder, "CKYC BASE DATA.xlsx")
audit_file = os.path.join(folder, "Custom_audit_report.xlsx")

# 📥 Load Excel files
base_df = pd.read_excel(base_file)
audit_df = pd.read_excel(audit_file)

# 🧹 Clean column names and values
base_df.columns = base_df.columns.str.strip()
audit_df.columns = audit_df.columns.str.strip()
audit_df = audit_df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

# ----------------------------------------------------------------------------
# ✅ Step 1: Compute 'Approved/Disbursed Date'
# ----------------------------------------------------------------------------
date_cols = ["App Form DisbursalDate", "Appform Approval Date", "Appform Posting Date"]
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
# 🗓️ Step 2: Extract Month from Disbursed Date
# ----------------------------------------------------------------------------
base_df["Month"] = pd.to_datetime(base_df["Approved/Disbursed Date"], errors='coerce').dt.strftime('%b')

# ----------------------------------------------------------------------------
# 🔄 Step 3: Extract Applicant_id from audit file
# ----------------------------------------------------------------------------
if "Los App Id" in audit_df.columns:
    audit_df["Applicant_id"] = audit_df["Los App Id"].astype(str).str.extract(r'_(\d+)$')[0].str.strip()

base_df["Applicant_id"] = base_df["Applicant_id"].astype(str).str.strip()
audit_df["Applicant_id"] = audit_df["Applicant_id"].astype(str).str.strip()

# ----------------------------------------------------------------------------
# 🔄 Step 4: Map Status.1 and Workflow from audit
# ----------------------------------------------------------------------------
if "CKYC Status" in audit_df.columns:
    status_map = audit_df.set_index("Applicant_id")["CKYC Status"].to_dict()
    base_df["Status.1"] = base_df["Applicant_id"].map(status_map)

if "Status" in audit_df.columns:
    workflow_map = audit_df.set_index("Applicant_id")["Status"].to_dict()
    base_df["Workflow"] = base_df["Applicant_id"].map(workflow_map)

# ----------------------------------------------------------------------------
# ✅ Step 5: Final Status from Status.1
# ----------------------------------------------------------------------------
completed_keywords = ["ckyc completed", "ckyc completed - manual", "issue with ckyc"]
pending_keywords = ["auto resolution", "ckyc upload pending", "pending with ckyc team", "under resolution"]

def determine_final_status(status1):
    if pd.isna(status1):
        return ""
    status_lower = str(status1).strip().lower()
    if status_lower in completed_keywords:
        return "Completed"
    elif status_lower in pending_keywords:
        return "Pending"
    return ""

base_df["Final Status"] = base_df["Status.1"].apply(determine_final_status)

# ----------------------------------------------------------------------------
# 📆 Step 6: InwardDate from Triggered Date
# ----------------------------------------------------------------------------
if "Triggered Date" in audit_df.columns:
    triggered_map = audit_df.set_index("Applicant_id")["Triggered Date"].to_dict()
    base_df["InwardDate"] = pd.to_datetime(base_df["Applicant_id"].map(triggered_map), errors='coerce').dt.date

# ----------------------------------------------------------------------------
# ✅ Step 7: Completion Date
# ----------------------------------------------------------------------------
if "CKYC Completion Date" in audit_df.columns:
    complete_map = audit_df.set_index("Applicant_id")["CKYC Completion Date"].to_dict()
    base_df["Completion Date"] = pd.to_datetime(base_df["Applicant_id"].map(complete_map), errors='coerce').dt.date

# ----------------------------------------------------------------------------
# ✅ Step 8: CKYC Number
# ----------------------------------------------------------------------------
if "CKYC Number" in audit_df.columns:
    ckyc_map = audit_df.set_index("Applicant_id")["CKYC Number"].to_dict()
    base_df["CKYC Number"] = base_df["Applicant_id"].map(ckyc_map)

# ----------------------------------------------------------------------------
# ✅ Step 9: Aging = Disbursed - Completion
# ----------------------------------------------------------------------------
base_df["Aging"] = (
    pd.to_datetime(base_df["Approved/Disbursed Date"], errors='coerce') -
    pd.to_datetime(base_df["Completion Date"], errors='coerce')
).dt.days

# ----------------------------------------------------------------------------
# ✅ Step 10: TAT = Inward - Disbursed
# ----------------------------------------------------------------------------
disbursed = pd.to_datetime(base_df["Approved/Disbursed Date"], errors='coerce')
inward = pd.to_datetime(base_df["InwardDate"], errors='coerce')
base_df["TAT"] = (inward - disbursed).dt.days

# ----------------------------------------------------------------------------
# ✅ Step 11: CKYC ID Length
# ----------------------------------------------------------------------------
base_df["CKYC ID Length"] = base_df["CKYC Number"].apply(lambda x: len(str(x)) if pd.notna(x) else "")

# ----------------------------------------------------------------------------
# 💾 Step 12: Save final output to same file
# ----------------------------------------------------------------------------
base_df.to_excel(base_file, index=False)

# ----------------------------------------------------------------------------
# ✅ Summary
# ----------------------------------------------------------------------------
print("\n✅ CKYC BASE DATA.xlsx updated with:")
print(" - Approved/Disbursed Date")
print(" - Month")
print(" - Status.1")
print(" - Workflow")
print(" - Final Status")
print(" - InwardDate")
print(" - Completion Date")
print(" - CKYC Number")
print(" - CKYC ID Length")
print(" - Aging")
print(" - TAT")
