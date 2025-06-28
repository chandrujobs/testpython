import pandas as pd
import os

# Define paths
folder_path = r"C:\Users\admin\Documents\CKYC Python"
base_file = os.path.join(folder_path, "Base data.xlsx")
audit_file = os.path.join(folder_path, "Custom audit report.xlsx")
mapping_folder = os.path.join(folder_path, "mapping_files")  # for future

# Load base and audit
base_df = pd.read_excel(base_file)
audit_df = pd.read_excel(audit_file)

# Strip column names
base_df = base_df.rename(columns=lambda x: x.strip())
audit_df = audit_df.rename(columns=lambda x: x.strip())

# Normalize dates
for df in [base_df, audit_df]:
    for col in ["Approved/Disbursed Date", "App Form DisbursalDate", "Appform Approval Date", "Appform Posting Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.date

# Define unique keys
keys = ["Applicant_id", "UCIC"]

# Prepare fallback date in custom audit report
def get_final_date(row):
    if pd.notna(row.get("App Form DisbursalDate")):
        return row["App Form DisbursalDate"]
    elif pd.notna(row.get("Appform Approval Date")):
        return row["Appform Approval Date"]
    elif pd.notna(row.get("Appform Posting Date")):
        return row["Appform Posting Date"]
    else:
        return pd.NaT

audit_df["Final Approved Date"] = audit_df.apply(get_final_date, axis=1)

# Merge Final Approved Date into base
merged_df = base_df.merge(
    audit_df[keys + ["Final Approved Date"]],
    on=keys, how="left"
)

# Update "Approved/Disbursed Date" if blank
merged_df["Approved/Disbursed Date"] = merged_df["Approved/Disbursed Date"].combine_first(merged_df["Final Approved Date"])

# Drop helper column
merged_df.drop(columns=["Final Approved Date"], inplace=True)

# Future: scan mapping folder
if os.path.exists(mapping_folder):
    mapping_files = [f for f in os.listdir(mapping_folder) if f.endswith('.xlsx')]
    print(f"📁 Mapping files ready: {mapping_files}")
else:
    print("📁 No mapping folder found (you can create one named 'mapping_files').")

# Save updated base file
merged_df.to_excel(base_file, index=False)

print("✅ Done! Base file updated using Custom Audit Report fallback logic.")
