import pandas as pd
import os

# File path
folder_path = r"C:\Users\admin\Documents\CKYC Python"
base_file = os.path.join(folder_path, "Base data.xlsx")

# Load base file
df = pd.read_excel(base_file)
df = df.rename(columns=lambda x: x.strip())

# Convert columns to date (safely)
for col in ["Approved/Disbursed Date", "App Form DisbursalDate", "Appform Approval Date", "Appform Posting Date"]:
    if col in df.columns:
        df[col] = pd.to_datetime(df[col], errors='coerce').dt.date

# Apply fallback logic only if Approved/Disbursed Date is empty
def fill_approved_date(row):
    if pd.notna(row["Approved/Disbursed Date"]):
        return row["Approved/Disbursed Date"]
    elif pd.notna(row.get("App Form DisbursalDate")):
        return row["App Form DisbursalDate"]
    elif pd.notna(row.get("Appform Approval Date")):
        return row["Appform Approval Date"]
    elif pd.notna(row.get("Appform Posting Date")):
        return row["Appform Posting Date"]
    else:
        return pd.NaT

df["Approved/Disbursed Date"] = df.apply(fill_approved_date, axis=1)

# Save back to base file
df.to_excel(base_file, index=False)

print("✅ Approved/Disbursed Date updated using fallback logic from within Base file.")
