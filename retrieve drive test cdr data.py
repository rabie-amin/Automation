import os
import pandas as pd

# Configuration
root_dir = r"C:\Users\XXXXXXXXXXXXXXXXXXX"
file_prefix = "CDR"
xlsm_file_paths = []
bad_paths = []
max_path_length = 0

# Walk through directories and collect .xlsm files
for dirpath, dirnames, filenames in os.walk(root_dir):
    for filename in filenames:
        if filename.lower().startswith(file_prefix.lower()) and filename.lower().endswith(".xlsm"):
            full_path = os.path.join(dirpath, filename)

            # Don't add \\?\ prefix — just use UNC path as-is for network shares
            safe_path = full_path

            if not os.path.exists(safe_path):
                bad_paths.append(safe_path)
            else:
                max_path_length = max(max_path_length, len(safe_path))
                xlsm_file_paths.append(safe_path)

print(f"\n✅ Number of valid files: {len(xlsm_file_paths)}")
print(f"❌ Number of invalid/inaccessible files: {len(bad_paths)}")

# Collect data from all valid files
collected_data = []
for path in xlsm_file_paths:
    try:
        df = pd.read_excel(
            path,
            sheet_name="CDR",
            engine="openpyxl",
            # usecols=["A File Name","A Call Start Time", "A Home Operator", "A Call Type","E2E call Status",""]
            usecols=["A File Name",	"A Call Start Time",	"A Device Label",	"A Home Operator",	"A Call Type",	"E2E call Status","A Dial System Band","CST (Dial->A_Connect)","CST (Dial->A_Alerting)","CST (Dial->B_Alerting)","CST (Dial->A_Connect - B_Alerting->B_connect)","A MOS avg","B MOS avg"]
        )
        df["file path"] = "West Amman"
        collected_data.append(df)
    except Exception as e:
        print(f"⚠️ Error reading file {path}:\n{e}\n")

# Combine all into a single DataFrame
if collected_data:
    all = pd.concat(collected_data, ignore_index=True)
    print(f"\n✅ Successfully collected data from {len(collected_data)} files.")
else:
    all = pd.DataFrame()
    print("\n⚠️ No valid data extracted.")
all.to_csv(r"C:\Users\XXXXXXXXXXXXXXXXX\all.csv")
print("Saved...")
