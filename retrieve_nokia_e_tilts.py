import pandas as pd
import xlwings as xw
import re
import sys
import os

def main():
    # Step 1: Ask user for the input XLSB file path
    file_path = input("Enter full path to the .xlsb file: ").strip()
    if not os.path.isfile(file_path):
        print("❌ File does not exist. Exiting.")
        sys.exit(1)

    print("📂 Opening Excel file...")

    # Step 2: Open workbook and sheet
    try:
        wb = xw.Book(file_path)
        sheet = wb.sheets['RETU']
    except Exception as e:
        print(f"❌ Error opening Excel file or sheet 'RETU': {e}")
        sys.exit(1)

    # Step 3: Detect where the real header starts (skip any junk rows)
    data_raw = sheet.used_range.options(pd.DataFrame, header=False, index=False).value

    # Find the header row (contains "MRBTS")
    header_row_index = None
    for idx, row in data_raw.iterrows():
        if any(str(cell).strip() == "MRBTS" for cell in row):
            header_row_index = idx
            break

    if header_row_index is None:
        print("❌ Could not find a row with 'MRBTS' column.")
        wb.close()
        sys.exit(1)

    # Extract only the relevant data starting from the header row
    data = data_raw.iloc[header_row_index + 1:].copy()
    data.columns = data_raw.iloc[header_row_index].astype(str).str.strip()
    data = data.reset_index(drop=True)
    wb.close()

    print("🔍 Columns found:", data.columns.tolist())

    df = data

    if "MRBTS" not in df.columns:
        print("❌ 'MRBTS' column not found. Exiting.")
        sys.exit(1)

    # Step 4: Create "Cluster" column
    # Note: The original code used df["MRBTS"].astype(str).str[:3].astype(int) // 1000,
    # but taking first 3 characters then integer then // 1000 will always be zero.
    # Probably intended to use first 4 chars or something else.
    # I keep it as original, but be aware it may need adjustment.
    # def split_col(text):
    #     text_split=text.split(".")
    #     return text_split[0]
    # try:
       
    #     # df["Cluster"] = df["MRBTS"].astype(str).apply(split_col).str[:-4].astype(int) // 100
    #     # df["MRBTS"].astype(str).apply(split_col).str.slice(0, -4)

    #     print(df.head(10))
    # except Exception as e:
    #     print(f"❌ Error creating Cluster column: {e}")
    #     sys.exit(1)
    def split_col(text):
        text_split = text.split(".")
        # print(text_split[0][-4:])
        # print(type(text_split[0]))
        return text_split[0][-4:]

    try:
        df["Cluster"] = (
            df["MRBTS"]
            .astype(str)
            .apply(split_col)         # remove decimals
            .astype(int) // 100       # convert to int and divide
        )
        print(df.head())
    except Exception as e:
        print("❌ Error while creating Cluster column:", e)


    available_clusters = sorted(df['Cluster'].dropna().unique())
    cluster_input = input(f"Enter Cluster (Available: {available_clusters}, or 'All'): ").strip()
    if cluster_input.lower() != "all":
        try:
            cluster_val = int(cluster_input)
            df = df[df["Cluster"] == cluster_val]
        except ValueError:
            print("❌ Invalid cluster input. Exiting.")
            sys.exit(1)

    # Step 6: Create "Vendor" column
    if "sectorID" not in df.columns:
        print("❌ 'sectorID' column not found. Exiting.")
        sys.exit(1)

    df["Vendor"] = df["sectorID"].apply(lambda x: "Huawei" if "-H" in str(x) else "Nokia")

    vendor_input = input("Enter Vendor (Huawei, Nokia, or 'All'): ").strip()
    if vendor_input.lower() != "all":
        df = df[df["Vendor"].str.lower() == vendor_input.lower()]

    def get_carrier(sector_id):
        sector_id = str(sector_id)
        if "L2100" in sector_id:
            return 2100
        elif "L900" in sector_id:
            return 900
        elif "L1800" in sector_id:
            return 1800
        return None

    df["Carrier"] = df["sectorID"].apply(get_carrier)

    carrier_input = input("Enter Carrier (2100, 900, 1800, or 'All'): ").strip()
    if carrier_input.lower() != "all":
        try:
            carrier_val = int(carrier_input)
            df = df[df["Carrier"] == carrier_val]
        except ValueError:
            print("❌ Invalid carrier input. Exiting.")
            sys.exit(1)

    def extract_base_site_id(mrbts):
        mrbts_str = str(int(mrbts))
        length = len(mrbts_str)
        if length == 4:
            site = mrbts_str
        elif length == 5:
            if mrbts_str.startswith("50"):
                # Keep last 3 chars (from right)
                site = int(mrbts_str[-3:])
            else:
                # Keep last 4 chars (from right)
                site = int(mrbts_str[-4:])
        else:
            site = None
        return site



    df["Base_Site_ID"] = df["MRBTS"].apply(extract_base_site_id)

    def sector_name(sector_id):
        sector_id_str = str(sector_id[0])
        return sector_id_str
        # sector_id_str = str(int(sector_id))
        # right2 = sector_id_str[-2:]
        # right1 = sector_id_str[-1:]
        # if right2 in ["21", "24"]:
        #     return "A"
        # elif right2 in ["22", "25"]:
        #     return "B"
        # elif right2 in ["23", "26"]:
        #     return "C"
        # elif right1 == "1":
        #     return "A"
        # elif right1 == "2":
        #     return "B"
        # elif right1 == "3":
        #     return "C"
        # elif right1 in ["4", "7"]:
        #     return "D"
        # elif right1 == "9":
        #     return "E"
        # elif right1 == "6":
        #     return "F"
        # return "?"

    df["Sector_Name"] = df["sectorID"].apply(sector_name)

    df["Sector"] = df["Base_Site_ID"].astype(str) + "_" + df["Sector_Name"]

    columns_to_export = ["Sector", "Vendor", "Carrier", "angle", "minAngle", "maxAngle"]
    columns_to_export = [col for col in columns_to_export if col in df.columns]
    final_df = df[columns_to_export]
    print(final_df.head())

    save_choice = input("Do you want to save the final data to a CSV file? (yes/no): ").strip().lower()
    if save_choice in ["yes", "y"]:
        save_path = input("Enter full path where to save the CSV file (including .csv extension): ").strip()
        try:
            final_df.to_csv(save_path, index=False)
            print(f"✅ File saved to: {save_path}")
        except Exception as e:
            print(f"❌ Error saving file: {e}")
    else:
        print("✅ Operation completed. File not saved.")

if __name__ == "__main__":
    main()
