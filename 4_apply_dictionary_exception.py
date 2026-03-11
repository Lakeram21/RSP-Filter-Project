import pandas as pd
import shutil
from datetime import datetime

PRODUCT_FILE = "Product_file.xlsx"
DICT_FILE = "filter_dictionary.xlsx"

SKU_COL = 2
FILTER_START_COL = 85
FILTER_HEADER_ROW = 3
DATA_START_ROW = 4

# ---------------------------------
# Create Backup
# ---------------------------------
timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
backup_name = f"backup_{timestamp}_Product_file.xlsx"
shutil.copy(PRODUCT_FILE, backup_name)
print(f"Backup created: {backup_name}")

# ---------------------------------
# Create Log File
# ---------------------------------
log_filename = f"change_log_{timestamp}.txt"
log_file = open(log_filename, "w", encoding="utf-8")

# ---------------------------------
# Load dictionary
# ---------------------------------
dictionary = pd.read_excel(DICT_FILE)

def norm(x):
    return str(x).strip().lower()

dictionary = dictionary.dropna(subset=["Canonical Value"])
dictionary["Filter Name"] = dictionary["Filter Name"].apply(norm)
dictionary["Filter Value"] = dictionary["Filter Value"].apply(norm)

mapping = {
    (row["Filter Name"], row["Filter Value"]): row["Canonical Value"].strip()
    for _, row in dictionary.iterrows()
}

print(f"Loaded {len(mapping)} dictionary mappings")

# ---------------------------------
# Read product file using CALAMINE
# ---------------------------------
print("Reading product file with calamine...")

sheets = pd.read_excel(
    PRODUCT_FILE,
    sheet_name=None,
    header=None,
    engine="calamine"
)

updated_sheets = {}
total_changes = 0

# ---------------------------------
# Process sheets
# ---------------------------------
for sheet_name, df in sheets.items():

    print(f"\nProcessing sheet: {sheet_name}")
    log_file.write(f"\nProcessing sheet: {sheet_name}\n")

    filter_headers = df.iloc[FILTER_HEADER_ROW-1, FILTER_START_COL-1:].apply(norm)

    sheet_changes = 0

    for row in range(DATA_START_ROW-1, len(df)):
        sku = df.iat[row, SKU_COL-1]

        for offset, filter_name in enumerate(filter_headers):
            col = FILTER_START_COL-1 + offset

            val = df.iat[row, col]

            if pd.isna(val):
                continue

            original = str(val).strip()
            key = (filter_name, original.lower())

            if key in mapping:
                new_val = mapping[key]

                if original != new_val:
                    df.iat[row, col] = new_val
                    sheet_changes += 1

                    log_file.write(
                        f"Sheet: {sheet_name} | "
                        f"SKU: {sku} | "
                        f"Row: {row+1} | "
                        f"Column: {col+1} | "
                        f"Filter: {filter_name} | "
                        f"Old: '{original}' → New: '{new_val}'\n"
                    )

    print(f"  → {sheet_changes} changes")
    total_changes += sheet_changes
    updated_sheets[sheet_name] = df

# ---------------------------------
# Write NEW Excel file
# ---------------------------------
OUTPUT_FILE = "Product_file_updated.xlsx"

with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
    for name, sheet in updated_sheets.items():
        sheet.to_excel(writer, sheet_name=name, header=False, index=False)

log_file.write(f"\nTotal changes: {total_changes}\n")
log_file.close()

print(f"\nDone. Total changes: {total_changes}")
print(f"Updated file saved as: {OUTPUT_FILE}")
print(f"Log saved to: {log_filename}")