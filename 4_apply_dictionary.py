from openpyxl import load_workbook
import pandas as pd
import shutil
from datetime import datetime

PRODUCT_FILE = "Product_file.xlsx"
DICT_FILE = "filter_dictionary.xlsx"

SKU_COL = 2
FILTER_START_COL = 85
FILTER_HEADER_ROW = 3
DATA_START_ROW = 4

# Backup
backup_name = f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}_Product_file.xlsx"
shutil.copy(PRODUCT_FILE, backup_name)
print(f"Backup created: {backup_name}")

# Load dictionary
dictionary = pd.read_excel(DICT_FILE)
dictionary = dictionary.dropna(subset=["Canonical Value"])

def norm(x):
    return str(x).strip().lower()

dictionary["Filter Name"] = dictionary["Filter Name"].apply(norm)
dictionary["Filter Value"] = dictionary["Filter Value"].apply(norm)

mapping = {
    (row["Filter Name"], row["Filter Value"]): row["Canonical Value"].strip()
    for _, row in dictionary.iterrows()
}

print(f"Loaded {len(mapping)} dictionary mappings.")

# Load workbook
wb = load_workbook(PRODUCT_FILE)

total_changes = 0

for ws in wb.worksheets:
    print(f"\nProcessing sheet: {ws.title}")

    # Read headers
    filter_headers = []
    for col in range(FILTER_START_COL, ws.max_column + 1):
        val = ws.cell(row=FILTER_HEADER_ROW, column=col).value
        filter_headers.append(norm(val) if val else "")

    sheet_changes = 0

    for row in range(DATA_START_ROW, ws.max_row + 1):
        for offset, filter_name in enumerate(filter_headers):
            col = FILTER_START_COL + offset
            cell = ws.cell(row=row, column=col)

            if cell.value is None:
                continue

            original = str(cell.value).strip()
            key = (filter_name, original.lower())

            if key in mapping:
                new_val = mapping[key]
                if original != new_val:
                    cell.value = new_val
                    sheet_changes += 1

    print(f"  → {sheet_changes} changes")
    total_changes += sheet_changes

# Save
wb.save(PRODUCT_FILE)

print(f"\nDone. Total changes across all sheets: {total_changes}")
