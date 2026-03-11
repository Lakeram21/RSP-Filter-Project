import pandas as pd

INPUT_FILE = "Product_file.xlsx"
OUTPUT_FILE = "filter_review.xlsx"

SKU_COL = 1
FILTER_START_COL = 84
FILTER_HEADER_ROW = 2
DATA_START_ROW = 3

print("Reading all sheets...")

all_records = []

# Load all sheets
try:
    sheets = pd.read_excel(INPUT_FILE, sheet_name=None, header=None)
except Exception as e:
    print(f"Error reading Excel file: {e}")
    sheets = pd.read_excel(
    INPUT_FILE,
    sheet_name=None,
    header=None,
    engine="calamine"
)

for sheet_name, raw in sheets.items():
    print(f"Processing sheet: {sheet_name}")

    filter_headers = raw.iloc[FILTER_HEADER_ROW, FILTER_START_COL:]

    for row_idx in range(DATA_START_ROW, raw.shape[0]):
        sku = raw.iat[row_idx, SKU_COL]

        if pd.isna(sku):
            continue

        for offset, filter_name in enumerate(filter_headers):
            col_idx = FILTER_START_COL + offset
            value = raw.iat[row_idx, col_idx]

            if pd.notna(value) and str(value).strip():
                all_records.append({
                    "Sheet": sheet_name,
                    "SKU": str(sku).strip(),
                    "Filter Name": str(filter_name).strip(),
                    "Filter Value": str(value).strip()
                })

df = pd.DataFrame(all_records)
df.to_excel(OUTPUT_FILE, index=False)

print(f"Done. Extracted {len(df)} rows across {len(sheets)} sheets.")
