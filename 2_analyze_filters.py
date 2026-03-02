import pandas as pd

INPUT_FILE = "filter_review.xlsx"
OUTPUT_FILE = "filter_summary.xlsx"

print("Loading extracted filters...")

df = pd.read_excel(INPUT_FILE)

# Group by Filter Name + Filter Value and count
summary = (
    df
    .groupby(["Filter Name", "Filter Value"])
    .size()
    .reset_index(name="Count")
    .sort_values(["Filter Name", "Count"], ascending=[True, False])
)

# Save summary
summary.to_excel(OUTPUT_FILE, index=False)

print(f"Done. Summary saved to {OUTPUT_FILE}")
