import pandas as pd
import os

SUMMARY_FILE = "filter_summary.xlsx"
DICT_FILE = "filter_dictionary.xlsx"

summary = pd.read_excel(SUMMARY_FILE)

summary["Filter Name"] = summary["Filter Name"].astype(str).str.strip()
summary["Filter Value"] = summary["Filter Value"].astype(str).str.strip()

# Load or create dictionary
if os.path.exists(DICT_FILE):
    dictionary = pd.read_excel(DICT_FILE)
else:
    dictionary = pd.DataFrame(columns=[
        "Filter Name", "Filter Value", "Canonical Value", "Count"
    ])

dictionary["Filter Name"] = dictionary["Filter Name"].astype(str).str.strip()
dictionary["Filter Value"] = dictionary["Filter Value"].astype(str).str.strip()

# ----------------------------
# Build safe lookup (dict only)
# ----------------------------
existing = {}

# Load existing dictionary into clean dict
for _, row in dictionary.iterrows():
    key = (row["Filter Name"], row["Filter Value"])
    existing[key] = {
        "Filter Name": row["Filter Name"],
        "Filter Value": row["Filter Value"],
        "Canonical Value": "" if pd.isna(row["Canonical Value"]) else row["Canonical Value"],
        "Count": int(row["Count"]) if not pd.isna(row["Count"]) else 0
    }

# Apply summary updates
for _, row in summary.iterrows():
    key = (row["Filter Name"], row["Filter Value"])

    if key in existing:
        # Update count only
        existing[key]["Count"] = int(row["Count"])
    else:
        # Add new entry
        existing[key] = {
            "Filter Name": row["Filter Name"],
            "Filter Value": row["Filter Value"],
            "Canonical Value": "",
            "Count": int(row["Count"])
        }

# Convert back to DataFrame
final = pd.DataFrame(list(existing.values()))

# Sort nicely
final = final.sort_values(["Filter Name", "Count"], ascending=[True, False])

# Save
final.to_excel(DICT_FILE, index=False)

print("Dictionary updated correctly.")
print("Existing entries preserved, new ones added.")
