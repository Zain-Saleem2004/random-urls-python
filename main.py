import pandas as pd
import random
import json
import webbrowser
import os

# File paths
excel_file = "urls.xlsx"
storage_file = "shown_links.json"

# Load URLs from Excel
try:
    df = pd.read_excel(excel_file)
    urls = df["URL"].dropna().tolist()
except Exception as e:
    print(f"Error reading Excel file: {e}")
    exit()

# Load shown indexes
if os.path.exists(storage_file):
    with open(storage_file, "r") as f:
        shown_indexes = json.load(f)
else:
    shown_indexes = []

# Calculate remaining indexes
remaining_indexes = [i for i in range(len(urls)) if i not in shown_indexes]

# Reset if all used
if not remaining_indexes:
    shown_indexes = []
    remaining_indexes = list(range(len(urls)))

# Pick a random URL
chosen_index = random.choice(remaining_indexes)
chosen_url = urls[chosen_index]

# Save updated shown indexes
shown_indexes.append(chosen_index)
with open(storage_file, "w") as f:
    json.dump(shown_indexes, f)

# Open in browser
webbrowser.open(chosen_url)
