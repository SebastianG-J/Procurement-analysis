import pandas as pd

# Load both Excel files
df_simplified = pd.read_excel("results_merged.xlsx")
df_new = pd.read_excel("results_new.xlsx")

# Append new data to the bottom
df_merged = pd.concat([df_simplified, df_new], ignore_index=True)

# Save to a new file (or overwrite one of the originals)
df_merged.to_excel("results_merged1.xlsx", index=False)

print("Merge complete. Saved as 'results_merged.xlsx'")