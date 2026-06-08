import pandas as pd

# file_path = r"C:/Users/Yash-Vichare/Downloads/New folder (3)/Life Cycle Report Dec-2025.csv"
file_path = r"C:/Users/Divyesh Parmar/Downloads/Life Cycle Report Dec-2025.csv"

# Read everything as text to avoid type conversion
df = pd.read_csv(file_path, sep=None, engine="python", dtype=str)

# Print the raw 'Agency Txn Id' column to inspect it
print("Raw 'Agency Txn Id' column values:")
print(df["Agency Txn Id"].head(10))

# Clean the 'Agency Txn Id' column:
# 1. Remove leading apostrophe using str.lstrip("'")
# 2. Keep the full value without splitting
df["Agency Txn Id"] = df["Agency Txn Id"].str.lstrip("'")

# Print only the cleaned "Agency Txn Id" column
print("\nCleaned 'Agency Txn Id' column values:")
# print(df[["Agency Txn Id"]].head(5))
print(df.head(5))