import pandas as pd

# Read the userlist Excel file
userlist_df = pd.read_excel("C:/Users/Kawser/Downloads/selisech-user-licenses.xlsx")

# Read the Läderach project user names from the Excel file
laderch_df = pd.read_excel("C:/Users/Kawser/kiosk.xlsx")

# Merge the two DataFrames on the "name" column, keeping only the rows that match
merged_df = pd.merge(userlist_df, laderch_df, how="inner", left_on="User Names", right_on="User Names")

# Create a new DataFrame with the "name", "email", and "Access Level" columns from the userlist_df,
# and the "Access Level" column from the merged_df
result_df = pd.DataFrame({
    "User Names": merged_df["User Names"],
    "Email": merged_df["Email"],
    "Access Level": merged_df["Access Level"]
})

# Write the result DataFrame to a new Excel file
result_df.to_excel("access_levelskiosk.xlsx", index=False)