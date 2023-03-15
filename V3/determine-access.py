import pandas as pd
import numpy as np
import glob
import os

# Read the userlist Excel file
userlist_df = pd.read_excel("C:/Users/Kawser/Downloads/selisech-user-licenses.xlsx")

# Get a list of all project Excel files in the directory
project_files = glob.glob("G:/Azure-Devops-Board-Userlist-Extract-Automation/selisech/*.xlsx")

# Loop through each project Excel file
for project_file in project_files:
    # Read the project user names from the Excel file
    # Read the project user names from the Excel file
    project_df = pd.read_excel(project_file)

    # Find the user names that are not in the userlist_df
    not_found_df = project_df[~project_df['User Names'].isin(userlist_df['User Names'])]

    # Print the user names to console
    if not_found_df.size > 0:
        print("\n\nUser names not found in userlist for file: " + os.path.basename(project_file))
        print(np.array(not_found_df['User Names']))

    # Merge the two DataFrames on the "name" column, keeping only the rows that match
    merged_df = pd.merge(userlist_df, project_df, how="inner", left_on="User Names", right_on="User Names")

    # Create a new DataFrame with the "name", "email", and "Access Level" columns from the userlist_df,
    # and the "Access Level" column from the merged_df
    result_df = pd.DataFrame({
        "User Names": merged_df["User Names"],
        "Email": merged_df["Email"],
        "Access Level": merged_df["Access Level"]
    })
    result_df.sort_values("Access Level", inplace=True)

    # Get the project name from the file path and create a new file name
    project_name = os.path.basename(project_file).split(".")[0]

    result_file_name = f"G:/Azure-Devops-Board-Userlist-Extract-Automation/selisech-a/{project_name}.xlsx"

    # Write the result DataFrame to a new Excel file
    result_df.to_excel(result_file_name, index=False)

    # Print duplicate user names to console
    duplicates_df = merged_df[merged_df.duplicated(subset='User Names', keep=False)]
    if len(duplicates_df) > 0:
        print("\n\nDuplicate user names found in file: " + os.path.basename(project_file))
        print(np.array(duplicates_df))
