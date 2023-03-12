import pandas as pd
import glob
import os

# Read the userlist Excel file
userlist_df = pd.read_excel("C:/Users/Kawser/Downloads/SELISESunrise-user-licenses.xlsx")

# Get a list of all project Excel files in the directory
project_files = glob.glob("G:/Azure-Devops-Board-Userlist-Extract-Automation/SELISESunrise/*.xlsx")

# Loop through each project Excel file
for project_file in project_files:
    # Read the project user names from the Excel file
    project_df = pd.read_excel(project_file, usecols=["User Names"])

    # Merge the two DataFrames on the "User Names" column, keeping all rows from both DataFrames
    merged_df = pd.merge(userlist_df, project_df, how="outer", on="User Names", indicator=True)

    # Filter the merged DataFrame to only include rows that are in the project_df and not in the userlist_df
    result_df = merged_df[merged_df["_merge"] == "right_only"].copy()
    result_df.drop(columns="_merge", inplace=True)

    # Add a column to result_df for concatenating the email addresses of duplicate usernames
    # result_df["Email Concat"] = result_df["Email"]

    # Group the merged DataFrame by "User Names" and concatenate the email addresses using a separator
    merged_df["Email"] = merged_df["Email"].astype(str)
    grouped_df = merged_df.groupby("User Names").agg({"Email": lambda x: "|".join(set(x))})

    # Join the email addresses to the result DataFrame for duplicate usernames
    for user_name, email_concat in grouped_df[grouped_df.duplicated()].iterrows():
        result_df.loc[result_df["User Names"] == user_name, "Email Concat"] = email_concat["Email"]

    # Sort the result DataFrame by "Access Level"
    result_df.sort_values("Access Level", inplace=True)

    # Get the project name from the file path and create a new file name
    project_name = os.path.basename(project_file).split(".")[0]
    result_file_name = f"G:/Azure-Devops-Board-Userlist-Extract-Automation/SELISESunrise-b/{project_name}.xlsx"

    # Write the result DataFrame to a new Excel file
    result_df.to_excel(result_file_name, index=False)
