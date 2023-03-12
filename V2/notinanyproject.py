import pandas as pd
import glob

# Read the userlist Excel file
userlist_df = pd.read_excel("C:/Users/Kawser/Downloads/selisech-user-licenses.xlsx")

# Get a list of all project Excel files in the directory
project_files = glob.glob("G:/Azure-Devops-Board-Userlist-Extract-Automation/projects/*.xlsx")

# Create an empty list to store the merged DataFrames
merged_dfs = []

# Loop through each project Excel file
for project_file in project_files:
    # Read the project user names from the Excel file
    project_df = pd.read_excel(project_file, usecols=["User Names"])

    # Merge the two DataFrames on the "name" column, keeping only the rows that match
    merged_df = pd.merge(userlist_df, project_df, how="inner", left_on="User Names", right_on="User Names")

    # Append the merged_df to the list of merged DataFrames
    merged_dfs.append(merged_df)

# Concatenate all the merged DataFrames into a single DataFrame
merged_df_all = pd.concat(merged_dfs)

# Drop duplicates based on "User Names" column to remove users that have already been included in a project
merged_df_all.drop_duplicates(subset=["User Names"], keep="first", inplace=True)

# Get the list of users that are not in any project
not_in_project_df = userlist_df[~userlist_df["User Names"].isin(merged_df_all["User Names"])]

# Create a new DataFrame with the "User Names", "Email", and "Access Level" columns from the not_in_project_df
not_in_project_result_df = pd.DataFrame({
    "User Names": not_in_project_df["User Names"],
    "Email": not_in_project_df["Email"],
    "Access Level": not_in_project_df["Access Level"]
})
not_in_project_result_df.sort_values("Access Level", inplace=True)

# Write the not_in_project_result_df to a new Excel file
not_in_project_file_name = "G:/Azure-Devops-Board-Userlist-Extract-Automation/not-in-any-project.xlsx"
not_in_project_result_df.to_excel(not_in_project_file_name, index=False)

# Print a message to indicate that the program has finished running
print("All project files have been processed successfully!")
