import os
import pandas as pd

# Ask user for path to directory containing the Excel files to merge
files_path = input("Enter path to directory containing the Excel files to merge: ")

# Get a list of all Excel files in the directory
excel_files = [f for f in os.listdir(files_path) if f.endswith('.xlsx')]

# Create an empty DataFrame to store the merged data
merged_data = pd.DataFrame()

# Loop through each Excel file
for file_name in excel_files:
    # Read the Excel file into a DataFrame
    df = pd.read_excel(os.path.join(files_path, file_name))
    
    # Get the project name from the file name
    project_name = os.path.splitext(file_name)[0]
    
    # Add the project name as a new column to the DataFrame
    df['Project Name'] = project_name
    
    # Append the DataFrame to the merged_data DataFrame
    merged_data = pd.concat([merged_data, df], ignore_index=True)
    
    # Add a blank row after the data for each project
    spacer = pd.DataFrame([[]])
    merged_data = pd.concat([merged_data, spacer], ignore_index=True)
    
# Write the merged data to a new Excel file
merged_data.to_excel(os.path.join(files_path, 'merged.xlsx'), index=False)
