import os
import pandas as pd

# Set the directory path where the Excel files are located
dir_path = 'G:/Azure-Devops-Board-Userlist-Extract-Automation/selisech-a'

# Get a list of all Excel files in the directory
excel_files = [f for f in os.listdir(dir_path) if f.endswith('.xlsx')]

# Initialize a dictionary to store the dataframes for each project
project_dataframes = {}

# Loop through each Excel file and extract the project name from the file name
for excel_file in excel_files:
    project_name = os.path.splitext(excel_file)[0]
    
    # Load the Excel file into a dataframe and add it to the project_dataframes dictionary
    excel_df = pd.read_excel(os.path.join(dir_path, excel_file))
    if project_name in project_dataframes:
        project_dataframes[project_name] = pd.concat([project_dataframes[project_name], excel_df])
    else:
        project_dataframes[project_name] = excel_df

# Calculate the number of users for each project based on their access level
project_users = {}
for project_name, project_df in project_dataframes.items():
    user_counts = project_df.groupby('Access Level')['User Names'].count()
    project_users[project_name] = user_counts

# Write the project dataframes to a new Excel file as tables in a single worksheet
with pd.ExcelWriter('merged_data2.xlsx', engine='xlsxwriter') as writer:
    workbook  = writer.book
    worksheet = workbook.add_worksheet('merged_data')
    bold      = workbook.add_format({'bold': True})
    blue_bold = workbook.add_format({'bold': True, 'font_color': 'blue'})
    
    row = 0
    col = 0
    
    for project_name, project_df in project_dataframes.items():
        # Write the project data as a table
        project_df.to_excel(writer, sheet_name='merged_data', startrow=row, startcol=col, index=False)
        
        # Write the project name in bold and colored in blue
        worksheet.write(row, col, project_name, blue_bold)
        
        # Calculate and write the number of users for each access level
        access_levels = project_df['Access Level'].unique()
        for i, access_level in enumerate(access_levels):
            user_count = project_users[project_name].get(access_level, 0)
            worksheet.write(row+i+1, col+3, access_level)
            worksheet.write(row+i+1, col+4, user_count)
        
        # Calculate and write the total number of users for the project
        total_users = project_users[project_name].sum()
        worksheet.write(row+len(access_levels)+1, col+3, 'Total')
        worksheet.write(row+len(access_levels)+1, col+4, total_users)
        
        # Update the row counter to the next project
        row += len(project_df) + 3
        
        # Apply conditional formatting to the project name cell
        worksheet.conditional_format(row-2, col, row-2, col, {'type': 'no_errors', 'format': blue_bold})