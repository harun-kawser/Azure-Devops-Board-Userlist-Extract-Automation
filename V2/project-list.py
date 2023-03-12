import os
import pandas as pd

# set path to folder containing project excel files
path = "G:/Azure-Devops-Board-Userlist-Extract-Automation/level-wise"

# get list of project excel file names
project_files = [f for f in os.listdir(path) if f.endswith('.xlsx')]

# create empty dataframe to store combined project access data
combined_access_data = pd.DataFrame(columns=['User Names', 'Email', 'Access Level', 'Project Name'])

# loop through each project file and extract access data
for project_file in project_files:
    # extract project name from file name
    project_name = os.path.splitext(project_file)[0]
    
    # load project access data from excel file into pandas dataframe
    project_access_data = pd.read_excel(os.path.join(path, project_file))
    
    # add project name column to project access data
    project_access_data['Project Name'] = project_name
    
    # append project access data to combined access data
    combined_access_data = pd.concat([combined_access_data, project_access_data[['User Names', 'Email', 'Access Level', 'Project Name']]], ignore_index=True)

# load user access data from excel file into pandas dataframe
user_access_data = pd.read_excel('C:/Users/Kawser/Downloads/selisech-user-licenses.xlsx')

# merge user access data with combined project access data using 'User Name' and 'Email' columns
final_access_data = pd.merge(user_access_data, combined_access_data, on=['User Names', 'Email'], how='left')

# convert entire dataframe to string
final_access_data = final_access_data.astype(str)

# join project names for each user and email
final_access_data['Project Name'] = final_access_data.groupby(['User Names', 'Email'])['Project Name'].transform(lambda x: ', '.join(x))

# drop duplicate rows
final_access_data.drop_duplicates(subset=['User Names', 'Email'], keep='first', inplace=True)

final_access_data = final_access_data.sort_values(by=['Access Level_x'], ascending=True)

# save merged access data to excel file
final_access_data.to_excel('file2.xlsx', index=False)
