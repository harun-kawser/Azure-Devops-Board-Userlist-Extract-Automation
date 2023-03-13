import pandas as pd

# Read the Excel file
df = pd.read_excel('G:/Azure-Devops-Board-Userlist-Extract-Automation/merged.xlsx')

# Group the data by project name and access level, and count the number of users
grouped = df.groupby(['Project Name', 'Access Level']).agg({'User Names': 'count'})

# Reset the index to make the grouped data a regular dataframe
grouped = grouped.reset_index()

# Pivot the data to show the number of users for each access level in each project
pivot = pd.pivot_table(grouped, values='User Names', index='Project Name', columns='Access Level', fill_value=0)

# Add a total column that sums up the number of users across all access levels for each project
pivot['Total'] = pivot.sum(axis=1)

# Reindex the columns to include all access levels, and fill NaN values with zeros
access_levels = ['Basic', 'Basic + Test Plans', 'Stakeholder', 'Visual Studio Enterprise subscription', 'Visual Studio Professional subscription']
pivot = pivot.reindex(columns=access_levels + ['Total'], fill_value=0)

# Add a column to the dataframe that combines the project name, access level, and total count
combined_names = []
for project in pivot.index:
    row = pivot.loc[project]
    access_level_counts = [f"{level}={row[level]}" for level in access_levels]
    combined_name = f"{project} - {', '.join(access_level_counts)}, Total={row['Total']}"
    combined_names.append(combined_name)

pivot['Project Name with Access Level and Total'] = combined_names

# Write the result to a new Excel file
pivot.to_excel('output.xlsx')
