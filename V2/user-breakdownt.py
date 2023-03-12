import pandas as pd

# Read data from Excel file
df = pd.read_excel("C:/Users/Kawser/Downloads/selisech-user-licenses.xlsx")

# Create four groups based on the access level
basic = df[df['Access Level'] == 'Basic'][['User Names', 'Email']]
basic['Project'] = 'Basic'

basic_tp = df[df['Access Level'] == 'Basic + Test Plans'][['User Names', 'Email']]
basic_tp['Project'] = 'Basic + Test Plans'

stakeholder = df[df['Access Level'] == 'Stakeholder'][['User Names', 'Email']]
stakeholder['Project'] = 'Stakeholder'

enterprise = df.loc[df['Access Level'] == 'Visual Studio Enterprise subscription', ['User Names', 'Email']]
enterprise['Project'] = 'Visual Studio Enterprise subscription'

# Sort the data by access level and User Names
basic = basic.sort_values(['Project', 'User Names'])
basic_tp = basic_tp.sort_values(['Project', 'User Names'])
stakeholder = stakeholder.sort_values(['Project', 'User Names'])
enterprise = enterprise.sort_values(['Project', 'User Names'])

# Save the data for each access level in a separate Excel file
basic.to_excel('Basic.xlsx', index=False)
basic_tp.to_excel('Basic + Test Plans.xlsx', index=False)
stakeholder.to_excel('Stakeholder.xlsx', index=False)
enterprise.to_excel('Visual Studio Enterprise subscription.xlsx', index=False)
