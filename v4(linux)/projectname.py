import pandas as pd
import requests

# Your personal access token and organization URL
token = 'zwp2pnwc6fxnbchunftsg63wywrpjlubpkgrnv552dy7ai3yuyoq'
url = 'https://dev.azure.com/selisech/_apis/projects?api-version=6.0'

# Make the API request and get the JSON response
response = requests.get(url, auth=('', token))
response.raise_for_status()
projects = response.json()

# Extract the relevant data from the JSON response
project_names = [project['name'] for project in projects['value']]

# Create a pandas DataFrame from the data
df = pd.DataFrame(project_names, columns=['name'])

# Save the DataFrame to an Excel file
output_file = 'project_names.xlsx'
df.to_excel(output_file, index=False)
print(f'Saved {len(df)} project names to {output_file}')
