import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from openpyxl import load_workbook
import pandas as pd

# Create a ChromeOptions object to specify the location of your user profile
chrome_options = Options()
chrome_options.add_argument("user-data-dir=C:/Users/Kawser/AppData/Local/Google/Chrome/User Data")

# Create a ChromeDriverService object to specify the location of your chromedriver executable
service = Service("H:\chromedriver.exe")

# Attach to the current session
driver = webdriver.Chrome(service=service, options=chrome_options)
driver.execute_cdp_cmd("Browser.grantPermissions", {
  "origin": "https://dev.azure.com",
  "permissions": ["clipboardReadWrite"]
})

# Read the project names from the Excel file
input_file = input("Enter the name of the input file: ")

# Read project names from an Excel file
projects_df = pd.read_excel(input_file, header=None, names=["Project Names"])

# Loop through the project names
for project_name in projects_df["Project Names"]:
    print(f"Processing project {project_name}")
    
    # Navigate to the project's permission settings page
    url = f"https://dev.azure.com/selisech/{project_name}/_settings/permissions"
    driver.get(url)

    # Wait for the page to load
    time.sleep(3)

    # Click the Users tab
    users_tab = driver.find_elements(By.CLASS_NAME, "bolt-tab-text")[1]
    ActionChains(driver).move_to_element(users_tab).click().perform()

    # Wait for the users list to load
    time.sleep(3)

    # Get the initial list of user names
    user_names = []
    user_name_elements = driver.find_elements(By.XPATH, "//span[@class='fontWeightSemiBold bolt-table-two-line-cell-item text-ellipsis']")
    for element in user_name_elements:
        user_names.append(element.text)

    # Create an empty DataFrame
    df = pd.DataFrame(columns=["User Names"])

    # Create an empty set to keep track of unique user names
    seen_user_names = set()

    # Scroll down the page to load more users
    while True:
        # Press the PAGE_DOWN key to scroll down
        action = ActionChains(driver)
        action.send_keys(Keys.PAGE_DOWN * 1).perform()  # Scroll by 200 pixels
        time.sleep(5)
        user_name_elements = driver.find_elements(By.XPATH, "//span[@class='fontWeightSemiBold bolt-table-two-line-cell-item text-ellipsis']")
        new_user_names = [element.text for element in user_name_elements if element.text not in seen_user_names]
        if not new_user_names:
            break
        else:
            # Add new user names to the set of seen names
            seen_user_names.update(new_user_names)
            # Add new user names to the DataFrame
            df = pd.concat([df, pd.DataFrame(new_user_names, columns=["User Names"])])

    # Save the DataFrame to an Excel file with the project name
    # Save the DataFrame to an Excel file
    output_file_name = f"{project_name}.xlsx"
    df.to_excel(output_file_name, index=False)
    print(f"{output_file_name} has been saved.")

# Close the browser
driver.quit()
