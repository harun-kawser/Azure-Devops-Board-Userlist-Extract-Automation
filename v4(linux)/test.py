from selenium import webdriver

options = webdriver.ChromeOptions()
options.add_argument("user-data-dir=/home/kawser/.config/google-chrome/Profile 1")

driver = webdriver.Chrome(options=options)

# Now you can use the driver object to navigate to web pages, click buttons, and so on
driver.get("https://www.google.com")
