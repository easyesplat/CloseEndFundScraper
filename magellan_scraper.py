from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl
import constants

# Create a new instance of the Chrome webdriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Navigate to the MGF page on the Magellan Group website
driver.get('https://www.magellangroup.com.au/funds/magellan-global-fund-closed-class-asx-mgf/')

# Find the NAV value element by class name
nav_element = driver.find_element(By.CLASS_NAME, 'nIopvPriceSol-MGF')

# Extract the text content of the element
nav_value = nav_element.text

# Print the NAV value to the console
print(f"MGF NAV: {nav_value.split()[1]}")

# Close the browser window
driver.close()

# Uploading data to Excel
workbook = openpyxl.load_workbook(constants.MGF_DATA)
worksheet = workbook['MGF Trade Calc']
worksheet.cell(row = 13, column = 6).value = nav_value.split()[1]
workbook.save(constants.MGF_DATA)
